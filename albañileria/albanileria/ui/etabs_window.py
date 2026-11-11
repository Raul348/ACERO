"""Pestaña ETABS: conexión COM en un hilo de trabajo (QThread) que nunca toca widgets.

El objeto TrabajadorEtabs vive en su propio QThread y es el único que usa COM (la instancia de ETABS
se crea y se usa en ese mismo hilo). La interfaz le pide trabajos con la señal `pedir` y recibe
resultados por señales que se atienden en slots del hilo principal.
"""
from __future__ import annotations

import datetime as _dt
import traceback
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QFormLayout, QGroupBox,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QPushButton,
                               QSplitter, QTableWidget, QVBoxLayout, QWidget)

from ..etabs.alternativas import Copiloto, aplicar_en_etabs, deshacer_en_etabs
from ..etabs.batch import ConfigEtabs, edificio_desde_snapshot
from ..etabs.clasificar import clasificar_muros
from ..etabs.client import SALIDA_REAL, ClienteEtabs
from ..etabs.model import SnapshotEtabs
from ..etabs.reader import leer_modelo
from .widgets import item


class TrabajadorEtabs(QObject):
    mensaje = Signal(str)
    conectado = Signal(list, list, str)
    snapshot_listo = Signal(object)
    error = Signal(str)
    hecho = Signal(str)

    def __init__(self):
        super().__init__()
        self.cli: ClienteEtabs | None = None
        self.historial: list = []

    @Slot(str, object)
    def ejecutar(self, orden: str, datos: object):
        try:
            getattr(self, "_" + orden)(datos)
        except Exception as e:      # noqa: BLE001
            self.error.emit(f"{e}\n\n{traceback.format_exc()}")

    def _com(self):
        try:
            import comtypes  # type: ignore
            comtypes.CoInitialize()
        except Exception:           # noqa: BLE001
            pass

    def _conectar(self, ruta):
        self._com()
        if self.cli is None:
            self.cli = ClienteEtabs()
            self.mensaje.emit("Iniciando ETABS (instancia oculta propia)…")
            self.cli.iniciar()
        copia = self.cli.abrir(ruta)
        self.mensaje.emit(f"Modelo copiado y abierto: {copia}")
        self.conectado.emit(self.cli.casos(), self.cli.combos(), str(copia))

    def _leer(self, cfg: dict):
        if self.cli is None:
            raise RuntimeError("Primero conecte con ETABS.")
        if cfg.get("crear_combo"):
            cfg["combo_gravedad"] = self.cli.crear_combo_gravedad()
            self.mensaje.emit(f"Combo creado en la copia: {cfg['combo_gravedad']}")
        snap = leer_modelo(self.cli, cfg["casos_sismo"], cfg["combo_gravedad"], cfg.get("combo_servicio", ""),
                           cfg.get("lista_usuario"), avisar=self.mensaje.emit)
        ruta = Path(str(self.cli.ruta_modelo.with_suffix("")) + "_snapshot.json")
        snap.guardar(ruta)
        self.mensaje.emit(f"Snapshot guardado: {ruta}")
        self.snapshot_listo.emit(snap)

    def _aplicar(self, datos):
        alt, ed, snap = datos
        h = aplicar_en_etabs(self.cli, ed, alt, snap)
        self.historial.append(h)
        self.cli.guardar()
        self.hecho.emit(f"Alternativa aplicada en la copia ({len(h)} áreas). Ejecute 'reanalizar'.")

    def _deshacer(self, _):
        if not self.historial:
            self.hecho.emit("No hay cambios que deshacer.")
            return
        deshacer_en_etabs(self.cli, self.historial.pop())
        self.cli.guardar()
        self.hecho.emit("Cambio deshecho en la copia. Ejecute 'reanalizar'.")

    def _reanalizar(self, cfg: dict):
        self.mensaje.emit("Reanalizando la copia…")
        self.cli.analizar()
        self._leer(cfg)

    def _cerrar(self, _):
        if self.cli is not None:
            self.cli.cerrar()
            self.cli = None
        self.hecho.emit("ETABS cerrado.")


class PanelEtabs(QWidget):
    pedir = Signal(str, object)

    def __init__(self, ventana):
        super().__init__()
        self.v = ventana
        self.snap: SnapshotEtabs | None = None
        self.hilo = QThread(self)
        self.trab = TrabajadorEtabs()
        self.trab.moveToThread(self.hilo)
        self.pedir.connect(self.trab.ejecutar, Qt.QueuedConnection)
        self.trab.mensaje.connect(self._log)
        self.trab.conectado.connect(self._al_conectar)
        self.trab.snapshot_listo.connect(self._al_snapshot)
        self.trab.error.connect(self._al_error)
        self.trab.hecho.connect(self._al_hecho)
        self.hilo.start()
        self.copiloto = Copiloto(self._obtener, self._cop_aplicar, self._cop_deshacer, self._cop_reanalizar)
        self._construir()

    # ------------------------------------------------------------ interfaz
    def _construir(self):
        lv = QVBoxLayout(self)
        gb = QGroupBox("Modelo ETABS 22 (se trabaja SIEMPRE sobre una copia en tests/salida/real)")
        f = QFormLayout(gb)
        hb = QHBoxLayout()
        self.ed_ruta = QLineEdit()
        b = QPushButton("…")
        b.clicked.connect(self._elegir_edb)
        hb.addWidget(self.ed_ruta)
        hb.addWidget(b)
        f.addRow("Archivo .EDB", hb)
        hb = QHBoxLayout()
        for txt, fn in (("Conectar", self.conectar), ("Leer modelo", self.leer), ("Abrir snapshot JSON", self.abrir_snapshot),
                        ("Clasificar muros", self.clasificar), ("Verificar edificio", self.verificar),
                        ("Exportar CSV", self.exportar_csv), ("Memoria Word automática", self.memoria),
                        ("Cerrar ETABS", self.cerrar)):
            bt = QPushButton(txt)
            if txt in ("Verificar edificio", "Memoria Word automática"):
                bt.setObjectName("primario")
            bt.clicked.connect(fn)
            hb.addWidget(bt)
        f.addRow(hb)
        hb = QHBoxLayout()
        self.cb_sx, self.cb_sy, self.cb_grav, self.cb_serv = QComboBox(), QComboBox(), QComboBox(), QComboBox()
        for c in (self.cb_sx, self.cb_sy, self.cb_grav, self.cb_serv):
            c.setEditable(True)
            c.setMinimumWidth(140)
        for et, c in (("Sismo X", self.cb_sx), ("Sismo Y", self.cb_sy), ("Gravedad D+0.25L", self.cb_grav),
                      ("Servicio D+L", self.cb_serv)):
            hb.addWidget(QLabel(et))
            hb.addWidget(c)
        f.addRow(hb)
        hb = QHBoxLayout()
        self.chk_moderado = QCheckBox("Los casos de sismo ya son el sismo MODERADO (no multiplicar por 0.5)")
        self.chk_combo = QCheckBox("Crear combo D+0.25L en la copia")
        hb.addWidget(self.chk_moderado)
        hb.addWidget(self.chk_combo)
        f.addRow(hb)
        lv.addWidget(gb)

        sp = QSplitter(Qt.Vertical)
        self.tabla = QTableWidget(0, 9)
        self.tabla.setHorizontalHeaderLabels(["Albañilería", "Pier", "Piso", "Material", "Tipo", "L (m)", "t (cm)", "Dir",
                                              "Motivo"])
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tabla.verticalHeader().setVisible(False)
        sp.addWidget(self.tabla)
        cop = QWidget()
        lc = QVBoxLayout(cop)
        lc.setContentsMargins(0, 0, 0, 0)
        lc.addWidget(QLabel("Copiloto — " + Copiloto.AYUDA))
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        lc.addWidget(self.log)
        self.ed_cmd = QLineEdit()
        self.ed_cmd.setPlaceholderText("Escriba un comando y presione Enter (p.ej. revisar edificio)")
        self.ed_cmd.returnPressed.connect(self._comando)
        lc.addWidget(self.ed_cmd)
        sp.addWidget(cop)
        lv.addWidget(sp)

    def _log(self, txt: str):
        self.log.appendPlainText(f"[{_dt.datetime.now():%H:%M:%S}] {txt}")
        self.v.statusBar().showMessage(txt.splitlines()[0][:150])

    # ------------------------------------------------------------ slots de resultados del hilo
    @Slot(list, list, str)
    def _al_conectar(self, casos: list, combos: list, copia: str):
        todos = [""] + casos + combos
        for c in (self.cb_sx, self.cb_sy, self.cb_grav, self.cb_serv):
            act = c.currentText()
            c.clear()
            c.addItems(todos)
            c.setCurrentText(act)
        for c, claves in ((self.cb_sx, ("SX", "SISX", "EX", "SDX", "DINX")), (self.cb_sy, ("SY", "SISY", "EY", "SDY", "DINY")),
                          (self.cb_grav, ("0.25", "PG", "D+0.25")), (self.cb_serv, ("D+L", "SERV"))):
            if not c.currentText():
                for n in todos:
                    if any(k in n.upper() for k in claves):
                        c.setCurrentText(n)
                        break
        self._log(f"Conectado. {len(casos)} casos y {len(combos)} combinaciones disponibles.")

    @Slot(object)
    def _al_snapshot(self, snap: SnapshotEtabs):
        self.snap = snap
        self._llenar_tabla()
        self._log(f"Modelo leído: {len(snap.pisos)} pisos, {len(snap.muros)} piers, {len(snap.fuerzas)} filas de fuerzas.")
        self.verificar()

    @Slot(str)
    def _al_error(self, txt: str):
        self._log("ERROR: " + txt.splitlines()[0])
        QMessageBox.critical(self, "ETABS", txt[:3000])

    @Slot(str)
    def _al_hecho(self, txt: str):
        self._log(txt)

    # ------------------------------------------------------------ acciones
    def _elegir_edb(self):
        r, _ = QFileDialog.getOpenFileName(self, "Modelo ETABS", "", "ETABS (*.EDB *.edb)")
        if r:
            self.ed_ruta.setText(r)

    def _cfg_dict(self) -> dict:
        return {"casos_sismo": {"X": self.cb_sx.currentText(), "Y": self.cb_sy.currentText()},
                "combo_gravedad": self.cb_grav.currentText(), "combo_servicio": self.cb_serv.currentText(),
                "crear_combo": self.chk_combo.isChecked() or not self.cb_grav.currentText(),
                "lista_usuario": self._lista_usuario()}

    def _lista_usuario(self):
        if self.tabla.rowCount() == 0:
            return None
        out = []
        for r in range(self.tabla.rowCount()):
            w = self.tabla.cellWidget(r, 0)
            if w is not None and w.isChecked():
                out.append(self.tabla.item(r, 1).text())
        return out

    def conectar(self):
        ruta = self.ed_ruta.text().strip()
        if not ruta:
            self._elegir_edb()
            ruta = self.ed_ruta.text().strip()
            if not ruta:
                return
        self._log("Conectando…")
        self.pedir.emit("conectar", ruta)

    def leer(self):
        self._log("Leyendo el modelo…")
        self.pedir.emit("leer", self._cfg_dict())

    def cerrar(self):
        self.pedir.emit("cerrar", None)

    def abrir_snapshot(self):
        r, _ = QFileDialog.getOpenFileName(self, "Snapshot ETABS", str(SALIDA_REAL), "Snapshot (*.json)")
        if not r:
            return
        snap = SnapshotEtabs.abrir(r)
        casos = sorted({f.caso for f in snap.fuerzas})
        self._al_conectar(casos, [], r)
        for c, k in ((self.cb_sx, "X"), (self.cb_sy, "Y")):
            if snap.casos_sismo.get(k):
                c.setCurrentText(snap.casos_sismo[k])
        if snap.combo_gravedad:
            self.cb_grav.setCurrentText(snap.combo_gravedad)
        if snap.combo_servicio:
            self.cb_serv.setCurrentText(snap.combo_servicio)
        self._al_snapshot(snap)

    def _llenar_tabla(self):
        t = self.tabla
        t.setRowCount(0)
        for m in self.snap.muros:
            r = t.rowCount()
            t.insertRow(r)
            chk = QCheckBox()
            chk.setChecked(bool(m.es_albanileria))
            t.setCellWidget(r, 0, chk)
            for j, v in enumerate((m.pier, m.piso, m.material, m.tipo_material, f"{m.longitud / 1000:.2f}",
                                   f"{m.espesor / 10:.1f}", m.direccion, m.motivo), start=1):
                t.setItem(r, j, item(v))

    def clasificar(self):
        if not self.snap:
            return
        lista = self._lista_usuario()
        clasificar_muros(self.snap.muros, lista)
        # la lista marcada manda: lo no marcado no es albañilería
        for m in self.snap.muros:
            if lista is not None and m.pier not in lista:
                m.es_albanileria, m.motivo = False, "desmarcado por el usuario"
        self._llenar_tabla()
        self._log(f"Clasificación: {sum(1 for m in self.snap.muros if m.es_albanileria)} muros de albañilería.")

    def _config(self) -> ConfigEtabs:
        ed = self.v.leer_edificio()
        return ConfigEtabs(casos_sismo={"X": self.cb_sx.currentText(), "Y": self.cb_sy.currentText()},
                           combo_gravedad=self.cb_grav.currentText(), combo_servicio=self.cb_serv.currentText(),
                           factor_moderado=1.0 if self.chk_moderado.isChecked() else 0.5, zona=ed.zona,
                           categoria=ed.categoria, suelo=ed.suelo,
                           Ap=ed.Ap if self.v.edificio.origen == "etabs" else None, material=ed.material,
                           proyecto=ed.memoria.proyecto)

    def verificar(self):
        if not self.snap:
            QMessageBox.information(self, "ETABS", "Primero lea el modelo o abra un snapshot.")
            return None
        try:
            ed, av = edificio_desde_snapshot(self.snap, self._config())
        except Exception as e:      # noqa: BLE001
            self._al_error(f"{e}\n\n{traceback.format_exc()}")
            return None
        ed.memoria = self.v.leer_edificio().memoria
        if not ed.memoria.proyecto:
            ed.memoria.proyecto = ed.nombre
        self.v.cargar_edificio(ed)
        res = self.v.verificar()
        if res is not None:
            res.advertencias = av + res.advertencias
            self.v.mostrar_resultado(res)
            self._log(self.copiloto.ejecutar("revisar edificio"))
        return res

    def exportar_csv(self):
        if self.v.resultado is None and self.verificar() is None:
            return
        SALIDA_REAL.mkdir(parents=True, exist_ok=True)
        ruta = SALIDA_REAL / f"resultados_{_dt.datetime.now():%Y%m%d_%H%M%S}.csv"
        self.v.exportar_resultados(str(ruta))
        self._log(f"CSV exportado: {ruta}")

    def memoria(self):
        if self.snap is not None and self.v.resultado is None:
            self.verificar()
        SALIDA_REAL.mkdir(parents=True, exist_ok=True)
        ruta = SALIDA_REAL / f"Memoria_{_dt.datetime.now():%Y%m%d_%H%M%S}.docx"
        p = self.v.memoria_word(str(ruta), preguntar=False)
        if p:
            self._log(f"Memoria generada: {p}")

    # ------------------------------------------------------------ copiloto
    def _obtener(self):
        return self.v.edificio, self.v.resultado

    def _comando(self):
        txt = self.ed_cmd.text()
        self.ed_cmd.clear()
        self._log("> " + txt)
        try:
            self._log(self.copiloto.ejecutar(txt))
        except Exception as e:      # noqa: BLE001
            self._log(f"Error: {e}")

    def _cop_aplicar(self, alt) -> str:
        if self.trab is None or self.snap is None:
            return "Aplicar requiere un modelo ETABS conectado."
        if alt.tipo != "espesor":
            return "FUERA DE ALCANCE: sólo el cambio de espesor se aplica automáticamente en el modelo copia."
        self.pedir.emit("aplicar", (alt, self.v.edificio, self.snap))
        return f"Aplicando en la copia: {alt.descripcion}…"

    def _cop_deshacer(self) -> str:
        self.pedir.emit("deshacer", None)
        return "Deshaciendo…"

    def _cop_reanalizar(self) -> str:
        self.pedir.emit("reanalizar", self._cfg_dict())
        return "Reanalizando y releyendo el modelo copia…"

    def cerrar_hilo(self):
        self.pedir.emit("cerrar", None)
        self.hilo.quit()
        self.hilo.wait(5000)
