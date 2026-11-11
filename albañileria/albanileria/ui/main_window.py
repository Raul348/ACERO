"""Ventana principal de ALBAÑILERÍA E.070."""
from __future__ import annotations

import datetime as _dt
import os
import traceback
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QKeySequence
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog, QFormLayout,
                               QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
                               QMessageBox, QPlainTextEdit, QPushButton, QSpinBox, QSplitter, QTableWidget,
                               QTabWidget, QTextBrowser, QToolBar, QVBoxLayout, QWidget)

from .. import __version__
from ..e030 import norma
from ..e070 import verificar_edificio
from ..e070.materiales import unidades_tabla9
from ..ejemplos import edificio_ejemplo
from ..io_csv import exportar_muros, exportar_tabla, importar_muros
from ..models import (KN_M3, KPA, Cargas, DatosMemoria, Edificio, Material, OpcionesAnalisis, ResultadoEdificio,
                      kgf_cm2, mpa, texto_valor)
from ..reports.word import generar_memoria
from . import theme
from .widgets import MuroView, PlantaView, TablaMuros, TablaNoPortantes, TablaPasos, item

SALIDA = Path(__file__).resolve().parents[2] / "tests" / "salida"


def _spin(minimo, maximo, valor, dec=2, paso=0.1, sufijo="") -> QDoubleSpinBox:
    s = QDoubleSpinBox()
    s.setRange(minimo, maximo)
    s.setDecimals(dec)
    s.setSingleStep(paso)
    s.setValue(valor)
    if sufijo:
        s.setSuffix(" " + sufijo)
    s.setKeyboardTracking(False)
    return s


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"ALBAÑILERÍA E.070 — Verificación y diseño de albañilería confinada  v{__version__}")
        self.resize(1500, 900)
        self.edificio = edificio_ejemplo()
        self.resultado: ResultadoEdificio | None = None
        self.ruta_proyecto: Path | None = None
        self._construir()
        self.cargar_edificio(self.edificio)
        self.statusBar().showMessage("Listo. Se cargó el edificio de ejemplo; presione Verificar (F5).")

    # ================================================================ construcción
    def _construir(self):
        tb = QToolBar("Principal")
        tb.setMovable(False)
        self.addToolBar(tb)
        for txt, fn, atajo in (("Nuevo", self.nuevo, QKeySequence.New), ("Abrir proyecto", self.abrir, QKeySequence.Open),
                               ("Guardar proyecto", self.guardar, QKeySequence.Save), ("Ejemplo", self.cargar_ejemplo, None)):
            a = QAction(txt, self)
            if atajo:
                a.setShortcut(atajo)
            a.triggered.connect(fn)
            tb.addAction(a)
        tb.addSeparator()
        b = QPushButton("▶ Verificar (F5)")
        b.setObjectName("primario")
        b.clicked.connect(self.verificar)
        tb.addWidget(b)
        a = QAction("Verificar", self)
        a.setShortcut(Qt.Key_F5)
        a.triggered.connect(self.verificar)
        self.addAction(a)
        b2 = QPushButton("📄 Memoria Word")
        b2.setObjectName("primario")
        b2.clicked.connect(self.memoria_word)
        tb.addWidget(b2)
        b3 = QPushButton("Exportar resultados CSV")
        b3.clicked.connect(self.exportar_resultados)
        tb.addWidget(b3)

        split = QSplitter(Qt.Horizontal)
        self.tabs = QTabWidget()
        self.tabs.addTab(self._tab_edificio(), "1  Edificio")
        self.tabs.addTab(self._tab_muros(), "2  Muros")
        self.tabs.addTab(self._tab_resultados(), "3  Resultados")
        self.tabs.addTab(self._tab_columnas(), "4  Columnas y soleras")
        self.tabs.addTab(self._tab_no_portantes(), "5  Muros no portantes")
        from .etabs_window import PanelEtabs
        self.panel_etabs = PanelEtabs(self)
        self.tabs.addTab(self.panel_etabs, "6  ETABS")
        self.tabs.addTab(self._tab_memoria(), "7  Datos de la memoria")
        split.addWidget(self.tabs)

        der = QWidget()
        lv = QVBoxLayout(der)
        lv.setContentsMargins(4, 4, 4, 4)
        self.lbl_estado = QLabel("Sin verificar")
        self.lbl_estado.setAlignment(Qt.AlignCenter)
        self.lbl_estado.setStyleSheet(f"font-size: 16pt; font-weight: bold; padding: 10px; background: {theme.PANEL};")
        lv.addWidget(self.lbl_estado)
        self.panel = QTextBrowser()
        self.panel.setOpenExternalLinks(False)
        lv.addWidget(self.panel)
        split.addWidget(der)
        split.setSizes([1000, 480])
        self.setCentralWidget(split)

    # ---------------------------------------------------------------- pestaña 1
    def _tab_edificio(self) -> QWidget:
        w = QWidget()
        g = QGridLayout(w)
        # sitio y sistema
        gb = QGroupBox("Parámetros sísmicos (E.030 2018)")
        f = QFormLayout(gb)
        self.cb_zona = QComboBox()
        self.cb_zona.addItems(["4", "3", "2", "1"])
        self.cb_cat = QComboBox()
        self.cb_cat.addItems(["C", "B", "A2", "A1"])
        self.cb_suelo = QComboBox()
        self.cb_suelo.addItems(["S0", "S1", "S2", "S3"])
        self.sp_tp = _spin(0, 3, 0, 2, 0.05, "s")
        self.sp_tl = _spin(0, 5, 0, 2, 0.1, "s")
        self.sp_tp.setSpecialValueText("automático")
        self.sp_tl.setSpecialValueText("automático")
        self.lbl_zus = QLabel()
        f.addRow("Zona sísmica", self.cb_zona)
        f.addRow("Categoría (uso)", self.cb_cat)
        f.addRow("Tipo de suelo", self.cb_suelo)
        f.addRow("Tp", self.sp_tp)
        f.addRow("Tl", self.sp_tl)
        f.addRow("Z / U / S", self.lbl_zus)
        self.chk_irreg = QCheckBox("Estructura irregular")
        self.sp_ia = _spin(0.4, 1.0, 1.0, 2, 0.05)
        self.sp_ip = _spin(0.4, 1.0, 1.0, 2, 0.05)
        f.addRow(self.chk_irreg)
        f.addRow("Ia (altura)", self.sp_ia)
        f.addRow("Ip (planta)", self.sp_ip)
        f.addRow("R0", QLabel("3 (albañilería confinada, fijo)"))
        self.chk_irreg.toggled.connect(lambda v: (self.sp_ia.setEnabled(v), self.sp_ip.setEnabled(v)))
        for c in (self.cb_zona, self.cb_cat, self.cb_suelo):
            c.currentTextChanged.connect(self._actualizar_zus)
        g.addWidget(gb, 0, 0)

        gb = QGroupBox("Geometría")
        f = QFormLayout(gb)
        self.sp_N = QSpinBox()
        self.sp_N.setRange(1, 5)
        self.sp_h = _spin(2.0, 4.0, 2.6, 2, 0.05, "m")
        self.sp_ap = _spin(10, 5000, 100, 2, 1, "m²")
        self.sp_lx = _spin(0, 200, 0, 2, 0.5, "m")
        self.sp_ly = _spin(0, 200, 0, 2, 0.5, "m")
        self.sp_lx.setSpecialValueText("—")
        self.sp_ly.setSpecialValueText("—")
        f.addRow("Número de pisos N", self.sp_N)
        f.addRow("Altura de entrepiso (piso a piso)", self.sp_h)
        f.addRow("Área en planta Ap", self.sp_ap)
        f.addRow("Lx (planta, para torsión)", self.sp_lx)
        f.addRow("Ly (planta, para torsión)", self.sp_ly)
        g.addWidget(gb, 1, 0)

        gb = QGroupBox("Materiales (E.070 Tabla 9) — kgf/cm²")
        f = QFormLayout(gb)
        self.cb_unidad = QComboBox()
        self.cb_unidad.addItems([u["nombre"] for u in unidades_tabla9()])
        self.cb_unidad.currentTextChanged.connect(self._unidad_cambiada)
        self.cb_familia = QComboBox()
        self.cb_familia.addItems(["arcilla", "silico", "concreto"])
        self.sp_fb = _spin(10, 500, 145, 1, 1)
        self.sp_fm = _spin(10, 300, 65, 1, 1)
        self.sp_vm = _spin(1, 30, 8.1, 2, 0.1)
        self.sp_em = _spin(0, 200000, 0, 0, 100, "MPa")
        self.sp_gm = _spin(0, 80000, 0, 0, 100, "MPa")
        self.sp_em.setSpecialValueText("automático")
        self.sp_gm.setSpecialValueText("automático (0.4·Em)")
        self.cb_mortero = QComboBox()
        self.cb_mortero.addItems(["P1", "P2", "NP"])
        self.sp_fc = _spin(140, 420, 175, 0, 5)
        self.sp_fy = _spin(2800, 5000, 4200, 0, 100)
        self.sp_ft = _spin(0.5, 5, 1.5, 2, 0.1)
        self.sp_gamma = _spin(10, 25, 18, 1, 0.5, "kN/m³")
        self.lbl_mpa = QLabel()
        for et, wdg in (("Tipo de unidad", self.cb_unidad), ("Familia", self.cb_familia), ("f'b", self.sp_fb),
                        ("f'm", self.sp_fm), ("v'm", self.sp_vm), ("Em", self.sp_em), ("Gm", self.sp_gm),
                        ("Mortero", self.cb_mortero), ("f'c confinamiento", self.sp_fc), ("fy", self.sp_fy),
                        ("f't (tracción por flexión)", self.sp_ft), ("γ albañilería", self.sp_gamma),
                        ("Equivalente", self.lbl_mpa)):
            f.addRow(et, wdg)
        for s in (self.sp_fm, self.sp_vm, self.sp_fc, self.sp_fy):
            s.valueChanged.connect(self._actualizar_mpa)
        g.addWidget(gb, 0, 1, 2, 1)

        gb = QGroupBox("Cargas (kN/m²) y opciones de análisis")
        f = QFormLayout(gb)
        self.sp_losa = _spin(0, 20, 3.0, 2, 0.1, "kN/m²")
        self.sp_acab = _spin(0, 10, 1.0, 2, 0.1, "kN/m²")
        self.sp_tabiq = _spin(0, 10, 1.0, 2, 0.1, "kN/m²")
        self.sp_sc = _spin(0, 10, 2.0, 2, 0.1, "kN/m²")
        self.sp_sca = _spin(0, 10, 1.0, 2, 0.1, "kN/m²")
        self.sp_hl = _spin(10, 50, 20, 0, 5, "cm")
        self.cb_momento = QComboBox()
        self.cb_momento.addItems(["Ve·h", "voladizo", "factor"])
        self.sp_brazo = _spin(0.5, 5, 1.0, 2, 0.1, "× h")
        self.cb_rigidez = QComboBox()
        self.cb_rigidez.addItems(["voladizo", "empotrado"])
        self.chk_torsion = QCheckBox("Considerar torsión (exc. real + accidental 0.05·B)")
        self.sp_dmin = _spin(15, 60, 25, 0, 5, "cm")
        for et, wdg in (("Peso de losa", self.sp_losa), ("Acabados", self.sp_acab), ("Tabiquería móvil", self.sp_tabiq),
                        ("Sobrecarga", self.sp_sc), ("Sobrecarga azotea", self.sp_sca), ("Peralte de losa (solera)", self.sp_hl),
                        ("Momento Me (modo manual)", self.cb_momento), ("Brazo (modo factor)", self.sp_brazo),
                        ("Rigidez de muros", self.cb_rigidez), ("", self.chk_torsion),
                        ("Peralte mínimo de columnas", self.sp_dmin)):
            f.addRow(et, wdg)
        g.addWidget(gb, 0, 2, 2, 1)
        g.setRowStretch(2, 1)
        return w

    # ---------------------------------------------------------------- pestaña 2
    def _tab_muros(self) -> QWidget:
        w = QWidget()
        lv = QVBoxLayout(w)
        hb = QHBoxLayout()
        self.tabla_muros = TablaMuros()
        for txt, fn in (("Agregar", self.tabla_muros.agregar), ("Quitar", self.tabla_muros.quitar),
                        ("Duplicar", self.tabla_muros.duplicar), ("Copiar piso 1 a todos los pisos", self.copiar_piso1),
                        ("Importar CSV", self.importar_csv), ("Exportar CSV", self.exportar_csv)):
            b = QPushButton(txt)
            b.clicked.connect(fn)
            hb.addWidget(b)
        hb.addStretch()
        hb.addWidget(QLabel("Piso a dibujar:"))
        self.cb_piso_planta = QComboBox()
        self.cb_piso_planta.currentIndexChanged.connect(self._dibujar)
        hb.addWidget(self.cb_piso_planta)
        lv.addLayout(hb)
        lv.addWidget(QLabel("Pg = D + 25 % L (vacío = por área tributaria); Pm = D + L; Lm vacío = automático; "
                            "Grupo = mismo muro en todos los pisos (vacío = nombre)."))
        sp = QSplitter(Qt.Vertical)
        sp.addWidget(self.tabla_muros)
        abajo = QSplitter(Qt.Horizontal)
        self.planta = PlantaView()
        self.vista_muro = MuroView()
        abajo.addWidget(self.planta)
        abajo.addWidget(self.vista_muro)
        sp.addWidget(abajo)
        sp.setSizes([420, 320])
        lv.addWidget(sp)
        self.tabla_muros.itemSelectionChanged.connect(self._muro_seleccionado)
        self.tabla_muros.itemChanged.connect(lambda *_: self._dibujar())
        self.planta.muro_clicado.connect(self._seleccionar_muro_por_nombre)
        return w

    # ---------------------------------------------------------------- pestaña 3
    def _tab_resultados(self) -> QWidget:
        w = QWidget()
        lv = QVBoxLayout(w)
        sp = QSplitter(Qt.Vertical)
        self.tabla_res = QTableWidget(0, 17)
        self.tabla_res.setHorizontalHeaderLabels(
            ["Muro", "Piso", "Dir", "σm MPa (kgf/cm²)", "Ve kN", "Me kN·m", "α", "Vm kN", "0.55Vm kN", "DCR fis.",
             "Vu kN", "Mu kN·m", "Diseño", "DCR verif.", "DCR máx", "Gobernante", "Estado"])
        self._config_tabla(self.tabla_res)
        self.tabla_res.itemSelectionChanged.connect(self._resultado_seleccionado)
        sp.addWidget(self.tabla_res)
        self.tabla_pisos = QTableWidget(0, 10)
        self.tabla_pisos.setHorizontalHeaderLabels(
            ["Piso", "Dir", "Densidad exist.", "Densidad req.", "DCR dens.", "ΣVm kN", "VE kN", "ΣVm/VE", "Deriva",
             "Muros agrietados / no cumplen"])
        self._config_tabla(self.tabla_pisos)
        self.tabla_pisos.itemSelectionChanged.connect(self._piso_seleccionado)
        sp.addWidget(self.tabla_pisos)
        abajo = QWidget()
        la = QVBoxLayout(abajo)
        la.setContentsMargins(0, 0, 0, 0)
        hb = QHBoxLayout()
        hb.addWidget(QLabel("Verificación:"))
        self.cb_check = QComboBox()
        self.cb_check.currentIndexChanged.connect(self._mostrar_check)
        hb.addWidget(self.cb_check, 1)
        self.lbl_check = QLabel()
        hb.addWidget(self.lbl_check)
        la.addLayout(hb)
        self.tabla_pasos = TablaPasos()
        la.addWidget(self.tabla_pasos)
        sp.addWidget(abajo)
        sp.setSizes([380, 170, 300])
        lv.addWidget(sp)
        self._checks_actuales = []
        return w

    def _config_tabla(self, t: QTableWidget):
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setSelectionMode(QAbstractItemView.SingleSelection)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        t.verticalHeader().setDefaultSectionSize(21)
        t.verticalHeader().setVisible(False)

    # ---------------------------------------------------------------- pestaña 4
    def _tab_columnas(self) -> QWidget:
        w = QWidget()
        lv = QVBoxLayout(w)
        self.tabla_col = QTableWidget(0, 15)
        self.tabla_col.setHorizontalHeaderLabels(
            ["Muro", "Piso", "Diseño", "Pc kN", "Vc kN", "T kN", "C kN", "Acf cm²", "An cm²", "Columna extrema",
             "Refuerzo vertical", "Estribos", "Col. interior", "Ts kN", "Solera (As, barras)"])
        self._config_tabla(self.tabla_col)
        self.tabla_col.itemSelectionChanged.connect(self._columna_seleccionada)
        sp = QSplitter(Qt.Vertical)
        sp.addWidget(self.tabla_col)
        self.tabla_pasos_col = TablaPasos()
        sp.addWidget(self.tabla_pasos_col)
        lv.addWidget(sp)
        return w

    # ---------------------------------------------------------------- pestaña 5
    def _tab_no_portantes(self) -> QWidget:
        w = QWidget()
        lv = QVBoxLayout(w)
        hb = QHBoxLayout()
        self.tabla_np = TablaNoPortantes()
        for txt, fn in (("Agregar", self.tabla_np.agregar), ("Quitar", self.tabla_np.quitar),
                        ("Duplicar", self.tabla_np.duplicar)):
            b = QPushButton(txt)
            b.clicked.connect(fn)
            hb.addWidget(b)
        hb.addStretch()
        lv.addLayout(hb)
        lv.addWidget(QLabel("Tipo: tabique | parapeto | cerco | voladizo.  Caso (Tabla 12): 1 = 4 bordes, 2 = 3 bordes "
                            "(a = borde libre), 3 = sólo bordes horizontales, 4 = voladizo.  C1 vacío = por tipo."))
        sp = QSplitter(Qt.Vertical)
        sp.addWidget(self.tabla_np)
        self.tabla_np_res = QTableWidget(0, 11)
        self.tabla_np_res.setHorizontalHeaderLabels(
            ["Elemento", "Tipo", "C1", "w kN/m²", "b/a", "m", "Ms kN·m/m", "fm MPa", "f't MPa", "Estado", "Arriostres"])
        self._config_tabla(self.tabla_np_res)
        self.tabla_np_res.itemSelectionChanged.connect(self._np_seleccionado)
        sp.addWidget(self.tabla_np_res)
        self.tabla_pasos_np = TablaPasos()
        sp.addWidget(self.tabla_pasos_np)
        lv.addWidget(sp)
        return w

    # ---------------------------------------------------------------- pestaña 7
    def _tab_memoria(self) -> QWidget:
        w = QWidget()
        f = QFormLayout(w)
        self.ed_proyecto = QLineEdit()
        self.ed_ubicacion = QLineEdit()
        self.ed_ingeniero = QLineEdit()
        self.ed_fecha = QLineEdit()
        self.ed_obs = QPlainTextEdit()
        f.addRow("Proyecto", self.ed_proyecto)
        f.addRow("Ubicación", self.ed_ubicacion)
        f.addRow("Ingeniero", self.ed_ingeniero)
        f.addRow("Fecha", self.ed_fecha)
        f.addRow("Observaciones", self.ed_obs)
        return w

    # ================================================================ datos <-> interfaz
    def cargar_edificio(self, ed: Edificio):
        self.edificio = ed
        self.cb_zona.setCurrentText(str(ed.zona))
        self.cb_cat.setCurrentText(ed.categoria)
        self.cb_suelo.setCurrentText(ed.suelo)
        self.sp_tp.setValue(ed.Tp or 0)
        self.sp_tl.setValue(ed.Tl or 0)
        irr = ed.Ia < 1 or ed.Ip < 1
        self.chk_irreg.setChecked(irr)
        self.sp_ia.setValue(ed.Ia)
        self.sp_ip.setValue(ed.Ip)
        self.sp_ia.setEnabled(irr)
        self.sp_ip.setEnabled(irr)
        self.sp_N.setValue(ed.N)
        self.sp_h.setValue(ed.h_entrepiso / 1000)
        self.sp_ap.setValue(ed.Ap / 1e6)
        self.sp_lx.setValue((ed.Lx or 0) / 1000)
        self.sp_ly.setValue((ed.Ly or 0) / 1000)
        m = ed.material
        self.cb_unidad.blockSignals(True)
        self.cb_unidad.setCurrentText(m.tipo_unidad)
        self.cb_unidad.blockSignals(False)
        self.cb_familia.setCurrentText(m.familia)
        self.sp_fb.setValue(kgf_cm2(m.fb))
        self.sp_fm.setValue(kgf_cm2(m.fm))
        self.sp_vm.setValue(kgf_cm2(m.vm))
        self.sp_em.setValue(m.Em or 0)
        self.sp_gm.setValue(m.Gm or 0)
        self.cb_mortero.setCurrentText(m.mortero)
        self.sp_fc.setValue(kgf_cm2(m.fc))
        self.sp_fy.setValue(kgf_cm2(m.fy))
        self.sp_ft.setValue(kgf_cm2(m.ft))
        self.sp_gamma.setValue(m.gamma / KN_M3)
        c = ed.cargas
        self.sp_losa.setValue(c.peso_losa / KPA)
        self.sp_acab.setValue(c.acabados / KPA)
        self.sp_tabiq.setValue(c.tabiqueria / KPA)
        self.sp_sc.setValue(c.sobrecarga / KPA)
        self.sp_sca.setValue(c.sobrecarga_azotea / KPA)
        o = ed.opciones
        self.sp_hl.setValue(o.peralte_losa / 10)
        self.cb_momento.setCurrentText(o.modo_momento)
        self.sp_brazo.setValue(o.factor_brazo)
        self.cb_rigidez.setCurrentText(o.rigidez)
        self.chk_torsion.setChecked(o.torsion)
        self.sp_dmin.setValue(o.dmin_columna / 10)
        self.tabla_muros.cargar(ed.muros)
        self.tabla_np.cargar(ed.no_portantes)
        mm = ed.memoria
        self.ed_proyecto.setText(mm.proyecto)
        self.ed_ubicacion.setText(mm.ubicacion)
        self.ed_ingeniero.setText(mm.ingeniero)
        self.ed_fecha.setText(mm.fecha)
        self.ed_obs.setPlainText(mm.observaciones)
        self._actualizar_zus()
        self._actualizar_mpa()
        self._actualizar_pisos_combo()
        self._dibujar()

    def leer_edificio(self) -> Edificio:
        ed0 = self.edificio
        mat = Material(tipo_unidad=self.cb_unidad.currentText(), familia=self.cb_familia.currentText(),
                       fb=mpa(self.sp_fb.value()), fm=mpa(self.sp_fm.value()), vm=mpa(self.sp_vm.value()),
                       Em=self.sp_em.value() or None, Gm=self.sp_gm.value() or None,
                       mortero=self.cb_mortero.currentText(), fc=mpa(self.sp_fc.value()), fy=mpa(self.sp_fy.value()),
                       ft=mpa(self.sp_ft.value()), gamma=self.sp_gamma.value() * KN_M3)
        car = Cargas(peso_losa=self.sp_losa.value() * KPA, acabados=self.sp_acab.value() * KPA,
                     tabiqueria=self.sp_tabiq.value() * KPA, sobrecarga=self.sp_sc.value() * KPA,
                     sobrecarga_azotea=self.sp_sca.value() * KPA, peralte_losa=self.sp_hl.value() * 10)
        op = OpcionesAnalisis(modo_momento=self.cb_momento.currentText(), factor_brazo=self.sp_brazo.value(),
                              rigidez=self.cb_rigidez.currentText(), torsion=self.chk_torsion.isChecked(),
                              dmin_columna=self.sp_dmin.value() * 10, peralte_losa=self.sp_hl.value() * 10,
                              modo_etabs=ed0.opciones.modo_etabs)
        mem = DatosMemoria(proyecto=self.ed_proyecto.text(), ubicacion=self.ed_ubicacion.text(),
                           ingeniero=self.ed_ingeniero.text(), fecha=self.ed_fecha.text(),
                           observaciones=self.ed_obs.toPlainText())
        irr = self.chk_irreg.isChecked()
        muros = self.tabla_muros.objetos()
        # conservar fuerzas leídas de ETABS (Ve, Me) que no están en la tabla
        previos = {(m.nombre, m.piso): m for m in ed0.muros}
        for m in muros:
            p = previos.get((m.nombre, m.piso))
            if p is not None:
                m.Ve, m.Me, m.material = p.Ve, p.Me, p.material
                m.Pc_ext, m.Pc_int, m.Pt = p.Pc_ext, p.Pc_int, p.Pt
                m.muro_transversal, m.junta_rugosa, m.confinado = p.muro_transversal, p.junta_rugosa, p.confinado
                m.carga_perpendicular, m.aberturas = p.carga_perpendicular, p.aberturas
        ed = Edificio(nombre=ed0.nombre, zona=int(self.cb_zona.currentText()), categoria=self.cb_cat.currentText(),
                      suelo=self.cb_suelo.currentText(), N=self.sp_N.value(), h_entrepiso=self.sp_h.value() * 1000,
                      Ap=self.sp_ap.value() * 1e6, Lx=self.sp_lx.value() * 1000 or None,
                      Ly=self.sp_ly.value() * 1000 or None, Ia=self.sp_ia.value() if irr else 1.0,
                      Ip=self.sp_ip.value() if irr else 1.0, Tp=self.sp_tp.value() or None, Tl=self.sp_tl.value() or None,
                      material=mat, cargas=car, muros=muros, no_portantes=self.tabla_np.objetos(), opciones=op,
                      memoria=mem, alturas=ed0.alturas if ed0.N == self.sp_N.value() else None,
                      pesos_override=ed0.pesos_override, VE_override=ed0.VE_override,
                      derivas=ed0.derivas, origen=ed0.origen)
        return ed

    def _actualizar_zus(self, *_):
        try:
            z = int(self.cb_zona.currentText())
            Z = norma.factor_Z(z)
            U = norma.factor_U(self.cb_cat.currentText())
            S = norma.factor_S(z, self.cb_suelo.currentText())
            tp, tl = norma.periodos_suelo(self.cb_suelo.currentText())
            self.lbl_zus.setText(f"{Z:.2f} / {U:.2f} / {S:.2f}   (Tp = {tp} s, Tl = {tl} s por tabla)")
        except (KeyError, ValueError):
            pass

    def _actualizar_mpa(self, *_):
        self.lbl_mpa.setText(f"f'm = {mpa(self.sp_fm.value()):.3f} MPa · v'm = {mpa(self.sp_vm.value()):.3f} MPa · "
                             f"f'c = {mpa(self.sp_fc.value()):.2f} MPa · fy = {mpa(self.sp_fy.value()):.0f} MPa")

    def _unidad_cambiada(self, nombre: str):
        for u in unidades_tabla9():
            if u["nombre"] == nombre:
                self.cb_familia.setCurrentText(u["familia"])
                self.sp_fb.setValue(u["fb"])
                self.sp_fm.setValue(u["fm"])
                self.sp_vm.setValue(u["vm"])

    def _actualizar_pisos_combo(self):
        self.cb_piso_planta.blockSignals(True)
        actual = self.cb_piso_planta.currentText()
        self.cb_piso_planta.clear()
        self.cb_piso_planta.addItems([str(i) for i in range(1, self.sp_N.value() + 1)])
        if actual:
            self.cb_piso_planta.setCurrentText(actual)
        self.cb_piso_planta.blockSignals(False)

    # ================================================================ dibujo
    def _muros_tabla_seguro(self):
        try:
            return self.tabla_muros.objetos()
        except ValueError:
            return []

    def _dibujar(self, *_):
        piso = int(self.cb_piso_planta.currentText() or 1)
        muros = [m for m in self._muros_tabla_seguro() if m.piso == piso]
        est = {}
        if self.resultado:
            est = {r.nombre: r.estado for r in self.resultado.muros if r.piso == piso}
        sel = ""
        filas = {i.row() for i in self.tabla_muros.selectedIndexes()}
        if filas:
            it = self.tabla_muros.item(min(filas), 0)
            sel = it.text() if it else ""
        self.planta.mostrar(muros, est, sel)

    def _muro_seleccionado(self):
        filas = sorted({i.row() for i in self.tabla_muros.selectedIndexes()})
        if not filas:
            return
        muros = self._muros_tabla_seguro()
        if filas[0] >= len(muros):
            return
        m = muros[filas[0]]
        txt = ""
        if self.resultado:
            r = self.resultado.muro(m.nombre, m.piso)
            if r and r.confinamiento:
                rc = r.confinamiento
                txt = (f"Columnas: {rc.seccion_ext} con {rc.refuerzo_ext}; {rc.estribos}. "
                       f"Solera {rc.seccion_solera}: {rc.refuerzo_solera}.  Estado: {r.estado} (DCR {r.dcr_max:.2f})")
        self.vista_muro.mostrar(m, txt)
        if str(m.piso) != self.cb_piso_planta.currentText():
            self.cb_piso_planta.setCurrentText(str(m.piso))
        self._dibujar()

    def _seleccionar_muro_por_nombre(self, nombre: str):
        piso = self.cb_piso_planta.currentText()
        for r in range(self.tabla_muros.rowCount()):
            if self.tabla_muros.item(r, 0).text() == nombre and self.tabla_muros.item(r, 1).text() == piso:
                self.tabla_muros.selectRow(r)
                return

    # ================================================================ acciones
    def verificar(self):
        try:
            ed = self.leer_edificio()
        except ValueError as e:
            QMessageBox.warning(self, "Datos", str(e))
            return
        try:
            res = verificar_edificio(ed)
        except Exception as e:      # noqa: BLE001
            QMessageBox.critical(self, "Error de cálculo", f"{e}\n\n{traceback.format_exc()}")
            return
        self.edificio, self.resultado = ed, res
        self.mostrar_resultado(res)
        self.statusBar().showMessage(f"Verificación terminada: {res.estado} — DCR máx = {res.dcr_max:.3f} "
                                     f"({len(res.muros)} muros, {len(res.no_portantes)} no portantes)")
        return res

    def mostrar_resultado(self, res: ResultadoEdificio):
        est = res.estado
        self.lbl_estado.setText(f"EDIFICIO: {est}   (DCR máx {res.dcr_max:.3f})")
        self.lbl_estado.setStyleSheet(f"font-size: 15pt; font-weight: bold; padding: 10px; color: white; "
                                      f"background: {theme.VERDE if est == 'CUMPLE' else theme.ROJO};")
        self._llenar_resultados(res)
        self._llenar_columnas(res)
        self._llenar_no_portantes(res)
        self._llenar_panel(res)
        self._dibujar()

    def _fila(self, t: QTableWidget, valores: list, dcr: float | None = None, estado: str | None = None):
        r = t.rowCount()
        t.insertRow(r)
        for j, v in enumerate(valores):
            it = item(v, alinear_derecha=isinstance(v, str) and v[:1].replace("-", "").replace(".", "").isdigit())
            if dcr is not None:
                it.setBackground(theme.color_dcr(dcr))
            t.setItem(r, j, it)
        if estado is not None:
            it = t.item(r, len(valores) - 1)
            it.setBackground(theme.color_estado(estado))
            it.setForeground(QColor("white"))

    def _llenar_resultados(self, res: ResultadoEdificio):
        t = self.tabla_res
        t.setRowCount(0)
        for rm in res.muros:
            v = rm.valores
            if rm.fuera_de_alcance:
                self._fila(t, [rm.nombre, str(rm.piso), rm.direccion] + ["—"] * 13 + [rm.estado], None, rm.estado)
                continue
            g = rm.check_gobernante
            self._fila(t, [rm.nombre, str(rm.piso), rm.direccion, f"{v['sigma']:.3f} ({kgf_cm2(v['sigma']):.2f})",
                           f"{v['Ve'] / 1000:.2f}", f"{v['Me'] / 1e6:.2f}", f"{v['alfa']:.3f}", f"{v['Vm'] / 1000:.2f}",
                           f"{v['Vadm'] / 1000:.2f}", f"{v['dcr_fis']:.3f}", f"{v['Vu'] / 1000:.2f}",
                           f"{v['Mu'] / 1e6:.2f}", v.get("tipo_diseno", ""), f"{rm.dcr_verificacion:.3f}",
                           f"{rm.dcr_max:.3f}", g.nombre if g else "", rm.estado], rm.dcr_verificacion, rm.estado)
        t2 = self.tabla_pisos
        t2.setRowCount(0)
        for pd in res.pisos:
            dmax = max([c.dcr for c in pd.checks if c.aplica] or [0])
            self._fila(t2, [str(pd.piso), pd.direccion, f"{pd.densidad_existente:.4f}", f"{pd.densidad_requerida:.4f}",
                            f"{pd.dcr_densidad:.3f}", f"{pd.suma_Vm / 1000:.1f}", f"{pd.VE / 1000:.1f}",
                            f"{pd.relacion:.2f}" + ("  (elástico)" if pd.elastico else ""),
                            f"{pd.deriva:.5f}" if pd.deriva is not None else "—",
                            (", ".join(pd.muros_agrietados) or "—") + " / " + (", ".join(pd.muros_no_cumplen) or "—")],
                      dmax)

    def _resultado_seleccionado(self):
        filas = {i.row() for i in self.tabla_res.selectedIndexes()}
        if not filas or not self.resultado:
            return
        rm = self.resultado.muros[min(filas)]
        self._poner_checks(rm.checks)

    def _piso_seleccionado(self):
        filas = {i.row() for i in self.tabla_pisos.selectedIndexes()}
        if not filas or not self.resultado:
            return
        pd = self.resultado.pisos[min(filas)]
        s = self.resultado.sismo[pd.direccion]
        from ..models import Check
        sis = Check(f"Análisis sísmico, dirección {pd.direccion}", "E.030 Art. 28", 0, 1, "", pasos=s.pasos, aplica=False)
        self._poner_checks(pd.checks + [sis])

    def _poner_checks(self, checks):
        self._checks_actuales = checks
        self.cb_check.blockSignals(True)
        self.cb_check.clear()
        for c in checks:
            est = "—" if not c.aplica else ("✔" if c.cumple else "✘")
            dcr = "" if not c.aplica else f"  DCR {c.dcr:.3f}"
            self.cb_check.addItem(f"{est} {c.nombre} ({c.referencia}){dcr}")
        self.cb_check.blockSignals(False)
        if checks:
            gob = max(range(len(checks)), key=lambda i: checks[i].dcr)
            self.cb_check.setCurrentIndex(gob)
            self._mostrar_check(gob)

    def _mostrar_check(self, i: int):
        if i < 0 or i >= len(self._checks_actuales):
            return
        c = self._checks_actuales[i]
        self.tabla_pasos.mostrar(c.pasos, c.notas)
        if c.aplica:
            self.lbl_check.setText(f"Demanda {texto_valor(c.demanda, c.unidad)} · Capacidad {texto_valor(c.capacidad, c.unidad)}"
                                   f" · DCR {c.dcr:.3f} → {'CUMPLE' if c.cumple else 'NO CUMPLE'}")
            self.lbl_check.setStyleSheet(f"color: {'#5fd38a' if c.cumple else '#ff7b7b'}; font-weight: bold;")
        else:
            self.lbl_check.setText("Informativo / FUERA DE ALCANCE")
            self.lbl_check.setStyleSheet("color: #aaaaaa;")

    def _llenar_columnas(self, res: ResultadoEdificio):
        t = self.tabla_col
        t.setRowCount(0)
        self._filas_col = []
        for rm in res.muros:
            rc = rm.confinamiento
            if rc is None:
                continue
            self._filas_col.append(rm)
            dmax = max([c.dcr for c in rc.checks] or [0])
            dmax = dmax if dmax > 1 else 0.0
            self._fila(t, [rm.nombre, str(rm.piso), rc.tipo, f"{rc.Pc_ext / 1000:.1f}", f"{rc.Vc_ext / 1000:.1f}",
                           f"{rc.T_ext / 1000:.1f}", f"{rc.C_ext / 1000:.1f}", f"{rc.Acf / 100:.0f}",
                           f"{rc.An_req / 100:.0f}", rc.seccion_ext, f"{rc.refuerzo_ext} ({rc.As_prov_ext / 100:.2f} cm²)",
                           rc.estribos, (f"{rc.seccion_int} {rc.refuerzo_int}" if rc.seccion_int else "—"),
                           f"{rc.Ts / 1000:.1f}", f"{rc.As_solera / 100:.2f} cm² → {rc.refuerzo_solera}"], dmax)

    def _columna_seleccionada(self):
        filas = {i.row() for i in self.tabla_col.selectedIndexes()}
        if not filas:
            return
        rm = self._filas_col[min(filas)]
        pasos = []
        notas = []
        for c in rm.confinamiento.checks:
            pasos += c.pasos
            notas += c.notas
        self.tabla_pasos_col.mostrar(pasos, notas + rm.confinamiento.notas)

    def _llenar_no_portantes(self, res: ResultadoEdificio):
        t = self.tabla_np_res
        t.setRowCount(0)
        for n in res.no_portantes:
            v = n.valores
            arr = f"@ {v['sep_arriostre'] / 1000:.2f} m" if "sep_arriostre" in v else "—"
            self._fila(t, [n.nombre, n.tipo, f"{v['C1']:.2f}", f"{v['w'] * 1000:.3f}",
                           f"{v['ba']:.2f}" if v["caso"] in (1, 2) else "—", f"{v['m']:.4f}", f"{v['Ms'] / 1000:.4f}",
                           f"{v['fm']:.3f}", f"{v['ft']:.3f}", n.estado, arr], n.dcr_max)
            it = t.item(t.rowCount() - 1, 9)
            it.setBackground(theme.color_estado(n.estado))

    def _np_seleccionado(self):
        filas = {i.row() for i in self.tabla_np_res.selectedIndexes()}
        if not filas or not self.resultado:
            return
        n = self.resultado.no_portantes[min(filas)]
        pasos, notas = [], []
        for c in n.checks:
            pasos += c.pasos
            notas += c.notas
        self.tabla_pasos_np.mostrar(pasos, notas)

    def _llenar_panel(self, res: ResultadoEdificio):
        def color(ok):
            return "#5fd38a" if ok else "#ff7b7b"
        h = [f"<h3 style='color:{theme.ACENTO}'>Resumen por piso y dirección</h3>",
             "<table cellspacing=0 cellpadding=3 border=1 style='border-color:#444'>"
             "<tr><th>Piso</th><th>Dir</th><th>Densidad</th><th>ΣVm/VE</th><th>Deriva</th></tr>"]
        for pd in res.pisos:
            h.append(f"<tr><td>{pd.piso}</td><td>{pd.direccion}</td>"
                     f"<td style='color:{color(pd.dcr_densidad <= 1)}'>{pd.densidad_existente:.4f} / "
                     f"{pd.densidad_requerida:.4f}</td>"
                     f"<td style='color:{color(pd.relacion >= 1)}'>{pd.relacion:.2f}{' (elástico)' if pd.elastico else ''}</td>"
                     f"<td>{'' if pd.deriva is None else f'{pd.deriva:.4f}'}</td></tr>")
        h.append("</table>")
        nc = [f"{m.nombre} (P{m.piso})" for m in res.muros if m.estado == "NO CUMPLE"]
        h.append(f"<h3 style='color:{theme.ACENTO}'>Muros que no cumplen</h3>"
                 f"<p style='color:{color(not nc)}'>{', '.join(nc) if nc else 'Ninguno'}</p>")
        for m in res.muros:
            if m.estado == "NO CUMPLE":
                g = m.check_gobernante
                h.append(f"<p>• <b>{m.nombre}</b> piso {m.piso}: {g.nombre} — DCR {g.dcr:.2f}</p>")
        npnc = [n for n in res.no_portantes if n.estado == "NO CUMPLE"]
        if npnc:
            h.append(f"<h3 style='color:{theme.ACENTO}'>No portantes</h3>")
            for n in npnc:
                sep = n.valores.get("sep_arriostre")
                h.append(f"<p>• {n.nombre}: fm = {n.valores['fm']:.3f} MPa &gt; f't"
                         + (f" → arriostres @ {sep / 1000:.2f} m" if sep else "") + "</p>")
        for c in res.checks_globales:
            if not c.cumple:
                h.append(f"<p style='color:#ff7b7b'>• {c.nombre}: {' '.join(c.notas)}</p>")
        if res.advertencias:
            h.append(f"<h3 style='color:{theme.ACENTO}'>Advertencias</h3>")
            for a in list(dict.fromkeys(res.advertencias))[:40]:
                h.append(f"<p style='color:#e0c070'>• {a}</p>")
        h.append(f"<h3 style='color:{theme.ACENTO}'>FUERA DE ALCANCE</h3>")
        for f in res.fuera_de_alcance:
            h.append(f"<p style='color:#aaaaaa'>• {f}</p>")
        self.panel.setHtml("".join(h))

    def closeEvent(self, ev):
        try:
            self.panel_etabs.cerrar_hilo()
        finally:
            super().closeEvent(ev)

    # ---------------------------------------------------------------- archivo
    def nuevo(self):
        ed = Edificio(N=2)
        self.ruta_proyecto = None
        self.resultado = None
        self.cargar_edificio(ed)

    def cargar_ejemplo(self):
        self.ruta_proyecto = None
        self.cargar_edificio(edificio_ejemplo())
        self.verificar()

    def abrir(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Abrir proyecto", "", "Proyecto albañilería (*.json)")
        if not ruta:
            return
        try:
            ed = Edificio.abrir(ruta)
        except Exception as e:      # noqa: BLE001
            QMessageBox.critical(self, "Abrir", f"No se pudo abrir el proyecto:\n{e}")
            return
        self.ruta_proyecto = Path(ruta)
        self.resultado = None
        self.cargar_edificio(ed)
        self.statusBar().showMessage(f"Proyecto abierto: {ruta}")

    def guardar(self):
        try:
            ed = self.leer_edificio()
        except ValueError as e:
            QMessageBox.warning(self, "Datos", str(e))
            return
        ruta = self.ruta_proyecto
        if ruta is None:
            r, _ = QFileDialog.getSaveFileName(self, "Guardar proyecto", "proyecto_albanileria.json", "Proyecto (*.json)")
            if not r:
                return
            ruta = Path(r)
        ed.guardar(ruta)
        self.ruta_proyecto = ruta
        self.edificio = ed
        self.statusBar().showMessage(f"Proyecto guardado: {ruta}")

    def copiar_piso1(self):
        try:
            muros = self.tabla_muros.objetos()
        except ValueError as e:
            QMessageBox.warning(self, "Datos", str(e))
            return
        from dataclasses import replace
        base = [m for m in muros if m.piso == 1]
        nuevos = list(base)
        for p in range(2, self.sp_N.value() + 1):
            nuevos += [replace(m, piso=p, Pg=None, Pm=None) for m in base]
        self.tabla_muros.cargar(nuevos)
        self._actualizar_pisos_combo()
        self._dibujar()

    def importar_csv(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Importar muros", "", "CSV (*.csv)")
        if not ruta:
            return
        try:
            self.tabla_muros.cargar(importar_muros(ruta))
        except Exception as e:      # noqa: BLE001
            QMessageBox.critical(self, "Importar", str(e))
        self._dibujar()

    def exportar_csv(self):
        ruta, _ = QFileDialog.getSaveFileName(self, "Exportar muros", "muros.csv", "CSV (*.csv)")
        if ruta:
            exportar_muros(self.tabla_muros.objetos(), ruta)

    def exportar_resultados(self, ruta: str | None = None):
        if not self.resultado and not self.verificar():
            return
        if not ruta:
            ruta, _ = QFileDialog.getSaveFileName(self, "Exportar resultados", "resultados_muros.csv", "CSV (*.csv)")
            if not ruta:
                return
        t = self.tabla_res
        enc = [t.horizontalHeaderItem(j).text() for j in range(t.columnCount())]
        filas = [[t.item(i, j).text() if t.item(i, j) else "" for j in range(t.columnCount())] for i in range(t.rowCount())]
        exportar_tabla(enc, filas, ruta)
        self.statusBar().showMessage(f"Resultados exportados: {ruta}")
        return ruta

    def memoria_word(self, ruta: str | None = None, preguntar: bool = True):
        res = self.verificar()
        if res is None:
            return None
        if not ruta and preguntar:
            nombre = f"Memoria_{(self.edificio.memoria.proyecto or 'albanileria').replace(' ', '_')[:40]}.docx"
            ruta, _ = QFileDialog.getSaveFileName(self, "Memoria de cálculo", nombre, "Word (*.docx)")
            if not ruta:
                return None
        if not ruta:
            SALIDA.mkdir(parents=True, exist_ok=True)
            ruta = str(SALIDA / f"Memoria_{_dt.datetime.now():%Y%m%d_%H%M%S}.docx")
        try:
            p = generar_memoria(self.edificio, res, ruta)
        except PermissionError:
            QMessageBox.warning(self, "Memoria", "No se pudo escribir el archivo (¿está abierto en Word?).")
            return None
        self.statusBar().showMessage(f"Memoria generada: {p}")
        if preguntar and os.name == "nt":
            try:
                os.startfile(str(p))         # noqa: S606
            except OSError:
                pass
        return p
