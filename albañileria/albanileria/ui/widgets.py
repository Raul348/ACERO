"""Widgets reutilizables: tablas editables, tabla de pasos, dibujo en planta y esquema del muro."""
from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QHeaderView, QTableWidget, QTableWidgetItem,
                               QWidget)

from ..io_csv import COLUMNAS_MURO, COLUMNAS_NP, a_texto, de_texto
from ..models import ElementoNoPortante, Muro, Step
from . import theme


def item(texto, editable: bool = False, alinear_derecha: bool = False) -> QTableWidgetItem:
    it = QTableWidgetItem(str(texto))
    if not editable:
        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    if alinear_derecha:
        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return it


class TablaEditable(QTableWidget):
    """Tabla genérica editable a partir de una lista de columnas (encabezado, atributo, tipo, factor, opcional)."""

    def __init__(self, columnas, fabrica, parent=None):
        super().__init__(0, len(columnas), parent)
        self.columnas = columnas
        self.fabrica = fabrica
        self.setHorizontalHeaderLabels([c[0] for c in columnas])
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.verticalHeader().setDefaultSectionSize(22)

    def cargar(self, objetos) -> None:
        self.blockSignals(True)
        self.setRowCount(0)
        for o in objetos:
            self._agregar_fila(o)
        self.blockSignals(False)

    def _agregar_fila(self, o, fila: int | None = None) -> None:
        r = self.rowCount() if fila is None else fila
        self.insertRow(r)
        for j, c in enumerate(self.columnas):
            self.setItem(r, j, item(a_texto(o, c), editable=True))

    def objetos(self) -> list:
        res = []
        for r in range(self.rowCount()):
            kw = {}
            for j, c in enumerate(self.columnas):
                it = self.item(r, j)
                try:
                    kw[c[1]] = de_texto(it.text() if it else "", c)
                except ValueError as e:
                    raise ValueError(f"Fila {r + 1}: {e}") from e
            res.append(self.fabrica(**kw))
        return res

    def agregar(self) -> None:
        objs = self.objetos() if self.rowCount() else []
        o = self.fabrica(nombre=f"M{len(objs) + 1}") if not objs else self._copiar(objs[-1], nuevo_nombre=True)
        self._agregar_fila(o)
        self.selectRow(self.rowCount() - 1)

    def quitar(self) -> None:
        filas = sorted({i.row() for i in self.selectedIndexes()}, reverse=True)
        for r in filas:
            self.removeRow(r)

    def duplicar(self) -> None:
        filas = sorted({i.row() for i in self.selectedIndexes()})
        objs = self.objetos()
        for k, r in enumerate(filas):
            self._agregar_fila(self._copiar(objs[r], nuevo_nombre=True))

    def _copiar(self, o, nuevo_nombre=False):
        from dataclasses import replace
        n = replace(o)
        if nuevo_nombre:
            n.nombre = o.nombre + "'"
        return n


class TablaMuros(TablaEditable):
    def __init__(self, parent=None):
        super().__init__(COLUMNAS_MURO, Muro, parent)


class TablaNoPortantes(TablaEditable):
    def __init__(self, parent=None):
        super().__init__(COLUMNAS_NP, ElementoNoPortante, parent)


class TablaPasos(QTableWidget):
    """Tabla inferior con los pasos de un Check: símbolo, fórmula, sustitución, valor, descripción."""

    def __init__(self, parent=None):
        super().__init__(0, 5, parent)
        self.setHorizontalHeaderLabels(["Símbolo", "Fórmula", "Sustitución", "Valor", "Descripción"])
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.horizontalHeader().setStretchLastSection(True)
        self.setAlternatingRowColors(True)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.verticalHeader().setVisible(False)
        for j, w in enumerate((70, 260, 260, 160)):
            self.setColumnWidth(j, w)

    def mostrar(self, pasos: list[Step], notas: list[str] | None = None) -> None:
        self.setRowCount(0)
        for s in pasos:
            r = self.rowCount()
            self.insertRow(r)
            for j, v in enumerate((s.simbolo, s.formula, s.sustitucion, s.valor, s.descripcion)):
                self.setItem(r, j, item(v))
        for n in notas or []:
            r = self.rowCount()
            self.insertRow(r)
            it = item("Nota")
            self.setItem(r, 0, it)
            it2 = item(n)
            it2.setForeground(QColor("#ffb3b3"))
            self.setItem(r, 4, it2)
            self.setSpan(r, 1, 1, 3)


class PlantaView(QWidget):
    """Dibujo esquemático en planta de los muros de un piso (líneas con nombre y dirección)."""
    muro_clicado = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.muros: list[Muro] = []
        self.estados: dict[str, str] = {}
        self.seleccionado = ""
        self.setMinimumSize(320, 260)

    def mostrar(self, muros: list[Muro], estados: dict[str, str] | None = None, seleccionado: str = "") -> None:
        self.muros = muros
        self.estados = estados or {}
        self.seleccionado = seleccionado
        self.update()

    def _extremos(self, m: Muro):
        if m.direccion == "X":
            return (m.x - m.L / 2, m.y), (m.x + m.L / 2, m.y)
        return (m.x, m.y - m.L / 2), (m.x, m.y + m.L / 2)

    def _transform(self):
        pts = [p for m in self.muros for p in self._extremos(m)]
        if not pts:
            return None
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        w, h = max(x1 - x0, 1000), max(y1 - y0, 1000)
        mg = 40
        esc = min((self.width() - 2 * mg) / w, (self.height() - 2 * mg) / h)
        ox = mg + ((self.width() - 2 * mg) - w * esc) / 2
        oy = mg + ((self.height() - 2 * mg) - h * esc) / 2
        return lambda x, y: QPointF(ox + (x - x0) * esc, self.height() - (oy + (y - y0) * esc))

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor(theme.FONDO))
        tr = self._transform()
        if tr is None:
            p.setPen(QColor("#888"))
            p.drawText(self.rect(), Qt.AlignCenter, "Sin muros en este piso")
            return
        p.setFont(QFont("Segoe UI", 8))
        for m in self.muros:
            a, b = self._extremos(m)
            est = self.estados.get(m.nombre, "")
            col = QColor(theme.VERDE if est == "CUMPLE" else theme.ROJO if est == "NO CUMPLE" else "#9aa7b8")
            if not m.portante:
                col = QColor("#666")
            ancho = 6 if m.nombre == self.seleccionado else 4
            p.setPen(QPen(col, ancho, Qt.SolidLine, Qt.FlatCap))
            pa, pb = tr(*a), tr(*b)
            p.drawLine(pa, pb)
            p.setPen(QPen(QColor(theme.ACENTO), 1))
            p.setBrush(QBrush(QColor(theme.ACENTO)))
            for k in range(max(m.Nc, 2)):
                t = k / (max(m.Nc, 2) - 1)
                q = QPointF(pa.x() + (pb.x() - pa.x()) * t, pa.y() + (pb.y() - pa.y()) * t)
                p.drawRect(QRectF(q.x() - 3, q.y() - 3, 6, 6))
            mid = (pa + pb) / 2
            p.setPen(QColor(theme.TEXTO))
            off = QPointF(0, -8) if m.direccion == "X" else QPointF(6, 0)
            p.drawText(mid + off, m.nombre)
        p.setPen(QColor("#888"))
        p.drawText(8, 16, "Planta esquemática — verde: cumple, rojo: no cumple, gris: no portante")
        p.end()

    def mousePressEvent(self, ev):
        tr = self._transform()
        if tr is None:
            return
        pos = ev.position()
        mejor, dist = None, 1e9
        for m in self.muros:
            a, b = self._extremos(m)
            pa, pb = tr(*a), tr(*b)
            mid = (pa + pb) / 2
            d = (mid - pos).manhattanLength()
            if d < dist:
                mejor, dist = m, d
        if mejor is not None and dist < 80:
            self.muro_clicado.emit(mejor.nombre)


class MuroView(QWidget):
    """Esquema en elevación del muro seleccionado con sus columnas y solera."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.muro: Muro | None = None
        self.texto = ""
        self.setMinimumSize(320, 200)

    def mostrar(self, muro: Muro | None, texto: str = "") -> None:
        self.muro = muro
        self.texto = texto
        self.update()

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor(theme.FONDO))
        m = self.muro
        if m is None:
            p.setPen(QColor("#888"))
            p.drawText(self.rect(), Qt.AlignCenter, "Seleccione un muro")
            return
        mg = 30
        hs = 200.0
        bc = 250.0
        W, H = m.L + bc, m.h + hs
        esc = min((self.width() - 2 * mg) / W, (self.height() - 2 * mg - 70) / H)
        ox = (self.width() - W * esc) / 2
        oy = mg + 20
        # albañilería
        p.setPen(QPen(QColor("#a0522d"), 1))
        p.setBrush(QBrush(QColor("#8b4a2b"), Qt.Dense5Pattern))
        p.drawRect(QRectF(ox, oy + hs * esc, W * esc, m.h * esc))
        # solera
        p.setBrush(QBrush(QColor("#7f8c8d")))
        p.setPen(QPen(QColor("#bdc3c7"), 1))
        p.drawRect(QRectF(ox, oy, W * esc, hs * esc))
        # columnas
        n = max(m.Nc, 2)
        for k in range(n):
            x = ox + (m.L * k / (n - 1)) * esc
            p.drawRect(QRectF(x, oy, bc * esc, H * esc))
        p.setPen(QColor(theme.TEXTO))
        p.setFont(QFont("Segoe UI", 8))
        p.drawText(QPointF(ox, oy - 6), f"{m.nombre} — piso {m.piso}, dir. {m.direccion}: L = {m.L / 1000:.2f} m, "
                                        f"t = {m.t / 10:.0f} cm, h = {m.h / 1000:.2f} m, Nc = {m.Nc}")
        if self.texto:
            p.drawText(QRectF(mg, oy + H * esc + 4, self.width() - 2 * mg, 60), Qt.TextWordWrap, self.texto)
        p.end()


class ComboTexto(QComboBox):
    def __init__(self, opciones, actual=None, parent=None):
        super().__init__(parent)
        self.addItems([str(o) for o in opciones])
        if actual is not None:
            self.setCurrentText(str(actual))
