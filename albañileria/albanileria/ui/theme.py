"""Tema oscuro tipo KALLPA (Fusion + paleta oscura)."""
from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

FONDO = "#1e1e1e"
FONDO2 = "#252526"
PANEL = "#2d2d30"
TEXTO = "#e0e0e0"
ACENTO = "#3a7bd5"
VERDE = "#1b7f3b"
ROJO = "#b01c1c"
AMBAR = "#8a6d1a"
GRIS = "#3c3c3c"

QSS = f"""
QWidget {{ color: {TEXTO}; font-size: 9pt; }}
QMainWindow, QDialog {{ background: {FONDO}; }}
QTabWidget::pane {{ border: 1px solid {GRIS}; background: {FONDO2}; }}
QTabBar::tab {{ background: {PANEL}; padding: 6px 14px; border: 1px solid {GRIS}; border-bottom: none; }}
QTabBar::tab:selected {{ background: {ACENTO}; color: white; }}
QGroupBox {{ border: 1px solid {GRIS}; border-radius: 4px; margin-top: 10px; padding-top: 6px; background: {PANEL}; }}
QGroupBox::title {{ subcontrol-origin: margin; left: 8px; padding: 0 4px; color: {ACENTO}; font-weight: bold; }}
QPushButton {{ background: {PANEL}; border: 1px solid {GRIS}; border-radius: 3px; padding: 5px 12px; }}
QPushButton:hover {{ border-color: {ACENTO}; }}
QPushButton#primario {{ background: {ACENTO}; color: white; font-weight: bold; border: none; }}
QPushButton#primario:hover {{ background: #4a8be5; }}
QTableView, QTableWidget {{ background: {FONDO2}; alternate-background-color: #2a2a2c; gridline-color: {GRIS};
    selection-background-color: {ACENTO}; }}
QHeaderView::section {{ background: {PANEL}; padding: 4px; border: 1px solid {GRIS}; font-weight: bold; }}
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QPlainTextEdit, QTextEdit, QTextBrowser {{
    background: {FONDO2}; border: 1px solid {GRIS}; border-radius: 2px; padding: 2px; }}
QStatusBar {{ background: {PANEL}; }}
QToolBar {{ background: {PANEL}; border: none; spacing: 4px; }}
QSplitter::handle {{ background: {GRIS}; }}
"""


def aplicar_tema(app: QApplication) -> None:
    app.setStyle("Fusion")
    p = QPalette()
    p.setColor(QPalette.Window, QColor(FONDO))
    p.setColor(QPalette.WindowText, QColor(TEXTO))
    p.setColor(QPalette.Base, QColor(FONDO2))
    p.setColor(QPalette.AlternateBase, QColor(PANEL))
    p.setColor(QPalette.ToolTipBase, QColor(PANEL))
    p.setColor(QPalette.ToolTipText, QColor(TEXTO))
    p.setColor(QPalette.Text, QColor(TEXTO))
    p.setColor(QPalette.Button, QColor(PANEL))
    p.setColor(QPalette.ButtonText, QColor(TEXTO))
    p.setColor(QPalette.BrightText, QColor("#ff5555"))
    p.setColor(QPalette.Highlight, QColor(ACENTO))
    p.setColor(QPalette.HighlightedText, QColor("white"))
    p.setColor(QPalette.PlaceholderText, QColor("#888888"))
    app.setPalette(p)
    app.setStyleSheet(QSS)


def color_dcr(dcr: float) -> QColor:
    """Color de fondo de una fila según el DCR."""
    if dcr > 1.0:
        return QColor("#5a1f1f")
    if dcr > 0.9:
        return QColor("#4d3f17")
    return QColor("#1f3d29")


def color_estado(estado: str) -> QColor:
    if estado == "CUMPLE":
        return QColor(VERDE)
    if estado == "NO CUMPLE":
        return QColor(ROJO)
    return QColor("#555555")
