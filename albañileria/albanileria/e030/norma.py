"""Tablas de la NTE E.030 (2018) usadas por el programa."""
from __future__ import annotations

from ..models import cargar_json

_T = cargar_json("e030.json")


def factor_Z(zona: int) -> float:
    return _T["Z"][str(int(zona))]


def factor_S(zona: int, suelo: str) -> float:
    return _T["S"][str(int(zona))][suelo]


def periodos_suelo(suelo: str) -> tuple[float, float]:
    tp, tl = _T["TpTl"][suelo]
    return tp, tl


def factor_U(categoria: str) -> float:
    return _T["U"][categoria]


def porcentaje_carga_viva(categoria: str) -> float:
    return _T["porcentaje_carga_viva"][categoria]


R0_ALBANILERIA = _T["R0_albanileria"]
CT_ALBANILERIA = _T["CT_albanileria"]
C_SOBRE_R_MIN = _T["C_sobre_R_min"]
DERIVA_MAX = _T["deriva_max_albanileria"]
C1_POR_TIPO = {k: v for k, v in _T["C1_elementos_no_estructurales"].items() if not k.startswith("_")}


def factor_C(T: float, Tp: float, Tl: float) -> float:
    """Factor de amplificación sísmica (E.030 Art. 14)."""
    if T < Tp:
        return 2.5
    if T < Tl:
        return 2.5 * Tp / T
    return 2.5 * Tp * Tl / T ** 2


def exponente_k(T: float) -> float:
    """Exponente de distribución en altura (E.030 Art. 28.3.2)."""
    return 1.0 if T <= 0.5 else min(0.75 + 0.5 * T, 2.0)
