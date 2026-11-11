"""Muros no portantes, tabiques, parapetos y cercos ante carga perpendicular (E.070 Art. 29, Tabla 12)."""
from __future__ import annotations

import math

from ..e030 import norma
from ..models import (Check, Edificio, ElementoNoPortante, Muro, ResultadoNoPortante, Step, cargar_json, fmt,
                      fmt_MPa, fmt_mm)

_T12 = cargar_json("tabla12.json")
FACTOR_CARGA_ARRIOSTRE = 1.25      # editable: Mu = 1.25·M en el arriostre (criterio usual; revisar)


def coeficiente_m(caso: int, a: float, b: float) -> tuple[float, float]:
    """Coeficiente de momentos m de la Tabla 12 (interpolación lineal). Devuelve (m, b/a)."""
    if caso == 3:
        return _T12["caso3"]["m"], b / a if a else 0.0
    if caso == 4:
        return _T12["caso4"]["m"], b / a if a else 0.0
    t = _T12["caso1" if caso == 1 else "caso2"]
    ba = b / a
    xs, ms = t["ba"], t["m"]
    if ba <= xs[0]:
        return ms[0], ba
    if ba > xs[-2]:
        # último tramo hasta b/a = ∞: interpolación lineal en a/b
        return ms[-2] + (ms[-1] - ms[-2]) * (1 - xs[-2] / ba), ba
    for i in range(1, len(xs)):
        if ba <= xs[i]:
            x0, x1, m0, m1 = xs[i - 1], xs[i], ms[i - 1], ms[i]
            return m0 + (m1 - m0) * (ba - x0) / (x1 - x0), ba
    return ms[-1], ba


def carga_sismica(Z: float, U: float, C1: float, gamma: float, e: float) -> float:
    """w = 0.8·Z·U·C1·γ·e (MPa = N/mm², con γ en N/mm³ y e en mm)."""
    return 0.8 * Z * U * C1 * gamma * e


def _ZU(ed: Edificio) -> tuple[float, float]:
    return norma.factor_Z(ed.zona), (ed.U if ed.U else norma.factor_U(ed.categoria))


def _casos_txt(caso: int) -> str:
    return {1: "4 bordes arriostrados", 2: "3 bordes arriostrados", 3: "arriostrado en bordes horizontales",
            4: "voladizo"}[caso]


def verificar_no_portante(el: ElementoNoPortante, ed: Edificio) -> ResultadoNoPortante:
    r = ResultadoNoPortante(nombre=el.nombre, tipo=el.tipo)
    Z, U = _ZU(ed)
    C1 = el.C1 if el.C1 is not None else norma.C1_POR_TIPO.get(el.tipo, 2.0)
    e = el.e or el.t
    w = carga_sismica(Z, U, C1, el.gamma, e)
    m, ba = coeficiente_m(el.caso, el.a, el.b)
    Ms = m * w * el.a ** 2               # N·mm/mm
    fm = 6 * Ms / el.t ** 2
    ft = ed.material.ft
    pasos = [
        Step("Caso", f"{el.caso}", f"{_casos_txt(el.caso)} (E.070 Tabla 12)"),
        Step("C1", fmt(C1, 2), f"Coeficiente sísmico para '{el.tipo}' (E.030; editable)"),
        Step("w", f"{w * 1000:.3f} kN/m²", "Carga sísmica perpendicular al plano", "w = 0.8·Z·U·C1·γ·e",
             f"0.8·{Z:.2f}·{U:.2f}·{C1:.2f}·{el.gamma * 1e6:.1f}·{e / 1000:.3f} (kN/m²)"),
        Step("a", fmt_mm(el.a), "Dimensión crítica"),
        Step("b/a", fmt(ba, 3) if el.caso in (1, 2) else "—", "Relación de aspecto"),
        Step("m", fmt(m, 4), "Coeficiente de momento (Tabla 12, interpolado)"),
        Step("Ms", f"{Ms / 1000:.4f} kN·m/m", "Momento flector distribuido por unidad de longitud", "Ms = m·w·a^2",
             f"{m:.4f}·{w * 1000:.3f}·{el.a / 1000:.2f}^2 (kN·m/m)"),
        Step("fm", fmt_MPa(fm), "Esfuerzo normal por flexión", "fm = 6·Ms/t^2", f"6·{Ms:.1f}/{el.t:.0f}^2"),
        Step("f't", fmt_MPa(ft), "Resistencia a tracción por flexión de la albañilería no reforzada"),
    ]
    c = Check(f"Flexión perpendicular al plano: {el.nombre}", "E.070 Art. 29 (Tabla 12)", fm, ft, "MPa", pasos=pasos)
    r.checks.append(c)
    r.valores = {"C1": C1, "w": w, "m": m, "ba": ba, "Ms": Ms, "fm": fm, "ft": ft, "caso": el.caso}
    if fm > ft:
        c.notas.append("No cumple: colocar arriostres (columnas de concreto) a la separación indicada.")
        arr = proponer_arriostres(el, w, ft, ed)
        r.valores.update(arr["valores"])
        r.checks.append(arr["check"])
    return r


def proponer_arriostres(el: ElementoNoPortante, w: float, ft: float, ed: Edificio) -> dict:
    """Busca la separación máxima de columnas de arriostre que hace fm ≤ f't y diseña el arriostre."""
    voladizo = el.caso in (2, 4)                  # el arriostre trabaja en voladizo desde la base
    if el.caso in (3, 4):
        h = el.a
    elif el.caso == 2:
        h = el.b                                  # 3 bordes: a = borde libre (horizontal), b = altura
    else:
        h = min(el.a, el.b)
    sep = None
    caso_nuevo = 2 if voladizo else 1
    for s in range(4000, 400, -50):
        if caso_nuevo == 2:
            m, _ = coeficiente_m(2, s, h)         # a = borde libre = s, b = altura
            a = s
        else:
            a, b = min(s, h), max(s, h)
            m, _ = coeficiente_m(1, a, b)
        fm = 6 * m * w * a ** 2 / el.t ** 2
        if fm <= ft:
            sep = s
            break
    pasos = []
    if sep is None:
        c = Check(f"Arriostres para {el.nombre}", "E.070 Art. 29", 1.0, 0.0, "",
                  notas=["No se encontró separación ≥ 0.45 m que cumpla; aumentar el espesor o reforzar el muro."])
        return {"check": c, "valores": {}}
    wcol = w * sep                                 # N/mm sobre la columna
    Mcol = wcol * h ** 2 / 2 if voladizo else wcol * h ** 2 / 8
    Mu = FACTOR_CARGA_ARRIOSTRE * Mcol
    mat = ed.material
    b, dcol = el.t, 200.0
    d_util = dcol - 40.0
    As = Mu / (0.9 * mat.fy * 0.9 * d_util)
    pasos += [
        Step("s", fmt_mm(sep), f"Separación máxima de arriostres (caso {caso_nuevo})"),
        Step("w·s", f"{wcol:.3f} N/mm", "Carga tributaria sobre la columna de arriostre", "q = w·s", ""),
        Step("M", f"{Mcol / 1e6:.3f} kN·m", "Momento en el arriostre", "M = q·h^2/2" if voladizo else "M = q·h^2/8",
             f"{wcol:.3f}·{h:.0f}^2/{'2' if voladizo else '8'}"),
        Step("Mu", f"{Mu / 1e6:.3f} kN·m", "Momento último (factor editable)", f"Mu = {FACTOR_CARGA_ARRIOSTRE}·M", ""),
        Step("As", f"{As:.1f} mm²", f"Acero por cara, columna {b / 10:.0f} x {dcol / 10:.0f} cm (d = {d_util:.0f} mm)",
             "As = Mu/(φ·fy·0.9·d)", f"{Mu:.0f}/(0.9·{mat.fy:.0f}·0.9·{d_util:.0f})"),
    ]
    As_min = 2 * 50.0
    nota = f"Arriostres {b / 10:.0f} x {dcol / 10:.0f} cm @ {sep / 1000:.2f} m; acero por cara: " \
           + ("2 Ø 8 mm" if As <= As_min else f"{As:.0f} mm² (elegir barras)")
    c = Check(f"Arriostres para {el.nombre}", "E.070 Art. 29", 0.0, 1.0, "", pasos=pasos, notas=[nota], aplica=False)
    return {"check": c, "valores": {"sep_arriostre": sep, "Mu_arriostre": Mu, "As_arriostre": As}}


def verificar_portante_perpendicular(m: Muro, ed: Edificio, Pg: float) -> Check:
    """Muro portante ante carga perpendicular (4 bordes, tabique interior C1 = 2.0)."""
    Z, U = _ZU(ed)
    C1 = norma.C1_POR_TIPO.get("tabique", 2.0)
    w = carga_sismica(Z, U, C1, ed.material.gamma, m.t)
    Lp = m.Lm or m.L / max(m.Nc - 1, 1)
    a, b = min(m.h, Lp), max(m.h, Lp)
    coef, ba = coeficiente_m(1, a, b)
    Ms = coef * w * a ** 2
    fm = 6 * Ms / m.t ** 2
    fa = Pg / (m.L * m.t)
    ft = ed.material.ft
    pasos = [
        Step("w", f"{w * 1000:.3f} kN/m²", "Carga sísmica perpendicular", "w = 0.8·Z·U·C1·γ·e",
             f"0.8·{Z:.2f}·{U:.2f}·{C1:.1f}·{ed.material.gamma * 1e6:.1f}·{m.t / 1000:.3f}"),
        Step("m", fmt(coef, 4), f"Caso 1 (4 bordes), b/a = {ba:.2f}"),
        Step("Ms", f"{Ms / 1000:.4f} kN·m/m", "", "Ms = m·w·a^2", f"{coef:.4f}·{w * 1000:.3f}·{a / 1000:.2f}^2"),
        Step("fm", fmt_MPa(fm), "", "fm = 6·Ms/t^2", f"6·{Ms:.1f}/{m.t:.0f}^2"),
        Step("fa", fmt_MPa(fa), "Esfuerzo axial de gravedad", "fa = Pg/(L·t)", f"{Pg:.0f}/({m.L:.0f}·{m.t:.0f})"),
        Step("fm - fa", fmt_MPa(fm - fa), "Tracción neta", "fm - fa ≤ f't", ""),
        Step("fa + fm", fmt_MPa(fa + fm), "Compresión", "fa + fm ≤ 0.25·f'm", f"0.25·{ed.material.fm:.3f}"),
    ]
    dem = max(fm - fa, 0.0)
    c = Check(f"Carga perpendicular en muro portante {m.nombre} (tracción neta)", "E.070 Art. 29", dem, ft, "MPa",
              pasos=pasos)
    if fa + fm > 0.25 * ed.material.fm:
        c.notas.append("fa + fm > 0.25·f'm: no cumple a compresión por flexocompresión.")
        c.demanda = max(c.demanda / ft, (fa + fm) / (0.25 * ed.material.fm)) * ft
    return c
