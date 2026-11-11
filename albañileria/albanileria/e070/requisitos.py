"""Requisitos estructurales mínimos (E.070 Art. 19 y Art. 20)."""
from __future__ import annotations

from ..e030 import norma
from ..models import Check, Edificio, Muro, Step, fmt, fmt_kN, fmt_MPa, fmt_mm

L_MIN_DENSIDAD = 1200.0       # mm: muros con L ≥ 1.20 m cuentan para la densidad


def check_espesor(m: Muro, zona: int) -> Check:
    div = 25.0 if zona == 1 else 20.0
    treq = m.h / div
    return Check(
        "Espesor efectivo mínimo", "E.070 Art. 19.1.a", treq, m.t, "mm",
        pasos=[Step("h", fmt_mm(m.h), "Altura libre entre arriostres horizontales"),
               Step("t mín", fmt_mm(treq, 1), f"Zona {zona}", f"t ≥ h/{div:.0f}", f"{m.h:.0f}/{div:.0f}"),
               Step("t", fmt_mm(m.t), "Espesor efectivo del muro (sin tarrajeo)")],
        notas=["Si no cumple: aumentar el espesor (muro en cabeza) o reducir la altura libre."] if m.t < treq else [])


def check_esfuerzo_axial(m: Muro, Pm: float, fm: float) -> Check:
    sigma = Pm / (m.L * m.t)
    esb = 0.2 * fm * (1 - (m.h / (35 * m.t)) ** 2)
    lim = 0.15 * fm
    cap = min(esb, lim)
    c = Check(
        "Esfuerzo axial máximo", "E.070 Art. 19.1.b", sigma, cap, "MPa",
        pasos=[Step("Pm", fmt_kN(Pm), "Carga de gravedad máxima de servicio (D + L)"),
               Step("σm", fmt_MPa(sigma), "Esfuerzo axial", "σm = Pm/(L·t)", f"{Pm:.0f}/({m.L:.0f}·{m.t:.0f})"),
               Step("Fa1", fmt_MPa(esb), "Límite por esbeltez", "Fa = 0.2·f'm·(1 - (h/(35·t))^2)",
                    f"0.2·{fm:.3f}·(1 - ({m.h:.0f}/(35·{m.t:.0f}))^2)"),
               Step("Fa2", fmt_MPa(lim), "Límite absoluto", "Fa ≤ 0.15·f'm", f"0.15·{fm:.3f}"),
               Step("Fa", fmt_MPa(cap), "Esfuerzo admisible", "Fa = mín(Fa1, Fa2)", "")])
    if sigma > cap:
        c.notas.append("No cumple: aumentar el espesor t, usar una unidad de mayor f'm o reducir la carga.")
    return c


def check_separacion_columnas(m: Muro) -> Check:
    Lpano = m.Lm if m.Lm else (m.L / max(m.Nc - 1, 1))
    lim = min(2 * m.h, 5000.0)
    return Check(
        "Separación máxima entre columnas de confinamiento", "E.070 Art. 20", Lpano, lim, "mm",
        pasos=[Step("Lpaño", fmt_mm(Lpano), "Longitud del paño mayor", "Lpaño = Lm  ó  L/(Nc - 1)",
                    f"{m.L:.0f}/({m.Nc} - 1)" if not m.Lm else f"{m.Lm:.0f}"),
               Step("Lmáx", fmt_mm(lim), "Separación máxima", "Lmáx = mín(2·h, 5 m)", f"mín(2·{m.h:.0f}, 5000)")],
        notas=["Agregar columnas de confinamiento intermedias."] if Lpano > lim else [])


def avisos_muro(m: Muro, N: int) -> list[str]:
    a = []
    if m.portante and not m.confinado:
        a.append(f"{m.nombre}: muro portante sin confinar. La E.070 exige confinamiento en los 4 bordes "
                 "(columnas y soleras) — FUERA DE ALCANCE el diseño de muros no confinados/armados.")
    if m.tabique:
        a.append(f"{m.nombre}: tiene tabique o alféizar no aislado que puede acortar la altura del muro "
                 "(columna corta). Aislar con junta ≥ 1 cm o considerarlo en el análisis (E.070 Art. 19).")
    if m.aberturas:
        a.append(f"{m.nombre}: aberturas '{m.aberturas}'. Los vanos deben estar bordeados por elementos de "
                 "confinamiento; verifique que L sea la longitud del paño sin vanos.")
    if m.L < L_MIN_DENSIDAD and m.portante:
        a.append(f"{m.nombre}: L = {m.L / 1000:.2f} m < 1.20 m; no se considera en la densidad de muros.")
    return a


def densidad_muros(ed: Edificio, piso: int, direccion: str) -> tuple[float, float, Check]:
    Z = norma.factor_Z(ed.zona)
    U = ed.U if ed.U else norma.factor_U(ed.categoria)
    S = norma.factor_S(ed.zona, ed.suelo)
    muros = [m for m in ed.muros_piso(piso, direccion)
             if m.portante and m.confinado and m.L >= L_MIN_DENSIDAD]
    suma = sum(m.L * m.t for m in muros)
    dens = suma / ed.Ap if ed.Ap else 0.0
    req = Z * U * S * ed.N / 56.0
    pasos = [Step(m.nombre, f"{m.L * m.t / 1e6:.4f} m²", f"L = {m.L / 1000:.2f} m, t = {m.t / 1000:.3f} m",
                  "L·t", f"{m.L / 1000:.2f}·{m.t / 1000:.3f}") for m in muros]
    pasos += [
        Step("ΣL·t", f"{suma / 1e6:.4f} m²", f"Muros portantes confinados con L ≥ 1.20 m, piso {piso}, dir. {direccion}"),
        Step("Ap", f"{ed.Ap / 1e6:.2f} m²", "Área de la planta típica"),
        Step("ΣL·t/Ap", fmt(dens, 4), "Densidad existente", "ΣL·t/Ap", f"{suma / 1e6:.4f}/{ed.Ap / 1e6:.2f}"),
        Step("Z·U·S·N/56", fmt(req, 4), "Densidad requerida", "Z·U·S·N/56", f"{Z:.2f}·{U:.2f}·{S:.2f}·{ed.N}/56"),
    ]
    c = Check(f"Densidad mínima de muros (piso {piso}, dirección {direccion})", "E.070 Art. 19.2.b",
              req, dens, "", pasos=pasos)
    if dens < req:
        c.notas.append("No cumple: agregar muros en esta dirección, aumentar su espesor o usar placas de concreto "
                       "(convertidas a albañilería equivalente con Ec/Em — FUERA DE ALCANCE).")
    return dens, req, c
