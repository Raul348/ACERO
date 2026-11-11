"""Diseño por sismo severo: factor de amplificación, Vu, Mu y agrietamiento (E.070 Art. 27)."""
from __future__ import annotations

from ..models import Check, Step, fmt, fmt_kN, fmt_kNm, fmt_MPa


def factor_amplificacion(Vm1: float, Ve1: float) -> tuple[float, float]:
    """Vm1/Ve1 limitado a 2 ≤ Vm1/Ve1 ≤ 3. Devuelve (acotado, sin acotar)."""
    if Ve1 <= 0:
        return 3.0, float("inf")
    r = Vm1 / Ve1
    return min(max(r, 2.0), 3.0), r


def diseno_severo(piso: int, Ve: float, Me: float, Vm: float, factor: float, factor_bruto: float,
                  origen_factor: str) -> tuple[float, float, bool, Check]:
    """Devuelve (Vu, Mu, agrietado, check informativo)."""
    Vu = Ve * factor
    Mu = Me * factor
    if piso == 1:
        agrietado = True
        nota = "Primer piso: se asume agrietado por corte ante el sismo severo; se diseña por capacidad (Tabla 11)."
    else:
        agrietado = Vu >= Vm
        nota = ("Vu ≥ Vm: el muro se agrieta por corte; se diseña como los del primer piso (Tabla 11)."
                if agrietado else
                "Vu < Vm: el muro no se agrieta; columnas extremas por tracción/compresión del momento Mu, "
                "resto con refuerzo mínimo (E.070 Art. 27.4).")
    pasos = [
        Step("Vm1/Ve1", fmt(factor_bruto), f"Relación del muro en el primer piso ({origen_factor})", "Vm1/Ve1", ""),
        Step("f", fmt(factor), "Factor de amplificación acotado", "2 ≤ Vm1/Ve1 ≤ 3", ""),
        Step("Vu", fmt_kN(Vu), "Cortante último (sismo severo)", "Vu = Ve·(Vm1/Ve1)", f"{Ve / 1000:.2f}·{factor:.3f}"),
        Step("Mu", fmt_kNm(Mu), "Momento último (sismo severo)", "Mu = Me·(Vm1/Ve1)", f"{Me / 1e6:.2f}·{factor:.3f}"),
        Step("Vm", fmt_kN(Vm), "Resistencia al agrietamiento diagonal del muro"),
    ]
    c = Check("Diseño por sismo severo (Vu, Mu, agrietamiento)", "E.070 Art. 27.1", Vu, Vm, "kN",
              pasos=pasos, notas=[nota, "Verificación informativa: define el tipo de diseño del confinamiento."],
              aplica=False)
    return Vu, Mu, agrietado, c


def refuerzo_horizontal(nombre: str, piso: int, N: int, Vu: float, Vm: float, sigma_m: float, fm: float,
                        t: float) -> tuple[bool, Check]:
    """Necesidad de refuerzo horizontal continuo (E.070 Art. 27.1)."""
    motivos = []
    if piso > 1 and Vu >= Vm:
        motivos.append("Vu ≥ Vm")
    if sigma_m >= 0.05 * fm:
        motivos.append("σm ≥ 0.05·f'm")
    if N > 3 and piso == 1:
        motivos.append("edificio de más de 3 pisos: muros portantes del primer nivel")
    req = bool(motivos)
    As_min = 0.001 * 200.0 * t
    pasos = [
        Step("σm", fmt_MPa(sigma_m), "Esfuerzo axial con toda la sobrecarga (Pm/(L·t))"),
        Step("0.05·f'm", fmt_MPa(0.05 * fm), "", "0.05·f'm", f"0.05·{fm:.3f}"),
        Step("ρ mín", "0.001", "Cuantía mínima de refuerzo horizontal", "ρ = As/(s·t) ≥ 0.001", ""),
        Step("As", f"{As_min:.1f} mm² cada 200 mm", "Refuerzo horizontal mínimo (s = 2 hiladas ≈ 200 mm)",
             "As = 0.001·s·t", f"0.001·200·{t:.0f}"),
    ]
    c = Check("Refuerzo horizontal continuo en el muro", "E.070 Art. 27.1", 0.0, 1.0, "", pasos=pasos,
              notas=([f"REQUIERE refuerzo horizontal ({', '.join(motivos)}): ρ ≥ 0.001, anclado en las columnas "
                      "(p.ej. 1 Ø 1/4\" cada 2 hiladas)."] if req else
                     ["No requiere refuerzo horizontal por este artículo."]), aplica=False)
    return req, c
