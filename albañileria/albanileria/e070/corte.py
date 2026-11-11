"""Control de fisuración y resistencia al corte (E.070 Art. 26)."""
from __future__ import annotations

from ..models import Check, Material, Muro, Step, fmt, fmt_kN, fmt_kNm


def factor_alfa(Ve: float, Me: float, L: float) -> tuple[float, float]:
    """α = Ve·L/Me acotado a 1/3 ≤ α ≤ 1. Devuelve (α acotado, α sin acotar)."""
    if Me <= 0:
        return 1.0, float("inf")
    a = Ve * L / Me
    return min(max(a, 1.0 / 3.0), 1.0), a


def resistencia_corte(m: Muro, mat: Material, Ve: float, Me: float, Pg: float) -> tuple[float, float, list[Step]]:
    """Vm = 0.5·v'm·α·t·L + 0.23·Pg (0.35 para sílico-calcáreo). Devuelve (Vm, α, pasos)."""
    alfa, a0 = factor_alfa(Ve, Me, m.L)
    k = mat.coef_vm
    Vm = k * mat.vm * alfa * m.t * m.L + 0.23 * Pg
    pasos = [
        Step("Ve", fmt_kN(Ve), "Cortante del muro ante sismo moderado"),
        Step("Me", fmt_kNm(Me), "Momento del muro ante sismo moderado"),
        Step("α", fmt(alfa), "Factor de reducción por esbeltez, 1/3 ≤ α ≤ 1"
             + (f" (sin acotar = {a0:.3f})" if abs(a0 - alfa) > 1e-9 else ""),
             "α = Ve·L/Me", f"{Ve:.0f}·{m.L:.0f}/{Me:.4e}"),
        Step("Pg", fmt_kN(Pg), "Carga de gravedad con sobrecarga reducida (D + 25 % L)"),
        Step("Vm", fmt_kN(Vm), "Resistencia al agrietamiento diagonal",
             f"Vm = {k}·v'm·α·t·L + 0.23·Pg",
             f"{k}·{mat.vm:.4f}·{alfa:.3f}·{m.t:.0f}·{m.L:.0f} + 0.23·{Pg:.0f}"),
    ]
    return Vm, alfa, pasos


def check_fisuracion(Ve: float, Vm: float, pasos_vm: list[Step]) -> Check:
    lim = 0.55 * Vm
    c = Check("Control de fisuración ante sismo moderado", "E.070 Art. 26.2", Ve, lim, "kN",
              pasos=list(pasos_vm) + [
                  Step("0.55·Vm", fmt_kN(lim), "Fuerza cortante admisible", "Ve ≤ 0.55·Vm", f"0.55·{Vm / 1000:.2f} kN")])
    if Ve > lim:
        c.notas.append("El muro se fisura ante sismo moderado: aumentar t o L, usar unidad de mayor v'm, "
                       "o agregar muros para reducir Ve.")
    return c


def check_resistencia_global(piso: int, direccion: str, suma_Vm: float, VE: float, nombres: list[str],
                             vms: list[float]) -> Check:
    pasos = [Step(f"Vm {n}", fmt_kN(v), "") for n, v in zip(nombres, vms)]
    rel = suma_Vm / VE if VE > 0 else float("inf")
    pasos += [
        Step("ΣVm", fmt_kN(suma_Vm), "Suma de resistencias de los muros portantes", "ΣVm",
             " + ".join(f"{v / 1000:.1f}" for v in vms)),
        Step("VE", fmt_kN(VE), "Cortante del entrepiso ante sismo severo"),
        Step("ΣVm/VE", fmt(rel), "≥ 3: comportamiento elástico ante sismo severo (refuerzo mínimo)",
             "ΣVm/VE", f"{suma_Vm / 1000:.1f}/{VE / 1000:.1f}"),
    ]
    c = Check(f"Resistencia al corte del entrepiso (piso {piso}, dirección {direccion})", "E.070 Art. 26.4",
              VE, suma_Vm, "kN", pasos=pasos)
    if rel >= 3:
        c.notas.append("ΣVm ≥ 3·VE: el entrepiso se comporta elásticamente ante el sismo severo; "
                       "los muros llevan refuerzo mínimo (E.070 Art. 26.4).")
    elif rel < 1:
        c.notas.append("ΣVm < VE: resistencia insuficiente; agregar muros, aumentar t o usar placas de concreto.")
    return c
