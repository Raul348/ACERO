"""Rigidez lateral de muros y reparto del cortante de piso (modo manual)."""
from __future__ import annotations

import math

from ..models import Edificio, Muro, Step


def rigidez_lateral(m: Muro, Em: float, Gm: float, condicion: str = "voladizo", f: float = 1.2) -> float:
    """K = 1/(h³/(3·Em·I) + f·h/(Gm·A))  (voladizo)  ó  h³/(12·Em·I) (doble empotramiento). N/mm."""
    c = 3.0 if condicion == "voladizo" else 12.0
    flex = m.h ** 3 / (c * Em * m.inercia)
    corte = f * m.h / (Gm * m.area)
    return 1.0 / (flex + corte)


def pasos_rigidez(m: Muro, Em: float, Gm: float, condicion: str, f: float) -> list[Step]:
    c = 3 if condicion == "voladizo" else 12
    K = rigidez_lateral(m, Em, Gm, condicion, f)
    return [
        Step("I", f"{m.inercia / 1e12:.4f}e12 mm⁴", "Inercia de la sección (sin alas)", "I = t·L^3/12",
             f"{m.t:.0f}·{m.L:.0f}^3/12"),
        Step("A", f"{m.area:.0f} mm²", "Área de corte", "A = t·L", f"{m.t:.0f}·{m.L:.0f}"),
        Step("K", f"{K:.1f} N/mm", f"Rigidez lateral ({condicion}, f = {f})",
             f"K = 1/(h^3/({c}·Em·I) + f·h/(Gm·A))",
             f"1/({m.h:.0f}^3/({c}·{Em:.0f}·{m.inercia:.4e}) + {f}·{m.h:.0f}/({Gm:.0f}·{m.area:.0f}))"),
    ]


def repartir_cortante(ed: Edificio, piso: int, direccion: str, V: float) -> dict[str, dict]:
    """Reparte el cortante V (N) del piso entre los muros portantes de la dirección,
    en proporción a su rigidez. Con torsión: excentricidad real + accidental (Lx, Ly)."""
    op = ed.opciones
    Em, Gm = ed.material.Em_calc, ed.material.Gm_calc
    todos = [m for m in ed.muros_piso(piso) if m.portante]
    muros = [m for m in todos if m.direccion == direccion]
    if not muros:
        return {}
    K = {m.nombre: rigidez_lateral(m, Em, Gm, op.rigidez, op.factor_forma) for m in todos}
    sumK = sum(K[m.nombre] for m in muros)
    res = {m.nombre: {"K": K[m.nombre], "sumK": sumK, "Vdir": V * K[m.nombre] / sumK, "Vtor": 0.0,
                      "V": V * K[m.nombre] / sumK} for m in muros}
    if not op.torsion:
        return res
    mx = [m for m in todos if m.direccion == "X"]
    my = [m for m in todos if m.direccion == "Y"]
    if not mx or not my:
        return res
    Kx = sum(K[m.nombre] for m in mx)
    Ky = sum(K[m.nombre] for m in my)
    xr = sum(K[m.nombre] * m.x for m in my) / Ky
    yr = sum(K[m.nombre] * m.y for m in mx) / Kx
    Lx = ed.Lx or math.sqrt(ed.Ap)
    Ly = ed.Ly or math.sqrt(ed.Ap)
    xcm, ycm = Lx / 2.0, Ly / 2.0
    J = sum(K[m.nombre] * (m.y - yr) ** 2 for m in mx) + sum(K[m.nombre] * (m.x - xr) ** 2 for m in my)
    if J <= 0:
        return res
    if direccion == "X":
        e0, B = ycm - yr, Ly
    else:
        e0, B = xcm - xr, Lx
    for m in muros:
        d = (m.y - yr) if direccion == "X" else (m.x - xr)
        mejores = []
        for signo in (+1, -1):
            e = e0 + signo * op.exc_accidental * B
            Mt = V * e
            # ΔV = Mt·K·d/J: aumenta el cortante en los muros del lado del centro de masas
            dv = Mt * K[m.nombre] * d / J
            mejores.append(dv)
        dv = max(mejores)
        r = res[m.nombre]
        r["Vtor"] = max(dv, 0.0)        # no se reduce el cortante directo (criterio conservador)
        r["V"] = r["Vdir"] + r["Vtor"]
        r.update({"xr": xr, "yr": yr, "e0": e0, "J": J})
    return res
