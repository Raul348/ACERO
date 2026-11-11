"""Columnas de confinamiento y vigas soleras (E.070 Art. 27, Tabla 11)."""
from __future__ import annotations

import math

from ..models import (Check, Material, Muro, OpcionesAnalisis, ResultadoConfinamiento, Step, cargar_json,
                      fmt, fmt_kN, fmt_kNm, fmt_mm, fmt_mm2)

_T11 = cargar_json("tabla11.json")
_B = cargar_json("barras.json")
BARRAS = {b["nombre"]: b for b in _B["barras"]}
LONGITUDINALES = _B["longitudinales_columna"]
AS_MIN_4D8 = 4 * BARRAS["8 mm"]["area"]       # 4 Ø 8 mm = 200 mm²
D = _T11["diseno"]


def proponer_barras(As_req: float, cantidades=(4, 6)) -> tuple[str, float, bool]:
    """Combinación comercial con As ≥ As_req. Devuelve (texto, As provisto, ok)."""
    for n in cantidades:
        for nom in LONGITUDINALES:
            a = n * BARRAS[nom]["area"]
            if a >= As_req:
                return f"{n} Ø {nom}", a, True
    n, nom = cantidades[-1], LONGITUDINALES[-1]
    return f"{n} Ø {nom}", n * BARRAS[nom]["area"], False


def longitud_pano(m: Muro) -> tuple[float, float]:
    """(Lpaño, Lm): Lm = paño mayor o 0.5·L (el mayor); en muros de un paño Lm = L."""
    npanos = max(m.Nc - 1, 1)
    lp = m.L / npanos
    if m.Lm:
        return m.Lm, max(m.Lm, 0.5 * m.L) if m.Nc > 2 else m.L
    return lp, (m.L if m.Nc <= 2 else max(lp, 0.5 * m.L))


def cargas_columnas(m: Muro, Pg: float) -> tuple[float, float, list[Step]]:
    lp, _ = longitud_pano(m)
    pce = m.Pc_ext if m.Pc_ext is not None else Pg * lp / (2 * m.L)
    pci = m.Pc_int if m.Pc_int is not None else Pg * lp / m.L
    pasos = [Step("Pc ext", fmt_kN(pce), "Carga de gravedad de la columna extrema (mitad del paño adyacente"
                  + ("; valor ingresado)" if m.Pc_ext is not None else ")"),
                  "Pc = Pg·Lpaño/(2·L)" if m.Pc_ext is None else "Pc",
                  f"{Pg:.0f}·{lp:.0f}/(2·{m.L:.0f})" if m.Pc_ext is None else "")]
    if m.Nc > 2:
        pasos.append(Step("Pc int", fmt_kN(pci), "Carga de gravedad de la columna interior (mitad de cada paño)",
                          "Pc = Pg·Lpaño/L" if m.Pc_int is None else "Pc",
                          f"{Pg:.0f}·{lp:.0f}/{m.L:.0f}" if m.Pc_int is None else ""))
    pasos.append(Step("Nota", "", "Agregar a Pc la carga de los muros transversales si existen (Pc ext ingresable)."))
    return pce, pci, pasos


def _redondear(x: float, paso: float = 50.0) -> float:
    return math.ceil(x / paso - 1e-9) * paso


def _estribos(b: float, d: float, mat: Material, r: float, An_req: float, minimo: bool) -> tuple[str, list[Step]]:
    if minimo:
        return "[] Ø 6 mm: 1 @ 5, 4 @ 10, resto @ 25 cm", [
            Step("Estribos", "[] Ø6 mm 1@5, 4@10, r@25 cm", "Estribaje mínimo de confinamiento (E.070 Art. 27.3)")]
    tn = b - 2 * r
    Ac = b * d
    An = (b - 2 * r) * (d - 2 * r)
    pasos = []
    for nom in ("6 mm", "8 mm", "3/8\""):
        Av = 2 * BARRAS[nom]["area"]
        s1 = Av * mat.fy / (0.3 * tn * mat.fc * (Ac / An - 1)) if Ac > An else math.inf
        s2 = Av * mat.fy / (0.12 * tn * mat.fc)
        s3 = max(d / 4, 50.0)
        s4 = 100.0
        s = min(s1, s2, s3, s4)
        if s >= 50.0:
            break
    s_ad = max(math.floor(s / 10.0) * 10.0, 50.0)
    zona = max(450.0, 1.5 * d)
    n = math.ceil((zona - 50.0) / s_ad)
    pasos += [
        Step("Av", f"{Av:.0f} mm²", f"Dos ramas de estribo Ø {nom}"),
        Step("tn", fmt_mm(tn), "Espesor del núcleo confinado", "tn = t - 2·r", f"{b:.0f} - 2·{r:.0f}"),
        Step("s1", fmt_mm(s1, 1), "", "s1 = Av·fy/(0.3·tn·f'c·(Ac/An - 1))",
             f"{Av:.0f}·{mat.fy:.0f}/(0.3·{tn:.0f}·{mat.fc:.2f}·({Ac:.0f}/{An:.0f} - 1))"),
        Step("s2", fmt_mm(s2, 1), "", "s2 = Av·fy/(0.12·tn·f'c)", f"{Av:.0f}·{mat.fy:.0f}/(0.12·{tn:.0f}·{mat.fc:.2f})"),
        Step("s3", fmt_mm(s3, 1), "", "s3 = d/4 ≥ 50 mm", f"{d:.0f}/4"),
        Step("s4", fmt_mm(s4), "", "s4 = 100 mm", ""),
        Step("s", fmt_mm(s_ad), "Separación adoptada en la zona confinada (≥ 50 mm)", "s = mín(s1, s2, s3, s4)", ""),
        Step("Lconf", fmt_mm(zona), "Longitud de la zona confinada en cada extremo", "Lconf = máx(450 mm, 1.5·d)",
             f"máx(450, 1.5·{d:.0f})"),
    ]
    txt = f"[] Ø {nom}: 1 @ 5, {n} @ {s_ad / 10:.0f}, resto @ 25 cm (zona de {zona / 10:.0f} cm en cada extremo)"
    return txt, pasos


def diseno_columna(nombre: str, Vc: float, T: float, C: float, m: Muro, mat: Material, op: OpcionesAnalisis,
                   agrietado: bool, delta: float) -> dict:
    """Dimensiona una columna (corte-fricción, tracción, núcleo) y propone refuerzo."""
    t = m.t
    r = op.recubrimiento_nucleo
    pasos: list[Step] = []
    phi = D["phi_cf"]
    mu = D["mu_rugosa"] if m.junta_rugosa else D["mu_sin_tratar"]
    Acf = Vc / (0.2 * mat.fc * phi) if agrietado else 0.0
    Ac_min = 150.0 * t                      # 15·t (cm², t en cm) = 150·t mm²
    if agrietado:
        Asf = Vc / (mat.fy * mu * phi)
        Ast = max(T, 0.0) / (mat.fy * phi)
        pasos += [
            Step("Acf", fmt_mm2(Acf), "Área por corte-fricción", "Acf = Vc/(0.2·f'c·φ)",
                 f"{Vc:.0f}/(0.2·{mat.fc:.2f}·{phi})"),
            Step("Asf", fmt_mm2(Asf), f"Acero por corte-fricción (μ = {mu})", "Asf = Vc/(fy·μ·φ)",
                 f"{Vc:.0f}/({mat.fy:.0f}·{mu}·{phi})"),
            Step("Ast", fmt_mm2(Ast), "Acero por tracción", "Ast = T/(fy·φ)", f"{max(T, 0):.0f}/({mat.fy:.0f}·{phi})"),
        ]
        As = Asf + Ast
        pasos.append(Step("As", fmt_mm2(As), "Acero vertical requerido", "As = Asf + Ast", f"{Asf:.1f} + {Ast:.1f}"))
    else:
        phi = D["phi_superior"]
        As = max(T, 0.0) / (mat.fy * phi)
        pasos.append(Step("As", fmt_mm2(As), "Acero por tracción (muro no agrietado)", "As = T/(φ·fy)",
                          f"{max(T, 0):.0f}/({phi}·{mat.fy:.0f})"))
    phic = D["phi_compresion_estribos"]
    if C > 0:
        An = As + (C / phic - As * mat.fy) / (0.85 * delta * mat.fc)
        An = max(An, 0.0)
        pasos.append(Step("An", fmt_mm2(An), f"Área del núcleo confinado (φ = {phic}, δ = {delta})",
                          "An = As + (C/φ - As·fy)/(0.85·δ·f'c)",
                          f"{As:.1f} + ({C:.0f}/{phic} - {As:.1f}·{mat.fy:.0f})/(0.85·{delta}·{mat.fc:.2f})"))
    else:
        An = 0.0
    d_cf = Acf / t
    d_min15 = Ac_min / t
    d_an = An / (t - 2 * r) + 2 * r if An > 0 else 0.0
    d = _redondear(max(op.dmin_columna, d_cf, d_min15, d_an))
    Ac = t * d
    pasos += [
        Step("Ac mín", fmt_mm2(Ac_min), "Área mínima de la columna: 15·t (cm², t en cm)", "Acmín = 15·t", f"15·{t / 10:.1f} cm²"),
        Step("dAn", fmt_mm(d_an), "Peralte para alojar el núcleo An", "dAn = An/(t - 2·r) + 2·r",
             f"{An:.0f}/({t:.0f} - 2·{r:.0f}) + 2·{r:.0f}" if An > 0 else "An = 0"),
        Step("d", fmt_mm(d), "Peralte adoptado (múltiplo de 5 cm)",
             "d ≥ máx(dmín, Acf/t, Acmín/t, dAn)",
             f"máx({op.dmin_columna:.0f}, {d_cf:.0f}, {d_min15:.0f}, {d_an:.0f})"),
        Step("Ac", fmt_mm2(Ac), "Sección adoptada", "Ac = t·d", f"{t:.0f}·{d:.0f}"),
    ]
    As_min = 0.1 * mat.fc * Ac / mat.fy
    As_dis = max(As, As_min, AS_MIN_4D8)
    txt, As_prov, ok = proponer_barras(As_dis)
    pasos += [
        Step("As mín", fmt_mm2(As_min), "Cuantía mínima", "As ≥ 0.1·f'c·Ac/fy", f"0.1·{mat.fc:.2f}·{Ac:.0f}/{mat.fy:.0f}"),
        Step("As dis", fmt_mm2(As_dis), "Acero de diseño (mínimo 4 Ø 8 mm = 200 mm²)", "As = máx(As, Asmín, 200 mm²)", ""),
        Step("Refuerzo", f"{txt} ({As_prov:.0f} mm²)", "Combinación comercial propuesta"),
    ]
    est_txt, est_pasos = _estribos(t, d, mat, r, An, minimo=not agrietado)
    pasos += est_pasos
    An_prov = (t - 2 * r) * (d - 2 * r)
    checks = []
    if agrietado:
        checks.append(Check(f"Columna {nombre}: sección por corte-fricción", "E.070 Art. 27.3", Acf, Ac, "mm²"))
    checks.append(Check(f"Columna {nombre}: acero vertical", "E.070 Art. 27.3" if agrietado else "E.070 Art. 27.4",
                        As_dis, As_prov, "mm²",
                        notas=[] if ok else ["Ninguna combinación de 4 ó 6 barras alcanza; aumentar la sección."]))
    if An > 0:
        checks.append(Check(f"Columna {nombre}: núcleo confinado a compresión", "E.070 Art. 27.3", An, An_prov, "mm²"))
    return {"Acf": Acf, "As": As, "As_dis": As_dis, "An": An, "d": d, "Ac": Ac, "txt": txt, "As_prov": As_prov,
            "estribos": est_txt, "pasos": pasos, "checks": checks}


def diseno_confinamiento(m: Muro, mat: Material, op: OpcionesAnalisis, Pg: float, Vm: float, Mu: float,
                         Vu: float, tipo: str) -> ResultadoConfinamiento:
    """tipo: 'agrietado' (Tabla 11), 'no agrietado' (Art. 27.4) o 'mínimo' (ΣVm ≥ 3·VE)."""
    rc = ResultadoConfinamiento(tipo=tipo)
    h, L, Nc = m.h, m.L, max(m.Nc, 2)
    lp, Lm = longitud_pano(m)
    pce, pci, pasos_pc = cargas_columnas(m, Pg)
    rc.Pc_ext, rc.Pc_int = pce, pci
    delta_ext = D["delta_con_muro_transversal"] if m.muro_transversal else D["delta_sin_muro_transversal"]
    delta_int = D["delta_sin_muro_transversal"]
    pasos_g = [Step("Lm", fmt_mm(Lm), "Longitud del paño mayor o 0.5·L (en muros de un paño Lm = L)"),
               Step("Nc", str(Nc), "Número de columnas de confinamiento")] + pasos_pc

    if tipo == "agrietado":
        k = _T11["extrema"]["kVc"]
        rc.M = Mu - 0.5 * Vm * h
        rc.F = rc.M / L
        rc.Vc_ext = k * Vm * Lm / (L * (Nc + 1))
        rc.T_ext = rc.F - pce
        rc.C_ext = pce + rc.F
        pasos_g += [
            Step("M", fmt_kNm(rc.M), "", "M = Mu - 0.5·Vm·h", f"{Mu / 1e6:.2f} - 0.5·{Vm / 1000:.2f}·{h / 1000:.2f}"),
            Step("F", fmt_kN(rc.F), "Fuerza axial por flexión", "F = M/L", f"{rc.M / 1e6:.2f}/{L / 1000:.2f}"),
            Step("Vc ext", fmt_kN(rc.Vc_ext), "Cortante en columna extrema (Tabla 11)", "Vc = 1.5·Vm·Lm/(L·(Nc + 1))",
                 f"{k}·{Vm / 1000:.2f}·{Lm / 1000:.2f}/({L / 1000:.2f}·({Nc} + 1))"),
            Step("T ext", fmt_kN(rc.T_ext), "Tracción en columna extrema", "T = F - Pc",
                 f"{rc.F / 1000:.2f} - {pce / 1000:.2f}"),
            Step("C ext", fmt_kN(rc.C_ext), "Compresión en columna extrema", "C = Pc + F",
                 f"{pce / 1000:.2f} + {rc.F / 1000:.2f}"),
        ]
        if Nc > 2:
            ki = _T11["interior"]["kVc"]
            rc.Vc_int = ki * Vm * Lm / (L * (Nc + 1))
            rc.T_int = Vm * h / L - pci - m.Pt
            rc.C_int = pci - 0.5 * Vm * h / L
            pasos_g += [
                Step("Vc int", fmt_kN(rc.Vc_int), "Cortante en columna interior", "Vc = Vm·Lm/(L·(Nc + 1))",
                     f"{Vm / 1000:.2f}·{Lm / 1000:.2f}/({L / 1000:.2f}·({Nc} + 1))"),
                Step("T int", fmt_kN(rc.T_int), "Tracción en columna interior", "T = Vm·h/L - Pc - Pt",
                     f"{Vm / 1000:.2f}·{h / 1000:.2f}/{L / 1000:.2f} - {pci / 1000:.2f} - {m.Pt / 1000:.2f}"),
                Step("C int", fmt_kN(rc.C_int), "Compresión en columna interior", "C = Pc - 0.5·Vm·h/L",
                     f"{pci / 1000:.2f} - 0.5·{Vm / 1000:.2f}·{h / 1000:.2f}/{L / 1000:.2f}"),
            ]
        ext = diseno_columna("extrema", rc.Vc_ext, rc.T_ext, rc.C_ext, m, mat, op, True, delta_ext)
        intr = diseno_columna("interior", rc.Vc_int, rc.T_int, rc.C_int, m, mat, op, True, delta_int) if Nc > 2 else None
        Ts = Vm * Lm / (2 * L)
        ts_f, ts_s = "Ts = Vm·Lm/(2·L)", f"{Vm / 1000:.2f}·{Lm / 1000:.2f}/(2·{L / 1000:.2f})"
    elif tipo == "no agrietado":
        rc.F = Mu / L
        rc.T_ext = rc.F - pce
        rc.C_ext = pce + rc.F
        pasos_g += [
            Step("F", fmt_kN(rc.F), "Fuerza axial por flexión", "F = Mu/L", f"{Mu / 1e6:.2f}/{L / 1000:.2f}"),
            Step("T ext", fmt_kN(rc.T_ext), "Tracción en columna extrema", "T = F - Pc", f"{rc.F / 1000:.2f} - {pce / 1000:.2f}"),
            Step("C ext", fmt_kN(rc.C_ext), "Compresión en columna extrema", "C = Pc + F", f"{pce / 1000:.2f} + {rc.F / 1000:.2f}"),
        ]
        ext = diseno_columna("extrema", 0.0, rc.T_ext, rc.C_ext, m, mat, op, False, delta_ext)
        intr = diseno_columna("interior", 0.0, 0.0, 0.0, m, mat, op, False, delta_int) if Nc > 2 else None
        Ts = 0.5 * Vu * Lm / L
        ts_f, ts_s = "Ts = 0.5·Vu·Lm/L", f"0.5·{Vu / 1000:.2f}·{Lm / 1000:.2f}/{L / 1000:.2f}"
    else:   # mínimo
        ext = diseno_columna("extrema", 0.0, 0.0, 0.0, m, mat, op, False, delta_ext)
        intr = diseno_columna("interior", 0.0, 0.0, 0.0, m, mat, op, False, delta_int) if Nc > 2 else None
        Ts = 0.0
        ts_f, ts_s = "Ts = 0 (refuerzo mínimo)", ""

    rc.Acf = ext["Acf"]
    rc.An_req = ext["An"]
    rc.As_req_ext = ext["As_dis"]
    rc.b_col, rc.d_col = m.t, ext["d"]
    rc.seccion_ext = f"{m.t / 10:.0f} x {ext['d'] / 10:.0f} cm"
    rc.refuerzo_ext = ext["txt"]
    rc.As_prov_ext = ext["As_prov"]
    rc.estribos = ext["estribos"]
    rc.checks += ext["checks"]
    if intr:
        rc.As_req_int = intr["As_dis"]
        rc.d_col_int = intr["d"]
        rc.seccion_int = f"{m.t / 10:.0f} x {intr['d'] / 10:.0f} cm"
        rc.refuerzo_int = intr["txt"]
        rc.As_prov_int = intr["As_prov"]
        rc.checks += intr["checks"]

    # ---- solera
    hs = op.peralte_losa
    Acs = m.t * hs
    phis = D["phi_solera"]
    As_s = Ts / (phis * mat.fy)
    As_smin = 0.1 * mat.fc * Acs / mat.fy
    As_sdis = max(As_s, As_smin, AS_MIN_4D8)
    txt_s, As_sprov, ok_s = proponer_barras(As_sdis, cantidades=(4, 6))
    rc.Ts, rc.As_solera, rc.refuerzo_solera = Ts, As_sdis, txt_s
    rc.seccion_solera = f"{m.t / 10:.0f} x {hs / 10:.0f} cm"
    rc.estribos_solera = "[] Ø 6 mm: 1 @ 5, 4 @ 10, resto @ 25 cm"
    pasos_s = [
        Step("Ts", fmt_kN(Ts), "Tracción en la solera", ts_f, ts_s),
        Step("As", fmt_mm2(As_s), f"Acero por tracción (φ = {phis})", "As = Ts/(φ·fy)", f"{Ts:.0f}/({phis}·{mat.fy:.0f})"),
        Step("Acs", fmt_mm2(Acs), "Sección de la solera = t × peralte de losa", "Acs = t·hs", f"{m.t:.0f}·{hs:.0f}"),
        Step("As mín", fmt_mm2(As_smin), "", "As ≥ 0.1·f'c·Acs/fy", f"0.1·{mat.fc:.2f}·{Acs:.0f}/{mat.fy:.0f}"),
        Step("Refuerzo", f"{txt_s} ({As_sprov:.0f} mm²)", "Mínimo 4 Ø 8 mm"),
        Step("Estribos", rc.estribos_solera, "Estribaje mínimo"),
    ]
    rc.checks.append(Check("Solera: acero longitudinal", "E.070 Art. 27.3" if tipo == "agrietado" else "E.070 Art. 27.4",
                           As_sdis, As_sprov, "mm²", pasos=pasos_s))

    # todos los pasos de la columna extrema van en su primer check
    for c in rc.checks:
        if c.nombre.startswith("Columna extrema"):
            c.pasos = pasos_g + ext["pasos"]
            break
    if intr:
        for c in rc.checks:
            if c.nombre.startswith("Columna interior"):
                c.pasos = intr["pasos"]
                break
    if tipo == "mínimo":
        rc.notas.append("Entrepiso con ΣVm ≥ 3·VE: columnas y soleras con refuerzo mínimo.")
    return rc
