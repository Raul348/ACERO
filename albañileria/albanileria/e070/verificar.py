"""Orquestador: verificar_edificio(edificio) -> ResultadoEdificio.

Independiente de la interfaz y de ETABS.
"""
from __future__ import annotations

from ..e030 import norma, sismo
from ..models import (Check, Edificio, Muro, ResultadoEdificio, ResultadoMuro, ResultadoPisoDir, Step,
                      check_fuera_de_alcance, fmt, fmt_kN, fmt_kNm)
from . import confinamiento, corte, materiales, perpendicular, requisitos, rigidez, severo

FUERA_DE_ALCANCE = [
    "Muros de albañilería armada y muros no confinados (E.070 Cap. 8 y Art. 28): no se diseñan.",
    "Contribución de las alas (muros transversales) en la rigidez y en la resistencia: no se considera en modo manual.",
    "Placas de concreto: se listan pero no se verifican (programa de concreto E.060).",
    "Análisis dinámico modal espectral en modo manual (sólo estático E.030; el dinámico se toma de ETABS).",
    "Cimentación, losas y escaleras.",
    "Interacción tabique-pórtico, alféizares y columnas cortas: sólo avisos y recomendaciones.",
    "Verificación por flexocompresión de muros con alas y longitud de anclaje detallada del refuerzo.",
]


def cargas_gravedad(ed: Edificio, m: Muro) -> tuple[float, float, list[Step], list[str]]:
    """Pg = D + %L (sobrecarga reducida) y Pm = D + L (N)."""
    av: list[str] = []
    pct = norma.porcentaje_carga_viva(ed.categoria)
    if m.Pg is not None:
        Pg = m.Pg
        if m.Pm is not None:
            Pm = m.Pm
        else:
            Pm = Pg
            av.append(f"{m.nombre} (piso {m.piso}): Pm (D + L) no ingresado; el esfuerzo axial se calcula con Pg.")
        return Pg, Pm, [Step("Pg", fmt_kN(Pg), "Carga de gravedad D + %L (dato)"),
                        Step("Pm", fmt_kN(Pm), "Carga de gravedad D + L (dato)")], av
    niveles = max(ed.N - m.piso + 1, 1)
    peso_muro = ed.material.gamma * m.t * m.L * m.h * niveles
    pasos = [Step("Pp", fmt_kN(peso_muro), f"Peso propio de los muros de los {niveles} niveles superiores (incluido)",
                  "Pp = γ·t·L·h·n", f"{ed.material.gamma * 1e6:.1f}e-6·{m.t:.0f}·{m.L:.0f}·{m.h:.0f}·{niveles}")]
    if not m.area_trib:
        av.append(f"{m.nombre} (piso {m.piso}): sin Pg ni área tributaria; se usa sólo el peso propio.")
        return peso_muro, peso_muro, pasos, av
    D = L = 0.0
    for j in range(m.piso, ed.N + 1):
        az = j == ed.N
        wd = (ed.cargas.peso_losa + ed.cargas.acabados) if az else ed.cargas.muerta
        wl = ed.cargas.sobrecarga_azotea if az else ed.cargas.sobrecarga
        D += m.area_trib * wd
        L += m.area_trib * wl
    Pg = D + peso_muro + pct * L
    Pm = D + peso_muro + L
    pasos += [
        Step("D losa", fmt_kN(D), f"Carga muerta de losa (área tributaria {m.area_trib / 1e6:.2f} m² por nivel)",
             "D = At·Σwd", ""),
        Step("L", fmt_kN(L), "Sobrecarga total", "L = At·Σwl", ""),
        Step("Pg", fmt_kN(Pg), f"Sobrecarga reducida al {pct * 100:.0f} %", "Pg = D + Pp + %·L",
             f"{D / 1000:.2f} + {peso_muro / 1000:.2f} + {pct}·{L / 1000:.2f}"),
        Step("Pm", fmt_kN(Pm), "Con toda la sobrecarga", "Pm = D + Pp + L",
             f"{D / 1000:.2f} + {peso_muro / 1000:.2f} + {L / 1000:.2f}"),
    ]
    return Pg, Pm, pasos, av


def _fuerzas_sismo(ed: Edificio, sis: dict, res: ResultadoEdificio) -> dict:
    """Ve, Me (sismo moderado) por muro: de ETABS o repartidos por rigidez."""
    out: dict = {}
    op = ed.opciones
    for d in ("X", "Y"):
        s = sis[d]
        for i in range(1, ed.N + 1):
            muros = [m for m in ed.muros_piso(i, d) if m.portante]
            if not muros:
                continue
            Vmod = s.Ve[i - 1]
            usar_dato = op.modo_etabs or all(m.Ve is not None and m.Me is not None for m in muros)
            if usar_dato:
                for m in muros:
                    Ve = abs(m.Ve or 0.0)
                    Me = abs(m.Me or 0.0)
                    out[(m.nombre, i)] = {"Ve": Ve, "Me": Me, "origen": "ETABS / dato", "pasos": [
                        Step("Ve", fmt_kN(Ve), "Cortante del muro, sismo moderado (dato del modelo)"),
                        Step("Me", fmt_kNm(Me), "Momento del muro, sismo moderado (dato del modelo)")]}
                continue
            rep = rigidez.repartir_cortante(ed, i, d, Vmod)
            for m in muros:
                r = rep[m.nombre]
                Ve = r["V"]
                pasos = rigidez.pasos_rigidez(m, ed.material.Em_calc, ed.material.Gm_calc, op.rigidez, op.factor_forma)
                pasos += [
                    Step("ΣK", f"{r['sumK']:.1f} N/mm", f"Suma de rigideces, piso {i}, dirección {d}"),
                    Step("Ve dir", fmt_kN(r["Vdir"]), "Cortante por rigidez (sismo moderado)", "Ve = Vpiso·K/ΣK",
                         f"{Vmod / 1000:.2f}·{r['K']:.1f}/{r['sumK']:.1f}"),
                ]
                if op.torsion:
                    pasos.append(Step("ΔVt", fmt_kN(r["Vtor"]), "Incremento por torsión (exc. real + accidental 0.05·B)",
                                      "ΔV = Mt·K·d/J", ""))
                if op.modo_momento == "voladizo":
                    H = s.alturas
                    Hb = H[i - 2] if i > 1 else 0.0
                    Mpiso = sum(s.F[j] * op.factor_moderado * (H[j] - Hb) for j in range(i - 1, ed.N))
                    Me = Mpiso * Ve / Vmod if Vmod else 0.0
                    pasos.append(Step("Me", fmt_kNm(Me), "Momento en voladizo acumulado (proporcional a Ve/Vpiso)",
                                      "Me = (Ve/Vpiso)·ΣFj·(Hj - Hi-1)", f"({Ve / 1000:.2f}/{Vmod / 1000:.2f})·{Mpiso / 1e6:.2f}"))
                elif op.modo_momento == "factor":
                    Me = Ve * op.factor_brazo * m.h
                    pasos.append(Step("Me", fmt_kNm(Me), "Momento con brazo ingresado", "Me = Ve·β·h",
                                      f"{Ve / 1000:.2f}·{op.factor_brazo}·{m.h / 1000:.2f}"))
                else:
                    Me = Ve * m.h
                    pasos.append(Step("Me", fmt_kNm(Me), "Momento con brazo igual a la altura", "Me = Ve·h",
                                      f"{Ve / 1000:.2f}·{m.h / 1000:.2f}"))
                out[(m.nombre, i)] = {"Ve": Ve, "Me": Me, "K": r["K"], "origen": "reparto por rigidez", "pasos": pasos}
    return out


def verificar_edificio(ed: Edificio) -> ResultadoEdificio:
    res = ResultadoEdificio(nombre=ed.memoria.proyecto or ed.nombre)
    res.fuera_de_alcance = list(FUERA_DE_ALCANCE)
    mat = ed.material
    res.checks_globales += materiales.checks_material(ed)

    sis = {d: sismo.analisis_estatico(ed, d) for d in ("X", "Y")}
    res.sismo = sis
    fuerzas = _fuerzas_sismo(ed, sis, res)

    # ---------------------------------------------------------------- muros: Art. 19 y 26
    datos: dict = {}
    for m in sorted(ed.muros, key=lambda w: (w.piso, w.direccion, w.nombre)):
        rm = ResultadoMuro(nombre=m.nombre, piso=m.piso, direccion=m.direccion)
        rm.advertencias += requisitos.avisos_muro(m, ed.N)
        res.muros.append(rm)
        if not m.portante or not m.confinado:
            rm.fuera_de_alcance = True
            rm.checks.append(check_fuera_de_alcance(
                "Verificación en el plano", "E.070",
                "muro no portante o no confinado: sólo se verifica ante carga perpendicular (Art. 29)"
                if not m.portante else "muro portante no confinado (albañilería simple/armada)"))
            if m.carga_perpendicular:
                Pg, Pm, _, _ = cargas_gravedad(ed, m)
                rm.checks.append(perpendicular.verificar_portante_perpendicular(m, ed, Pg))
            continue
        Pg, Pm, pasos_g, av = cargas_gravedad(ed, m)
        rm.advertencias += av
        f = fuerzas.get((m.nombre, m.piso), {"Ve": 0.0, "Me": 0.0, "pasos": []})
        Ve, Me = f["Ve"], f["Me"]
        rm.checks.append(requisitos.check_espesor(m, ed.zona))
        ca = requisitos.check_esfuerzo_axial(m, Pm, mat.fm)
        ca.pasos = pasos_g + ca.pasos
        rm.checks.append(ca)
        rm.checks.append(requisitos.check_separacion_columnas(m))
        Vm, alfa, pasos_vm = corte.resistencia_corte(m, mat, Ve, Me, Pg)
        cf = corte.check_fisuracion(Ve, Vm, f["pasos"] + pasos_vm)
        rm.checks.append(cf)
        sigma = Pm / (m.L * m.t)
        rm.valores.update({"Pg": Pg, "Pm": Pm, "sigma": sigma, "Ve": Ve, "Me": Me, "alfa": alfa, "Vm": Vm,
                           "Vadm": 0.55 * Vm, "dcr_fis": cf.dcr, "K": f.get("K"), "origen": f.get("origen", "")})
        datos[(m.nombre, m.piso)] = (m, rm)
        if m.carga_perpendicular:
            rm.checks.append(perpendicular.verificar_portante_perpendicular(m, ed, Pg))

    # ---------------------------------------------------------------- por piso y dirección
    for i in range(1, ed.N + 1):
        for d in ("X", "Y"):
            pd = ResultadoPisoDir(piso=i, direccion=d)
            dens, req, cd = requisitos.densidad_muros(ed, i, d)
            pd.densidad_existente, pd.densidad_requerida = dens, req
            pd.checks.append(cd)
            pd.VE = sis[d].VE[i - 1]
            pd.Ve_total = sis[d].Ve[i - 1]
            items = [(m, rm) for (n, p), (m, rm) in datos.items() if p == i and m.direccion == d]
            pd.muros = [m.nombre for m, _ in items]
            pd.suma_Vm = sum(rm.valores["Vm"] for _, rm in items)
            pd.relacion = pd.suma_Vm / pd.VE if pd.VE > 0 else 0.0
            pd.elastico = pd.relacion >= 3.0
            if items:
                pd.checks.append(corte.check_resistencia_global(i, d, pd.suma_Vm, pd.VE, [m.nombre for m, _ in items],
                                                                [rm.valores["Vm"] for _, rm in items]))
            res.pisos.append(pd)

    # ---------------------------------------------------------------- sismo severo (Art. 27)
    for (nombre, piso), (m, rm) in datos.items():
        v = rm.valores
        base = (m, rm)
        if m.piso > 1:
            base = next(((mm, rr) for (n2, p2), (mm, rr) in datos.items()
                         if p2 == 1 and mm.id_grupo == m.id_grupo and mm.direccion == m.direccion), None)
        if base is not None:
            fac, fac0 = severo.factor_amplificacion(base[1].valores["Vm"], base[1].valores["Ve"])
            origen = f"muro {base[0].nombre} del piso 1"
        else:
            pd1 = res.piso_dir(1, m.direccion)
            fac, fac0 = severo.factor_amplificacion(pd1.suma_Vm if pd1 else 0.0, pd1.Ve_total if pd1 else 0.0)
            origen = "sin muro equivalente en el piso 1: se usa ΣVm1/ΣVe1 del entrepiso"
            rm.advertencias.append(f"{m.nombre} (piso {m.piso}): no tiene muro con el mismo grupo en el piso 1; "
                                   f"factor Vm1/Ve1 del entrepiso = {fac:.2f}.")
        Vu, Mu, agr, cs = severo.diseno_severo(m.piso, v["Ve"], v["Me"], v["Vm"], fac, fac0, origen)
        rm.checks.append(cs)
        v.update({"factor": fac, "factor_bruto": fac0, "Vu": Vu, "Mu": Mu, "agrietado": agr})
        req_h, ch = severo.refuerzo_horizontal(m.nombre, m.piso, ed.N, Vu, v["Vm"], v["sigma"], mat.fm, m.t)
        v["refuerzo_horizontal"] = req_h
        rm.checks.append(ch)
        pd = res.piso_dir(m.piso, m.direccion)
        if pd.elastico:
            tipo = "mínimo"
        elif agr:
            tipo = "agrietado"
        else:
            tipo = "no agrietado"
        v["tipo_diseno"] = tipo
        rc = confinamiento.diseno_confinamiento(m, mat, ed.opciones, v["Pg"], v["Vm"], Mu, Vu, tipo)
        rm.confinamiento = rc
        rm.checks += rc.checks
        if agr and not pd.elastico:
            pd.muros_agrietados.append(m.nombre)

    # ---------------------------------------------------------------- derivas
    for pd in res.pisos:
        s = sis[pd.direccion]
        muros = [m for m in ed.muros_piso(pd.piso, pd.direccion) if m.portante and m.confinado]
        h = ed.alturas_entrepiso[pd.piso - 1]
        if ed.derivas and ed.derivas.get(pd.direccion):
            de = ed.derivas[pd.direccion][pd.piso - 1]
            der = 0.75 * s.R * de
            pasos = [Step("Δe/h", fmt(de, 5), "Deriva elástica del modelo (ETABS)"),
                     Step("Δ/h", fmt(der, 5), "Deriva inelástica (estructura regular)", "Δ/h = 0.75·R·Δe/h",
                          f"0.75·{s.R:.2f}·{de:.5f}")]
        elif muros:
            sumK = sum(rigidez.rigidez_lateral(m, mat.Em_calc, mat.Gm_calc, ed.opciones.rigidez, ed.opciones.factor_forma)
                       for m in muros)
            de = pd.VE / sumK / h
            der = 0.75 * s.R * de
            pasos = [Step("ΣK", f"{sumK:.1f} N/mm", "Rigidez del entrepiso (muros en voladizo por piso)"),
                     Step("δe", f"{pd.VE / sumK:.3f} mm", "Desplazamiento relativo elástico", "δe = VE/ΣK",
                          f"{pd.VE:.0f}/{sumK:.1f}"),
                     Step("Δ/h", fmt(der, 5), "Deriva inelástica (estructura regular)", "Δ/h = 0.75·R·δe/h",
                          f"0.75·{s.R:.2f}·{pd.VE / sumK:.3f}/{h:.0f}")]
        else:
            continue
        pd.deriva = der
        c = Check(f"Deriva de entrepiso (piso {pd.piso}, dirección {pd.direccion})", "E.030 Art. 32 (Tabla N°11)",
                  der, norma.DERIVA_MAX, "", pasos=pasos)
        if not (ed.derivas and ed.derivas.get(pd.direccion)):
            c.notas.append("Aproximación: rigidez por piso de muros en voladizo (no incluye la flexión global "
                           "del edificio). Confirmar con el modelo de ETABS.")
        pd.checks.append(c)

    # ---------------------------------------------------------------- listas de no cumplimiento
    for pd in res.pisos:
        for rm in res.muros:
            if rm.piso == pd.piso and rm.direccion == pd.direccion and rm.estado == "NO CUMPLE":
                pd.muros_no_cumplen.append(rm.nombre)

    # ---------------------------------------------------------------- no portantes
    for el in ed.no_portantes:
        res.no_portantes.append(perpendicular.verificar_no_portante(el, ed))

    # ---------------------------------------------------------------- avisos generales
    for rm in res.muros:
        res.advertencias += rm.advertencias
    if any(m.tabique for m in ed.muros):
        res.advertencias.append("Hay tabiques o alféizares no aislados: la interacción tabique-pórtico y la columna "
                                "corta de albañilería no se modelan (sólo avisos). Aislar con junta ≥ 1 cm.")
    if ed.opciones.torsion and not (ed.Lx and ed.Ly):
        res.advertencias.append("Torsión: no se ingresaron Lx y Ly; se usó B = √Ap para la excentricidad accidental.")
    return res
