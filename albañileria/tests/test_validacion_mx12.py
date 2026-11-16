"""Validación externa: memoria de cálculo del muro MX12, piso 1 (hoja "Análisis y diseño de elementos de
albañilería confinada - E.070" entregada por el usuario). Unidades de la memoria: tonf, m, kgf/cm².

Datos: f'm = 65, v'm = 8.05, E = 32500 kgf/cm²; L = 2.60 m, t = 0.23 m, h = 2.40 m; Pm = 16 t, Pg = 14 t,
Ve = 8.72 t, Me = 42.98 t·m; f'c = 210, fy = 4200 kgf/cm²; μ = 1.0; δ = 1.0; φ = 0.7; columna 25 x 23 cm.
"""
import pytest

from albanileria.e070 import verificar_edificio
from albanileria.e070.confinamiento import BARRAS
from albanileria.models import Edificio, Material, Muro, mpa

TF = 9806.65            # N por tonelada-fuerza

# (magnitud, valor de la memoria, función que extrae el valor del programa en las unidades de la memoria)
REFERENCIA = {
    "σm (t/m²)": 26.76, "α": 0.53, "Vm (t)": 15.92, "0.55·Vm (t)": 8.75, "Vm1/Ve1": 2.00, "Vu (t)": 17.44,
    "Mu (t·m)": 85.96, "M (t·m)": 66.85, "F (t)": 25.71, "Pc (t)": 7.00, "Vc (t)": 7.96, "T (t)": 18.713,
    "C (t)": 32.71, "Acf (cm²)": 222.93, "As = Asf + Ast (cm²)": 7.47, "Ts (t)": 7.96, "As solera Ts/φfy (cm²)": 2.11,
    "Ac (cm²)": 575.0, "An núcleo (cm²)": 399.0, "s1 con Ø1/4\" (cm)": 5.09, "s2 con Ø1/4\" (cm)": 5.61,
}


def resolver():
    mat = Material(fm=mpa(65), vm=mpa(8.05), Em=mpa(32500), fc=mpa(210), fy=mpa(4200))
    m = Muro("MX12", 1, "X", L=2600, t=230, h=2400, Pg=14 * TF, Pm=16 * TF, Ve=8.72 * TF, Me=42.98 * TF * 1000,
             Nc=2, junta_rugosa=True, muro_transversal=True)
    ed = Edificio(N=1, Ap=50e6, material=mat, muros=[m], VE_override={"X": [100 * TF], "Y": [100 * TF]})
    ed.opciones.modo_etabs = True
    r = verificar_edificio(ed)
    rm = r.muros[0]
    v, rc = rm.valores, rm.confinamiento
    # estribos con Ø1/4" (Av = 2·0.32 cm²) como en la memoria: s1, s2
    Av, tn = 2 * BARRAS["1/4\""]["area"], 230 - 40
    Ac, An = rc.b_col * rc.d_col, (rc.b_col - 40) * (rc.d_col - 40)
    s1 = Av * mat.fy / (0.3 * tn * mat.fc * (Ac / An - 1))
    s2 = Av * mat.fy / (0.12 * tn * mat.fc)
    prog = {
        "σm (t/m²)": v["sigma"] * 1e6 / TF, "α": v["alfa"], "Vm (t)": v["Vm"] / TF, "0.55·Vm (t)": v["Vadm"] / TF,
        "Vm1/Ve1": v["factor"], "Vu (t)": v["Vu"] / TF, "Mu (t·m)": v["Mu"] / TF / 1000, "M (t·m)": rc.M / TF / 1000,
        "F (t)": rc.F / TF, "Pc (t)": rc.Pc_ext / TF, "Vc (t)": rc.Vc_ext / TF, "T (t)": rc.T_ext / TF,
        "C (t)": rc.C_ext / TF, "Acf (cm²)": rc.Acf / 100, "As = Asf + Ast (cm²)": rc.As_req_ext / 100,
        "Ts (t)": rc.Ts / TF, "As solera Ts/φfy (cm²)": rc.Ts / (0.9 * mat.fy) / 100, "Ac (cm²)": Ac / 100,
        "An núcleo (cm²)": An / 100, "s1 con Ø1/4\" (cm)": s1 / 10, "s2 con Ø1/4\" (cm)": s2 / 10,
    }
    return r, rm, prog


def tabla_diferencias():
    _, _, prog = resolver()
    return [(k, ref, prog[k], (prog[k] - ref) / ref * 100) for k, ref in REFERENCIA.items()]


@pytest.mark.parametrize("magnitud", list(REFERENCIA))
def test_coincide_con_memoria(magnitud):
    _, _, prog = resolver()
    ref = REFERENCIA[magnitud]
    tol = 0.006 if magnitud == "α" else 0.0015 if magnitud in ("0.55·Vm (t)", "M (t·m)") else 0.001
    assert prog[magnitud] == pytest.approx(ref, rel=max(tol, 0.005 / ref))


def test_estados_y_decisiones():
    _, rm, _ = resolver()
    v, rc = rm.valores, rm.confinamiento
    fis = next(c for c in rm.checks if c.nombre.startswith("Control de fisuración"))
    assert fis.cumple                              # 8.72 ≤ 8.75 t: "Cumple - No hay fisuramiento"
    assert v["agrietado"] and v["tipo_diseno"] == "agrietado"
    assert v["refuerzo_horizontal"]                # Vu = 17.44 > Vm = 15.92: "necesita refuerzo horizontal"
    assert rc.seccion_ext == "23 x 25 cm"          # memoria: 25 x 23 cm (15·t = 345 cm² gobierna sobre Acf)
    assert rc.As_prov_ext >= rc.As_req_ext         # la memoria adopta 10 Ø 3/8" = 7.10 < 7.47 cm² requeridos
    assert rc.refuerzo_solera == "4 Ø 3/8\""       # igual que la memoria


if __name__ == "__main__":
    print(f"{'Magnitud':28s} {'Memoria':>10s} {'Programa':>10s} {'Dif %':>8s}")
    for k, ref, p, d in tabla_diferencias():
        print(f"{k:28s} {ref:10.3f} {p:10.3f} {d:8.2f}")
