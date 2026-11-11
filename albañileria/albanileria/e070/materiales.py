"""Materiales de la albañilería y del confinamiento (E.070 Cap. 5, Tabla 9, Art. 24.7)."""
from __future__ import annotations

from ..models import Check, Edificio, Material, Step, cargar_json, fmt_MPa, mpa

FC_MIN = mpa(175)          # 17.16 MPa


def unidades_tabla9() -> list[dict]:
    return cargar_json("tabla9.json")["unidades"]


def pasos_material(mat: Material) -> list[Step]:
    p = [
        Step("Unidad", mat.tipo_unidad, f"Familia: {mat.familia}; mortero {mat.mortero}"),
        Step("f'b", fmt_MPa(mat.fb), "Resistencia característica a compresión de la unidad", "f'b", "E.070 Tabla 9 / ensayo"),
        Step("f'm", fmt_MPa(mat.fm), "Resistencia característica a compresión axial de pilas", "f'm", "E.070 Tabla 9 / ensayo de prismas"),
        Step("v'm", fmt_MPa(mat.vm), "Resistencia característica a corte de muretes", "v'm", "E.070 Tabla 9 / ensayo de muretes"),
    ]
    if mat.Em:
        p.append(Step("Em", fmt_MPa(mat.Em, 0), "Módulo de elasticidad (valor ingresado)", "Em", ""))
    else:
        p.append(Step("Em", fmt_MPa(mat.Em_calc, 0), "Módulo de elasticidad (E.070 Art. 24.7)",
                      f"Em = {mat.coef_Em:.0f}·f'm", f"{mat.coef_Em:.0f}·{mat.fm:.3f}"))
    if mat.Gm:
        p.append(Step("Gm", fmt_MPa(mat.Gm, 0), "Módulo de corte (valor ingresado)", "Gm", ""))
    else:
        p.append(Step("Gm", fmt_MPa(mat.Gm_calc, 0), "Módulo de corte (E.070 Art. 24.8)", "Gm = 0.4·Em",
                      f"0.4·{mat.Em_calc:.0f}"))
    p += [
        Step("f'c", fmt_MPa(mat.fc), "Concreto de confinamiento (≥ 17.15 MPa = 175 kgf/cm²)", "f'c", ""),
        Step("fy", fmt_MPa(mat.fy, 0), "Acero de refuerzo grado 60", "fy", ""),
        Step("f't", fmt_MPa(mat.ft), "Resistencia a tracción por flexión (cargas perpendiculares)", "f't", "E.070 Art. 29"),
        Step("γ", f"{mat.gamma * 1e6:.1f} kN/m³", "Peso específico de la albañilería", "γ", "E.020"),
    ]
    return p


def checks_material(ed: Edificio) -> list[Check]:
    mat = ed.material
    cs = []
    c = Check("Resistencia mínima del concreto de confinamiento", "E.070 Art. 11", FC_MIN, mat.fc, "MPa",
              pasos=[Step("f'c mín", fmt_MPa(FC_MIN), "", "f'c ≥ 17.15 MPa", "175 kgf/cm²"),
                     Step("f'c", fmt_MPa(mat.fc), "Concreto especificado")])
    cs.append(c)
    if mat.mortero == "NP":
        cs.append(Check("Tipo de mortero en muros portantes", "E.070 Art. 6.4 (Tabla 4)", 1.0, 0.0, "",
                        notas=["El mortero NP no debe usarse en muros portantes; use P1 o P2."]))
    if ed.N > 3 and mat.tipo_unidad.lower().startswith("king kong artesanal"):
        cs.append(Check("Unidades artesanales en edificios de más de 3 pisos", "E.070 Art. 5.3 (Tabla 2)", 1.0, 0.0, "",
                        notas=["Las unidades artesanales no se recomiendan como muros portantes en zonas sísmicas 2, 3 y 4 "
                               "para edificios de más de 3 pisos (limitaciones de la Tabla 2). Verifique."]))
    return cs
