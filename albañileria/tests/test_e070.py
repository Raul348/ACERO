"""Pruebas del motor E.070 con números calculados a mano."""
import math

import pytest

from albanileria.e070 import corte, materiales, requisitos, rigidez, severo, verificar_edificio
from albanileria.ejemplos import edificio_ejemplo
from albanileria.models import Check, Edificio, Material, Muro, mpa

FM = 65 / 10.197          # MPa
VM = 8.1 / 10.197


def test_conversion_kgf():
    assert mpa(10.197) == pytest.approx(1.0)
    assert FM == pytest.approx(6.3744, rel=1e-4)


def test_check_dcr():
    assert Check("a", "r", 5, 10, "kN").dcr == 0.5
    assert Check("a", "r", 5, 0, "kN").dcr == math.inf
    assert Check("a", "r", 0, 0, "kN").dcr == 0.0
    c = Check("a", "r", 50, 10, "kN", aplica=False)
    assert c.dcr == 0 and c.cumple


def test_modulos():
    m = Material.desde_tabla9("King Kong industrial (arcilla)")
    assert m.fm == pytest.approx(FM)
    assert m.Em_calc == pytest.approx(500 * FM)          # 3187.2 MPa
    assert m.Gm_calc == pytest.approx(0.4 * 500 * FM)
    c = Material(familia="concreto", fm=mpa(73))
    assert c.Em_calc == pytest.approx(700 * 73 / 10.197)
    s = Material(familia="silico")
    assert s.coef_vm == 0.35


def test_espesor():
    m = Muro("M", h=2400, t=130)
    c = requisitos.check_espesor(m, 4)
    assert c.demanda == pytest.approx(120.0) and c.cumple       # 2400/20
    c1 = requisitos.check_espesor(m, 1)
    assert c1.demanda == pytest.approx(96.0)                    # 2400/25
    assert not requisitos.check_espesor(Muro("M", h=2800, t=130), 4).cumple   # 140 > 130


def test_esfuerzo_axial():
    m = Muro("M", L=3000, t=130, h=2400)
    c = requisitos.check_esfuerzo_axial(m, 200_000, FM)
    assert c.demanda == pytest.approx(200_000 / 390_000)         # 0.5128 MPa
    fa1 = 0.2 * FM * (1 - (2400 / (35 * 130)) ** 2)               # 0.9201 MPa
    assert c.capacidad == pytest.approx(min(fa1, 0.15 * FM))
    assert c.capacidad == pytest.approx(0.9201, rel=1e-3)


def test_densidad():
    ed = Edificio(zona=4, categoria="C", suelo="S1", N=3, Ap=60e6)
    ed.muros = [Muro("A", L=3000), Muro("B", L=4000), Muro("C", L=5000), Muro("D", L=1000)]
    dens, req, c = requisitos.densidad_muros(ed, 1, "X")
    assert dens == pytest.approx(12000 * 130 / 60e6)    # 0.026 (D no cuenta: L < 1.20 m)
    assert req == pytest.approx(0.45 * 1 * 1 * 3 / 56)  # 0.024107
    assert c.dcr == pytest.approx(0.024107 / 0.026, rel=1e-4)


def test_alfa():
    assert corte.factor_alfa(50_000, 150e6, 3000)[0] == pytest.approx(1.0)
    assert corte.factor_alfa(50_000, 300e6, 3000)[0] == pytest.approx(0.5)
    a, a0 = corte.factor_alfa(50_000, 600e6, 3000)
    assert a == pytest.approx(1 / 3) and a0 == pytest.approx(0.25)
    assert corte.factor_alfa(50_000, 50e6, 3000)[0] == 1.0      # 3.0 → 1


def test_vm_y_fisuracion():
    m = Muro("M", L=3000, t=130, h=2400)
    mat = Material()
    Vm, alfa, _ = corte.resistencia_corte(m, mat, 50_000, 150e6, 100_000)
    # Vm = 0.5·0.79435·1·130·3000 + 0.23·100000 = 154898.5 + 23000
    assert Vm == pytest.approx(177_898.5, rel=1e-5)
    c = corte.check_fisuracion(50_000, Vm, [])
    assert c.capacidad == pytest.approx(0.55 * 177_898.5, rel=1e-5)    # 97 844 N
    assert c.dcr == pytest.approx(50_000 / 97_844.2, rel=1e-4)
    mat_s = Material(familia="silico", vm=VM)
    Vs, _, _ = corte.resistencia_corte(m, mat_s, 50_000, 150e6, 100_000)
    assert Vs == pytest.approx(0.35 * VM * 390_000 + 23_000)


def test_factor_amplificacion():
    assert severo.factor_amplificacion(200, 60)[0] == pytest.approx(3.0)
    assert severo.factor_amplificacion(90, 60)[0] == pytest.approx(2.0)
    assert severo.factor_amplificacion(150, 60)[0] == pytest.approx(2.5)
    Vu, Mu, agr, c = severo.diseno_severo(2, 40_000, 100e6, 120_000, 2.5, 2.5, "")
    assert Vu == pytest.approx(100_000) and Mu == pytest.approx(250e6) and not agr
    Vu, _, agr, _ = severo.diseno_severo(2, 50_000, 100e6, 120_000, 2.5, 2.5, "")
    assert agr                                              # 125 kN ≥ 120 kN


def test_rigidez():
    m = Muro("M", L=3000, t=130, h=2400)
    Em, Gm = 500 * FM, 0.4 * 500 * FM
    I = 130 * 3000 ** 3 / 12
    k = 1 / (2400 ** 3 / (3 * Em * I) + 1.2 * 2400 / (Gm * 130 * 3000))
    assert rigidez.rigidez_lateral(m, Em, Gm) == pytest.approx(k)
    assert k == pytest.approx(93_160, rel=1e-3)
    k12 = 1 / (2400 ** 3 / (12 * Em * I) + 1.2 * 2400 / (Gm * 130 * 3000))
    assert rigidez.rigidez_lateral(m, Em, Gm, "empotrado") == pytest.approx(k12)


def test_reparto_por_rigidez():
    ed = Edificio(N=1)
    ed.muros = [Muro("A", L=3000), Muro("B", L=3000), Muro("C", L=6000)]
    rep = rigidez.repartir_cortante(ed, 1, "X", 100_000)
    assert rep["A"]["V"] == pytest.approx(rep["B"]["V"])
    assert sum(r["V"] for r in rep.values()) == pytest.approx(100_000)
    assert rep["C"]["V"] > 2 * rep["A"]["V"]               # muro largo: más del doble


def test_edificio_ejemplo_completo():
    ed = edificio_ejemplo()
    r = verificar_edificio(ed)
    assert len(r.muros) == 36
    sx = r.sismo["X"]
    assert sx.C == 2.5 and sx.T == pytest.approx(7.8 / 60)
    assert sx.V == pytest.approx(0.45 * 1 * 2.5 * 1 * sx.P / 3)
    # cortante moderado = mitad del severo
    assert sx.Ve[0] == pytest.approx(0.5 * sx.VE[0])
    # la suma de Ve de los muros de un piso = Ve del piso (sin torsión)
    ve1 = sum(m.valores["Ve"] for m in r.muros if m.piso == 1 and m.direccion == "X")
    assert ve1 == pytest.approx(sx.Ve[0])
    for pd in r.pisos:
        assert pd.densidad_existente > pd.densidad_requerida
        assert pd.suma_Vm > pd.VE
    # muros del primer piso: diseño de columnas por Tabla 11
    assert all(m.valores["tipo_diseno"] == "agrietado" for m in r.muros if m.piso == 1)
    assert all(m.estado == "CUMPLE" for m in r.muros)


def test_fuera_de_alcance_no_portante():
    ed = Edificio(N=1, Ap=20e6)
    ed.muros = [Muro("A", L=3000), Muro("B", L=3000, direccion="Y"), Muro("T", L=2000, portante=False)]
    r = verificar_edificio(ed)
    t = r.muro("T")
    assert t.estado == "FUERA DE ALCANCE"
    assert any("FUERA DE ALCANCE" in n for c in t.checks for n in c.notas)
    assert r.fuera_de_alcance


def test_pg_por_area_tributaria():
    from albanileria.e070.verificar import cargas_gravedad
    ed = Edificio(N=2)
    m = Muro("A", piso=1, L=3000, t=130, h=2400, area_trib=10e6)
    Pg, Pm, _, _ = cargas_gravedad(ed, m)
    pp = 18e-6 * 130 * 3000 * 2400 * 2                          # 33.70 kN
    D = 10e6 * (5.0e-3) + 10e6 * (4.0e-3)                       # piso 1 (5 kN/m²) + azotea (4 kN/m²)
    L = 10e6 * 2.0e-3 + 10e6 * 1.0e-3
    assert Pg == pytest.approx(D + pp + 0.25 * L)
    assert Pm == pytest.approx(D + pp + L)


def test_torsion_aumenta_cortante():
    ed = edificio_ejemplo()
    ed.opciones.torsion = True
    r = verificar_edificio(ed)
    ve1 = sum(m.valores["Ve"] for m in r.muros if m.piso == 1 and m.direccion == "X")
    assert ve1 >= r.sismo["X"].Ve[0] - 1e-6


def test_guardar_abrir(tmp_path):
    ed = edificio_ejemplo()
    p = tmp_path / "p.json"
    ed.guardar(p)
    ed2 = Edificio.abrir(p)
    assert ed2.a_dict() == ed.a_dict()
    assert verificar_edificio(ed2).dcr_max == pytest.approx(verificar_edificio(ed).dcr_max)
