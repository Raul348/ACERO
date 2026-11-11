"""Muros no portantes ante carga perpendicular (E.070 Art. 29, Tabla 12)."""
import pytest

from albanileria.e070 import perpendicular as pp
from albanileria.models import Edificio, ElementoNoPortante


def test_coeficientes_tabla12():
    assert pp.coeficiente_m(1, 2400, 2400)[0] == pytest.approx(0.0479)
    assert pp.coeficiente_m(1, 2400, 3000)[0] == pytest.approx(0.0659)      # b/a = 1.25
    assert pp.coeficiente_m(2, 2000, 1000)[0] == pytest.approx(0.060)       # b/a = 0.5
    assert pp.coeficiente_m(3, 2400, 3000)[0] == 0.125
    assert pp.coeficiente_m(4, 1000, 3000)[0] == 0.5
    assert pp.coeficiente_m(1, 1000, 50_000)[0] == pytest.approx(0.125, rel=5e-3)     # b/a → ∞
    assert pp.coeficiente_m(1, 1000, 3000)[0] == pytest.approx(0.1180)


def test_carga_w():
    # w = 0.8·0.45·1·2·18·0.13 = 1.6848 kN/m²
    w = pp.carga_sismica(0.45, 1.0, 2.0, 18e-6, 130)
    assert w * 1000 == pytest.approx(1.6848)


def test_tabique_4_bordes():
    ed = Edificio(zona=4, categoria="C")
    el = ElementoNoPortante("T", tipo="tabique", caso=1, a=2400, b=3000, t=130)
    r = pp.verificar_no_portante(el, ed)
    # Ms = 0.0659·1.6848·2.4² = 0.63952 kN·m/m ; fm = 6·639.52/130² = 0.22705 MPa
    assert r.valores["Ms"] == pytest.approx(639.52, rel=1e-4)
    assert r.valores["fm"] == pytest.approx(0.22705, rel=1e-4)
    assert r.estado == "NO CUMPLE"
    assert r.valores["sep_arriostre"] > 0


def test_parapeto_cumple_bajo():
    ed = Edificio(zona=2, categoria="C")
    el = ElementoNoPortante("P", tipo="parapeto", caso=4, a=400, b=3000, t=130)
    r = pp.verificar_no_portante(el, ed)
    w = 0.8 * 0.25 * 1 * 3.0 * 18 * 0.13          # kN/m²
    Ms = 0.5 * w * 0.4 ** 2                        # kN·m/m
    assert r.valores["fm"] == pytest.approx(6 * Ms * 1000 / 130 ** 2, rel=1e-6)
    assert r.estado == "CUMPLE"
