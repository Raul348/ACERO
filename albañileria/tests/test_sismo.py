"""Pruebas del análisis sísmico E.030."""
import pytest

from albanileria.e030 import norma, sismo
from albanileria.models import Edificio


def test_tablas():
    assert norma.factor_Z(4) == 0.45 and norma.factor_Z(1) == 0.10
    assert norma.factor_S(4, "S2") == 1.05 and norma.factor_S(1, "S3") == 2.00
    assert norma.periodos_suelo("S3") == (1.0, 1.6)
    assert norma.factor_U("B") == 1.3


def test_factor_C():
    assert norma.factor_C(0.2, 0.4, 2.5) == 2.5
    assert norma.factor_C(1.0, 0.4, 2.5) == pytest.approx(1.0)       # 2.5·0.4/1
    assert norma.factor_C(3.0, 0.4, 2.5) == pytest.approx(2.5 * 0.4 * 2.5 / 9)


def test_cortante_y_distribucion():
    ed = Edificio(zona=4, categoria="C", suelo="S1", N=3, h_entrepiso=2600)
    ed.pesos_override = [1000e3, 1000e3, 1000e3]
    r = sismo.analisis_estatico(ed, "X")
    assert r.T == pytest.approx(7.8 / 60)
    assert r.C == 2.5 and r.R == 3
    assert r.V == pytest.approx(0.45 * 1 * 2.5 * 1 * 3000e3 / 3)      # 1125 kN
    assert r.F == pytest.approx([r.V / 6, 2 * r.V / 6, 3 * r.V / 6])
    assert r.VE == pytest.approx([r.V, 5 * r.V / 6, 3 * r.V / 6])
    assert r.Ve == pytest.approx([v / 2 for v in r.VE])


def test_C_sobre_R_minimo():
    ed = Edificio(zona=4, suelo="S1", N=1, h_entrepiso=60_000)       # T = 1 s, C = 1.0
    ed.Ia = 0.5                                                        # R = 1.5 → C/R = 0.67
    ed.pesos_override = [1e6]
    r = sismo.analisis_estatico(ed)
    assert r.R == 1.5
    ed.Ia, ed.h_entrepiso = 1.0, 300_000                               # T = 5 s, C = 0.1 → C/R = 0.033
    r = sismo.analisis_estatico(ed)
    assert r.V == pytest.approx(0.45 * 1 * 0.11 * 1 * 1e6)
