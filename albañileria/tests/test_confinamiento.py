"""Columnas de confinamiento y soleras (E.070 Art. 27, Tabla 11) con números a mano."""
import pytest

from albanileria.e070 import confinamiento as cf
from albanileria.models import Material, Muro, OpcionesAnalisis

FC = 175 / 10.197
FY = 4200 / 10.197


def _muro(**kw):
    d = dict(nombre="M", L=3000, t=130, h=2400, Nc=2)
    d.update(kw)
    return Muro(**d)


def test_barras():
    assert cf.proponer_barras(150)[0] == "4 Ø 8 mm"
    assert cf.proponer_barras(250)[0] == "4 Ø 3/8\""      # 284 mm²
    assert cf.proponer_barras(500)[0] == "4 Ø 1/2\""      # 516 mm²
    assert cf.proponer_barras(700)[0] == "4 Ø 5/8\""      # 796 mm²
    assert cf.proponer_barras(1000)[0] == "6 Ø 5/8\""
    assert not cf.proponer_barras(2000)[2]


def test_tabla11_extrema_un_pano():
    m = _muro()
    rc = cf.diseno_confinamiento(m, Material(), OpcionesAnalisis(), Pg=100_000, Vm=180_000, Mu=375e6,
                                 Vu=180_000, tipo="agrietado")
    # M = 375 - 0.5·180·2.4 = 159 kN·m ; F = 53 kN ; Pc = 50 kN
    assert rc.M == pytest.approx(159e6)
    assert rc.F == pytest.approx(53_000)
    assert rc.Pc_ext == pytest.approx(50_000)
    assert rc.T_ext == pytest.approx(3_000)
    assert rc.C_ext == pytest.approx(103_000)
    assert rc.Vc_ext == pytest.approx(1.5 * 180_000 * 3000 / (3000 * 3))      # 90 kN
    acf = 90_000 / (0.2 * FC * 0.85)
    assert rc.Acf == pytest.approx(acf)                                        # 30 848 mm²
    asf = 90_000 / (FY * 0.8 * 0.85)
    ast = 3_000 / (FY * 0.85)
    As = asf + ast                                                             # 329.9 mm²
    An = As + (103_000 / 0.7 - As * FY) / (0.85 * 0.8 * FC)
    assert rc.An_req == pytest.approx(An)
    assert rc.As_req_ext == pytest.approx(As)
    assert rc.refuerzo_ext == "4 Ø 1/2\""
    # d ≥ máx(250, Acf/t = 237.3, 150, An/(130-40)+40) → 250 mm
    assert rc.d_col == pytest.approx(250)
    # solera: Ts = Vm·Lm/(2L) = 90 kN ; As = 90000/(0.9·fy) = 242.8 mm²
    assert rc.Ts == pytest.approx(90_000)
    assert rc.As_solera == pytest.approx(90_000 / (0.9 * FY))
    assert rc.refuerzo_solera == "4 Ø 3/8\""


def test_tabla11_varios_panos():
    m = _muro(L=6000, Nc=3)
    rc = cf.diseno_confinamiento(m, Material(), OpcionesAnalisis(), Pg=200_000, Vm=300_000, Mu=900e6,
                                 Vu=300_000, tipo="agrietado")
    Lm = 3000.0                                    # paño mayor = 0.5·L
    assert rc.Vc_ext == pytest.approx(1.5 * 300_000 * Lm / (6000 * 4))
    assert rc.Vc_int == pytest.approx(300_000 * Lm / (6000 * 4))
    assert rc.Pc_ext == pytest.approx(200_000 * 3000 / (2 * 6000))        # 50 kN
    assert rc.Pc_int == pytest.approx(200_000 * 3000 / 6000)              # 100 kN
    assert rc.T_int == pytest.approx(300_000 * 2400 / 6000 - 100_000)     # 20 kN
    assert rc.C_int == pytest.approx(100_000 - 0.5 * 300_000 * 2400 / 6000)
    M = 900e6 - 0.5 * 300_000 * 2400
    assert rc.F == pytest.approx(M / 6000)
    assert rc.Ts == pytest.approx(300_000 * Lm / (2 * 6000))


def test_no_agrietado():
    m = _muro()
    rc = cf.diseno_confinamiento(m, Material(), OpcionesAnalisis(), Pg=80_000, Vm=200_000, Mu=300e6,
                                 Vu=120_000, tipo="no agrietado")
    assert rc.F == pytest.approx(100_000)                    # Mu/L
    assert rc.T_ext == pytest.approx(60_000)                 # F - Pc
    assert rc.As_req_ext == pytest.approx(max(60_000 / (0.9 * FY), 200))
    assert rc.Ts == pytest.approx(0.5 * 120_000)             # 0.5·Vu·Lm/L


def test_minimo():
    rc = cf.diseno_confinamiento(_muro(), Material(), OpcionesAnalisis(), 80_000, 200_000, 300e6, 100_000, "mínimo")
    assert rc.refuerzo_ext == "4 Ø 8 mm"
    assert "1 @ 5, 4 @ 10" in rc.estribos
    assert rc.seccion_ext == "13 x 25 cm"
