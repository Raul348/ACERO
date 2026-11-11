"""Lectura de ETABS a partir de un snapshot guardado (sin ETABS abierto)."""
from pathlib import Path

import pytest

from albanileria.e070 import verificar_edificio
from albanileria.ejemplos import edificio_ejemplo
from albanileria.etabs.alternativas import Copiloto
from albanileria.etabs.batch import ConfigEtabs, edificio_desde_snapshot, verificar_snapshot
from albanileria.etabs.clasificar import es_albanileria
from albanileria.etabs.client import filas_tabla
from albanileria.etabs.model import SnapshotEtabs
from albanileria.etabs.reader import parse_pier_forces, parse_pier_props

DATA = Path(__file__).parent / "data" / "snapshot_ejemplo.json"


def test_filas_tabla_orden_comtypes():
    # [FieldKeyList, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret]
    res = [[], 1, ["Story", "Pier", "P"], 2, ["Story1", "X1", "-1000", "Story2", "X1", "-500"], 0]
    f = filas_tabla(res)
    assert f == [{"Story": "Story1", "Pier": "X1", "P": "-1000"}, {"Story": "Story2", "Pier": "X1", "P": "-500"}]
    assert filas_tabla([[], 1, ["A"], 0, [], 0]) == []
    assert filas_tabla([[], 1, ["A"], 1, ["x"], 1]) == []     # ret ≠ 0


def test_parse_tablas():
    props = parse_pier_props([{"Story": "Story1", "Pier": "X1", "AxisAngle": "90", "WidthBot": "3000",
                               "ThickBot": "130", "MatProp": "ALB", "CGBotX": "100", "CGBotY": "200"}])
    assert props[0].direccion == "Y" and props[0].longitud == 3000 and props[0].espesor == 130
    fz = parse_pier_forces([{"Story": "Story1", "Pier": "X1", "OutputCase": "PG", "Location": "Bottom",
                             "P": "-85000", "V2": "0", "V3": "0", "M2": "0", "M3": "0"}])
    assert fz[0].P == -85000


def test_clasificacion():
    assert es_albanileria("Masonry", "X", "")[0]
    assert es_albanileria("", "MAT1", "MURO_KK13")[0]
    assert not es_albanileria("Concrete", "C210", "PLACA")[0]
    assert es_albanileria("Concrete", "C210", "PLACA", "P1", ["P1"])[0]


def test_snapshot_guardar_abrir(tmp_path):
    s = SnapshotEtabs.abrir(DATA)
    p = s.guardar(tmp_path / "s.json")
    s2 = SnapshotEtabs.abrir(p)
    assert s2.a_dict() == s.a_dict()


def test_snapshot_reproduce_resultado_manual():
    """Pg = -P, Ve = 0.5·V2, Me = 0.5·M3 del sismo severo: deben salir los mismos Vm y DCR que en modo manual."""
    snap = SnapshotEtabs.abrir(DATA)
    cfg = ConfigEtabs(casos_sismo={"X": "SDX", "Y": "SDY"}, combo_gravedad="PG", combo_servicio="SERV",
                      Ap=96e6)
    ed, res = verificar_snapshot(snap, cfg)
    man = verificar_edificio(edificio_ejemplo())
    assert ed.origen == "etabs" and ed.opciones.modo_etabs
    assert len(res.muros) == 36                                  # la placa PL1 no se verifica
    assert any("PL1" in a for a in res.advertencias)
    for rm in man.muros:
        re_ = res.muro(rm.nombre, rm.piso)
        assert re_.valores["Pg"] == pytest.approx(rm.valores["Pg"], rel=1e-9)
        assert re_.valores["Ve"] == pytest.approx(rm.valores["Ve"], rel=1e-9)
        assert re_.valores["Me"] == pytest.approx(rm.valores["Me"], rel=1e-9)
        assert re_.valores["Vm"] == pytest.approx(rm.valores["Vm"], rel=1e-9)
    for pd in man.pisos:
        pe = res.piso_dir(pd.piso, pd.direccion)
        assert pe.VE == pytest.approx(pd.VE, rel=1e-9)
        assert pe.deriva == pytest.approx(pd.deriva, rel=1e-6)


def test_moderado_directo():
    snap = SnapshotEtabs.abrir(DATA)
    cfg = ConfigEtabs(casos_sismo={"X": "SDX", "Y": "SDY"}, combo_gravedad="PG", factor_moderado=1.0, Ap=96e6)
    ed, _ = edificio_desde_snapshot(snap, cfg)
    m = next(m for m in ed.muros if m.nombre == "X1" and m.piso == 1)
    f = snap.fuerza("X1", "Story1", "SDX")
    assert m.Ve == pytest.approx(f.V2)
    assert m.Pm is None                  # sin combo de servicio


def test_copiloto_con_snapshot():
    snap = SnapshotEtabs.abrir(DATA)
    ed, res = verificar_snapshot(snap, ConfigEtabs(casos_sismo={"X": "SDX", "Y": "SDY"}, combo_gravedad="PG", Ap=96e6))
    c = Copiloto(lambda: (ed, res))
    assert "Edificio" in c.ejecutar("revisar edificio")
    assert "X1 piso 1" in c.ejecutar("muro X1 1")
    assert "Piso 2" in c.ejecutar("piso 2")
    assert "requiere" in c.ejecutar("deshacer")
