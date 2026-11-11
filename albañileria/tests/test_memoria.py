"""Memoria Word (OMML) y CSV."""
import zipfile
from pathlib import Path

from albanileria.e070 import verificar_edificio
from albanileria.ejemplos import edificio_ejemplo
from albanileria.io_csv import exportar_muros, importar_muros
from albanileria.reports.omml import omml
from albanileria.reports.word import generar_memoria

SALIDA = Path(__file__).parent / "salida"


def test_omml_basico():
    for f in ["Vm = 0.5·v'm·α·t·L + 0.23·Pg", "Fa = 0.2·f'm·(1 - (h/(35·t))^2)", "x = √(a² + b)",
              "Vc = 1.5·Vm·Lm/(L·(Nc + 1))", "ΣL·t/Ap ≥ Z·U·S·N/56", "φ·γ·Δ"]:
        x = omml(f)
        assert x and x.startswith("<m:oMath")
    assert "<m:f>" in omml("a/b") and "<m:rad>" in omml("√(x)") and "<m:sSup>" in omml("h³")
    assert "<m:sSub>" in omml("f'm")
    assert omml("") is None and omml("a/(b") is None


def test_memoria_word():
    ed = edificio_ejemplo()
    res = verificar_edificio(ed)
    p = generar_memoria(ed, res, SALIDA / "memoria_prueba.docx")
    assert p.exists() and p.stat().st_size > 50_000
    st = generar_memoria.ultimas_estadisticas
    assert st["omml"] > 500 and st["texto"] == 0
    xml = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8")
    assert xml.count("<m:oMath") > 500
    for t in ("Densidad mínima", "Control de fisuración", "Tabla 11", "FUERA DE ALCANCE", "Resumen final", "DCR ="):
        assert t in xml


def test_csv_ida_y_vuelta(tmp_path):
    ed = edificio_ejemplo()
    ruta = tmp_path / "m.csv"
    exportar_muros(ed.muros, ruta)
    muros = importar_muros(ruta)
    assert len(muros) == len(ed.muros)
    for a, b in zip(ed.muros, muros):
        assert (a.nombre, a.piso, a.direccion, a.Nc) == (b.nombre, b.piso, b.direccion, b.Nc)
        assert abs(a.L - b.L) < 1e-6 and abs(a.t - b.t) < 1e-6 and abs(a.area_trib - b.area_trib) < 1
