"""Lista las tablas disponibles de un .EDB (sobre una COPIA en tests/salida/real) y muestra los campos
de las tablas que usa el programa.

Uso (Windows con ETABS 22):  python tools/explorar_tablas.py "C:\\ruta\\modelo.EDB"
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from albanileria.etabs.client import ClienteEtabs  # noqa: E402

TABLAS = ["Pier Section Properties", "Pier Forces", "Story Forces", "Story Drifts", "Mass Summary by Story",
          "Modal Participating Mass Ratios", "Material Properties - Basic Mechanical Properties"]


def main(ruta: str) -> None:
    with ClienteEtabs() as cli:
        cli.iniciar()
        copia = cli.abrir(ruta)
        print(f"Copia abierta: {copia}")
        disp = cli.tablas_disponibles()
        print(f"\n{len(disp)} tablas disponibles:")
        for t in disp:
            print("  -", t)
        casos = cli.casos()
        combos = cli.combos()
        print("\nCasos:", casos)
        print("Combinaciones:", combos)
        for t in TABLAS:
            try:
                filas = cli.tabla(t, casos, combos)
            except Exception as e:      # noqa: BLE001
                print(f"\n[{t}] error: {e}")
                continue
            print(f"\n[{t}] {len(filas)} filas; campos: {list(filas[0].keys()) if filas else '—'}")
            for f in filas[:3]:
                print("   ", f)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1])
