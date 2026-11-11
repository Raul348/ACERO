"""Genera tests/data/snapshot_ejemplo.json: un snapshot con el formato de ETABS construido a partir del
edificio de ejemplo (fuerzas del reparto por rigidez), más una placa de concreto para probar la clasificación.

Uso: python tools/generar_snapshot_ejemplo.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from albanileria.e070 import verificar_edificio  # noqa: E402
from albanileria.e070.verificar import cargas_gravedad  # noqa: E402
from albanileria.ejemplos import edificio_ejemplo  # noqa: E402
from albanileria.etabs.model import (DerivaPiso, FuerzaPier, FuerzaPiso, MuroSnap, PisoSnap,  # noqa: E402
                                     SnapshotEtabs)


def main() -> Path:
    ed = edificio_ejemplo()
    res = verificar_edificio(ed)
    s = SnapshotEtabs(archivo="EJEMPLO_ALBANILERIA.EDB")
    s.pisos = [PisoSnap(f"Story{i}", 2600.0 * i, 2600.0) for i in range(1, ed.N + 1)]
    s.casos = ["Dead", "Live", "SDX", "SDY"]
    s.combos = ["PG", "SERV"]
    s.casos_sismo = {"X": "SDX", "Y": "SDY"}
    s.combo_gravedad, s.combo_servicio = "PG", "SERV"
    s.materiales = {"ALB_KK": {"tipo": "Masonry", "E": 3187.0}, "C210": {"tipo": "Concrete", "E": 21737.0}}
    for m in ed.muros:
        piso = f"Story{m.piso}"
        s.muros.append(MuroSnap(pier=m.nombre, piso=piso, material="ALB_KK", tipo_material="Masonry",
                                seccion="MURO13", espesor=m.t, longitud=m.L, altura=2600.0,
                                angulo=0.0 if m.direccion == "X" else 90.0, x=m.x, y=m.y))
        r = res.muro(m.nombre, m.piso)
        Pg, Pm, _, _ = cargas_gravedad(ed, m)
        caso = "SDX" if m.direccion == "X" else "SDY"
        # en ETABS P > 0 es tracción; el sismo del modelo es el SEVERO (el programa lo multiplica por 0.5)
        s.fuerzas += [
            FuerzaPier(m.nombre, piso, "PG", "Bottom", P=-Pg),
            FuerzaPier(m.nombre, piso, "SERV", "Bottom", P=-Pm),
            FuerzaPier(m.nombre, piso, caso, "Bottom", P=0.0, V2=2 * r.valores["Ve"], M3=2 * r.valores["Me"], paso="Max"),
            FuerzaPier(m.nombre, piso, caso, "Top", P=0.0, V2=2 * r.valores["Ve"], M3=0.0, paso="Max"),
        ]
    # placa de concreto (debe listarse pero no verificarse)
    for i in range(1, ed.N + 1):
        s.muros.append(MuroSnap(pier="PL1", piso=f"Story{i}", material="C210", tipo_material="Concrete",
                                seccion="PLACA20", espesor=200.0, longitud=1500.0, altura=2600.0, angulo=90.0,
                                x=4000.0, y=6000.0))
    for d, caso in (("X", "SDX"), ("Y", "SDY")):
        sis = res.sismo[d]
        for i, p in enumerate(s.pisos):
            fp = FuerzaPiso(p.nombre, caso, "Bottom", paso="Max")
            setattr(fp, "VX" if d == "X" else "VY", sis.VE[i])
            s.fuerzas_piso.append(fp)
            pd = res.piso_dir(i + 1, d)
            s.derivas.append(DerivaPiso(p.nombre, caso, d, pd.deriva / (0.75 * 3.0), paso="Max"))
    s.masas = {p.nombre: w / 9806.65 for p, w in zip(s.pisos, res.sismo["X"].pesos)}
    s.notas.append("Snapshot sintético generado desde el edificio de ejemplo (no proviene de un .EDB real).")
    ruta = Path(__file__).resolve().parents[1] / "tests" / "data" / "snapshot_ejemplo.json"
    s.guardar(ruta)
    print(f"Snapshot guardado en {ruta}")
    return ruta


if __name__ == "__main__":
    main()
