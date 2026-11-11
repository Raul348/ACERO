"""Edificios de ejemplo."""
from __future__ import annotations

from ..models import Edificio, ElementoNoPortante, Muro, DatosMemoria, KPA


def edificio_ejemplo() -> Edificio:
    """Vivienda multifamiliar de 3 pisos en Lima (zona 4, S1), planta 8 x 12 m, muros de 13 cm."""
    ed = Edificio(nombre="Ejemplo 3 pisos", zona=4, categoria="C", suelo="S1", N=3, h_entrepiso=2600.0,
                  Ap=96.0e6, Lx=8000.0, Ly=12000.0)
    ed.memoria = DatosMemoria(proyecto="Vivienda multifamiliar de 3 pisos (ejemplo)", ubicacion="Lima",
                              ingeniero="", fecha="")
    muros_planta = [
        # nombre, dir, L, x, y, area trib (m²), Nc
        ("X1", "X", 3500, 1750, 0, 6.0, 2),
        ("X2", "X", 3500, 6250, 0, 6.0, 2),
        ("X3", "X", 4000, 2000, 6000, 10.0, 2),
        ("X4", "X", 3000, 6500, 6000, 8.0, 2),
        ("X5", "X", 3500, 1750, 12000, 6.0, 2),
        ("X6", "X", 3500, 6250, 12000, 6.0, 2),
        ("Y1", "Y", 6000, 0, 3000, 9.0, 3),
        ("Y2", "Y", 6000, 0, 9000, 9.0, 3),
        ("Y3", "Y", 4500, 4000, 3000, 12.0, 2),
        ("Y4", "Y", 4500, 4000, 9000, 12.0, 2),
        ("Y5", "Y", 6000, 8000, 3000, 9.0, 3),
        ("Y6", "Y", 6000, 8000, 9000, 9.0, 3),
    ]
    for piso in range(1, ed.N + 1):
        for nom, d, L, x, y, at, nc in muros_planta:
            ed.muros.append(Muro(nombre=nom, piso=piso, direccion=d, L=float(L), t=130.0, h=2400.0,
                                 x=float(x), y=float(y), area_trib=at * 1e6, Nc=nc))
    ed.no_portantes = [
        ElementoNoPortante("Parapeto azotea", tipo="parapeto", caso=4, a=1000.0, b=3000.0, t=130.0, piso=3),
        ElementoNoPortante("Tabique cocina", tipo="tabique", caso=1, a=2400.0, b=3000.0, t=130.0, piso=1),
        ElementoNoPortante("Cerco perimétrico", tipo="cerco", caso=2, a=3000.0, b=2400.0, t=130.0, piso=1),
    ]
    return ed
