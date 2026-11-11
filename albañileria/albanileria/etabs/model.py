"""Dataclasses del snapshot de un modelo ETABS (todo lo leído, para trabajar sin ETABS abierto).

Unidades: N, mm (las del modelo con SetPresentUnits(9) = N, mm, C).
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class PisoSnap:
    nombre: str
    elevacion: float          # mm (cota superior del piso)
    altura: float             # mm (altura de entrepiso)


@dataclass
class MuroSnap:
    pier: str
    piso: str                 # nombre del piso en ETABS
    material: str = ""
    tipo_material: str = ""   # "Masonry", "Concrete", ...
    seccion: str = ""
    espesor: float = 0.0      # mm
    longitud: float = 0.0     # mm
    altura: float = 0.0       # mm
    angulo: float = 0.0       # grados (AxisAngle del pier)
    x: float = 0.0            # centro de gravedad en planta (mm)
    y: float = 0.0
    areas: list[str] = field(default_factory=list)
    es_albanileria: Optional[bool] = None     # resultado de la clasificación
    motivo: str = ""

    @property
    def direccion(self) -> str:
        a = abs(self.angulo) % 180
        return "X" if a < 45 or a > 135 else "Y"


@dataclass
class FuerzaPier:
    pier: str
    piso: str
    caso: str
    ubicacion: str            # "Top" | "Bottom"
    P: float = 0.0            # N (positivo = tracción en ETABS)
    V2: float = 0.0
    V3: float = 0.0
    M2: float = 0.0           # N·mm
    M3: float = 0.0
    paso: str = ""            # StepType ("Max", "Min", ...)


@dataclass
class FuerzaPiso:
    piso: str
    caso: str
    ubicacion: str
    VX: float = 0.0
    VY: float = 0.0
    paso: str = ""


@dataclass
class DerivaPiso:
    piso: str
    caso: str
    direccion: str
    deriva: float
    paso: str = ""


@dataclass
class SnapshotEtabs:
    archivo: str = ""
    version: str = "1"
    unidades: str = "N, mm, C"
    pisos: list[PisoSnap] = field(default_factory=list)            # de abajo hacia arriba
    muros: list[MuroSnap] = field(default_factory=list)
    fuerzas: list[FuerzaPier] = field(default_factory=list)
    fuerzas_piso: list[FuerzaPiso] = field(default_factory=list)
    derivas: list[DerivaPiso] = field(default_factory=list)
    masas: dict = field(default_factory=dict)                        # piso -> masa (N·s²/mm)
    materiales: dict = field(default_factory=dict)                   # nombre -> {"tipo":..., "E":..., ...}
    casos: list[str] = field(default_factory=list)
    combos: list[str] = field(default_factory=list)
    casos_sismo: dict = field(default_factory=dict)                  # {"X": caso, "Y": caso}
    combo_gravedad: str = ""                                         # D + 0.25 L
    combo_servicio: str = ""                                         # D + L
    modal: list[dict] = field(default_factory=list)
    notas: list[str] = field(default_factory=list)

    # ------------------------------------------------------------ utilidades
    def indice_piso(self, nombre: str) -> int:
        for i, p in enumerate(self.pisos, start=1):
            if p.nombre == nombre:
                return i
        return 0

    def fuerza(self, pier: str, piso: str, caso: str, ubicacion: str = "Bottom") -> Optional[FuerzaPier]:
        """Fuerza de un pier; con varios pasos (Max/Min) devuelve la de mayor |V2|."""
        cands = [f for f in self.fuerzas if f.pier == pier and f.piso == piso and f.caso == caso
                 and f.ubicacion.lower().startswith(ubicacion.lower()[:3])]
        if not cands:
            return None
        return max(cands, key=lambda f: abs(f.V2))

    def a_dict(self) -> dict:
        return asdict(self)

    def guardar(self, ruta: str | Path) -> Path:
        ruta = Path(ruta)
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text(json.dumps(self.a_dict(), indent=1, ensure_ascii=False), encoding="utf-8")
        return ruta

    @staticmethod
    def desde_dict(d: dict) -> "SnapshotEtabs":
        d = dict(d)
        s = SnapshotEtabs(**{k: v for k, v in d.items()
                             if k not in ("pisos", "muros", "fuerzas", "fuerzas_piso", "derivas")})
        s.pisos = [PisoSnap(**p) for p in d.get("pisos", [])]
        s.muros = [MuroSnap(**m) for m in d.get("muros", [])]
        s.fuerzas = [FuerzaPier(**f) for f in d.get("fuerzas", [])]
        s.fuerzas_piso = [FuerzaPiso(**f) for f in d.get("fuerzas_piso", [])]
        s.derivas = [DerivaPiso(**f) for f in d.get("derivas", [])]
        return s

    @staticmethod
    def abrir(ruta: str | Path) -> "SnapshotEtabs":
        return SnapshotEtabs.desde_dict(json.loads(Path(ruta).read_text(encoding="utf-8")))
