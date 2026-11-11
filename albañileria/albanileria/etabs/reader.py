"""Lectura del modelo ETABS a un SnapshotEtabs (funciones puras + lectura COM)."""
from __future__ import annotations

from typing import Iterable

from .clasificar import clasificar_muros
from .model import DerivaPiso, FuerzaPier, FuerzaPiso, MuroSnap, PisoSnap, SnapshotEtabs


def campo(fila: dict, *candidatos, defecto=None):
    """Busca un campo por nombre exacto o, si no, sin distinguir mayúsculas / espacios."""
    for c in candidatos:
        if c in fila:
            return fila[c]
    norm = {k.lower().replace(" ", "").replace("_", ""): v for k, v in fila.items()}
    for c in candidatos:
        k = c.lower().replace(" ", "").replace("_", "")
        if k in norm:
            return norm[k]
    return defecto


def num(v, defecto: float = 0.0) -> float:
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return defecto


# ------------------------------------------------------------------ parsers de tablas
def parse_pier_props(filas: Iterable[dict]) -> list[MuroSnap]:
    """Tabla 'Pier Section Properties'."""
    out = []
    for f in filas:
        out.append(MuroSnap(
            pier=str(campo(f, "Pier", "Label")), piso=str(campo(f, "Story")),
            material=str(campo(f, "MatProp", "Material", defecto="")),
            espesor=num(campo(f, "ThickBot", "Thickness Bottom", "ThickTop")),
            longitud=num(campo(f, "WidthBot", "Width Bottom", "WidthTop")),
            angulo=num(campo(f, "AxisAngle", "Axis Angle")),
            x=num(campo(f, "CGBotX", "CG Bottom X", "CGTopX")), y=num(campo(f, "CGBotY", "CG Bottom Y", "CGTopY"))))
    return out


def parse_pier_forces(filas: Iterable[dict]) -> list[FuerzaPier]:
    """Tabla 'Pier Forces'."""
    return [FuerzaPier(pier=str(campo(f, "Pier")), piso=str(campo(f, "Story")), caso=str(campo(f, "OutputCase")),
                       ubicacion=str(campo(f, "Location", defecto="Bottom")), P=num(campo(f, "P")),
                       V2=num(campo(f, "V2")), V3=num(campo(f, "V3")), M2=num(campo(f, "M2")), M3=num(campo(f, "M3")),
                       paso=str(campo(f, "StepType", defecto="") or "")) for f in filas]


def parse_story_forces(filas: Iterable[dict]) -> list[FuerzaPiso]:
    return [FuerzaPiso(piso=str(campo(f, "Story")), caso=str(campo(f, "OutputCase")),
                       ubicacion=str(campo(f, "Location", defecto="Bottom")), VX=num(campo(f, "VX")),
                       VY=num(campo(f, "VY")), paso=str(campo(f, "StepType", defecto="") or "")) for f in filas]


def parse_drifts(filas: Iterable[dict]) -> list[DerivaPiso]:
    return [DerivaPiso(piso=str(campo(f, "Story")), caso=str(campo(f, "OutputCase")),
                       direccion=str(campo(f, "Direction", defecto="X")).upper()[:1], deriva=num(campo(f, "Drift")),
                       paso=str(campo(f, "StepType", defecto="") or "")) for f in filas]


def parse_masas(filas: Iterable[dict]) -> dict:
    """'Mass Summary by Story': masa en N·s²/mm (1 N·s²/mm = 1000 kg)."""
    return {str(campo(f, "Story")): num(campo(f, "UX", "MassX")) for f in filas if campo(f, "Story")}


# ------------------------------------------------------------------ lectura COM
def leer_modelo(cli, casos_sismo: dict | None = None, combo_gravedad: str = "", combo_servicio: str = "",
                lista_usuario: list[str] | None = None, avisar=print) -> SnapshotEtabs:
    """Lee el modelo abierto en `cli` (ClienteEtabs) y devuelve un snapshot."""
    s = SnapshotEtabs(archivo=str(cli.ruta_modelo or ""))
    avisar("Leyendo pisos…")
    s.pisos = [PisoSnap(n, e, h) for n, e, h in cli.pisos()]
    s.casos = cli.casos()
    s.combos = cli.combos()
    avisar("Leyendo materiales…")
    s.materiales = cli.materiales()
    avisar("Leyendo propiedades de piers…")
    s.muros = parse_pier_props(cli.tabla("Pier Section Properties"))
    alt = {p.nombre: p.altura for p in s.pisos}
    for m in s.muros:
        m.altura = alt.get(m.piso, 0.0)
        m.tipo_material = s.materiales.get(m.material, {}).get("tipo", "")
    clasificar_muros(s.muros, lista_usuario)
    s.casos_sismo = dict(casos_sismo or {})
    s.combo_gravedad, s.combo_servicio = combo_gravedad, combo_servicio
    pedidos = [c for c in list(s.casos_sismo.values()) + [combo_gravedad, combo_servicio] if c]
    casos_sel = [c for c in pedidos if c in s.casos]
    combos_sel = [c for c in pedidos if c in s.combos]
    faltan = [c for c in pedidos if c not in s.casos and c not in s.combos]
    if faltan:
        s.notas.append(f"No existen en el modelo: {', '.join(faltan)}")
    if not cli.bloqueado():
        avisar("El modelo no tiene resultados: ejecutando el análisis…")
        cli.analizar()
    avisar("Leyendo fuerzas en piers…")
    s.fuerzas = parse_pier_forces(cli.tabla("Pier Forces", casos_sel, combos_sel))
    avisar("Leyendo cortantes y derivas de piso…")
    s.fuerzas_piso = parse_story_forces(cli.tabla("Story Forces", casos_sel, combos_sel))
    s.derivas = parse_drifts(cli.tabla("Story Drifts", casos_sel, combos_sel))
    try:
        s.masas = parse_masas(cli.tabla("Mass Summary by Story"))
    except Exception as e:      # noqa: BLE001
        s.notas.append(f"No se leyó 'Mass Summary by Story': {e}")
    try:
        s.modal = cli.tabla("Modal Participating Mass Ratios")
    except Exception:           # noqa: BLE001
        pass
    return s
