"""Del snapshot de ETABS al Edificio del motor E.070 y verificación de todo el edificio."""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from ..e070 import verificar_edificio
from ..models import DatosMemoria, Edificio, Material, Muro, OpcionesAnalisis, ResultadoEdificio
from .model import SnapshotEtabs

G_MM_S2 = 9806.65


@dataclass
class ConfigEtabs:
    casos_sismo: dict = field(default_factory=lambda: {"X": "", "Y": ""})
    combo_gravedad: str = ""          # D + 0.25 L
    combo_servicio: str = ""          # D + L (opcional)
    factor_moderado: float = 0.5      # 0.5 si los casos son del sismo severo; 1.0 si son combos "moderado"
    zona: int = 4
    categoria: str = "C"
    suelo: str = "S1"
    Ap: Optional[float] = None        # mm²; None = estimada con los muros del piso 1
    material: Material = field(default_factory=Material)
    lista_usuario: Optional[list[str]] = None
    proyecto: str = ""


def edificio_desde_snapshot(snap: SnapshotEtabs, cfg: ConfigEtabs) -> tuple[Edificio, list[str]]:
    """Construye el Edificio (modo ETABS) y devuelve también la lista de avisos."""
    av: list[str] = list(snap.notas)
    op = OpcionesAnalisis(modo_etabs=True)
    N = len(snap.pisos)
    if N == 0:
        raise ValueError("El snapshot no tiene pisos.")
    if N > 5:
        av.append(f"El modelo tiene {N} pisos: la E.070 (albañilería confinada) está pensada para edificios bajos; "
                  "el programa verifica, pero revise el alcance.")
    ed = Edificio(nombre=Path(snap.archivo).stem or "ETABS", zona=cfg.zona, categoria=cfg.categoria, suelo=cfg.suelo,
                  N=N, h_entrepiso=snap.pisos[0].altura, material=cfg.material, opciones=op, origen="etabs")
    ed.alturas = [p.altura for p in snap.pisos]
    ed.memoria = DatosMemoria(proyecto=cfg.proyecto or ed.nombre)
    fm = cfg.factor_moderado
    if cfg.lista_usuario is not None or any(m.es_albanileria is None for m in snap.muros):
        from .clasificar import clasificar_muros
        clasificar_muros(snap.muros, cfg.lista_usuario)
    placas = []
    for ms in snap.muros:
        i = snap.indice_piso(ms.piso)
        if i == 0:
            continue
        if not ms.es_albanileria:
            placas.append(f"{ms.pier}@{ms.piso}")
            continue
        d = ms.direccion
        caso = cfg.casos_sismo.get(d, "")
        fs = snap.fuerza(ms.pier, ms.piso, caso) if caso else None
        fg = snap.fuerza(ms.pier, ms.piso, cfg.combo_gravedad) if cfg.combo_gravedad else None
        fsv = snap.fuerza(ms.pier, ms.piso, cfg.combo_servicio) if cfg.combo_servicio else None
        h_libre = max((ms.altura or snap.pisos[i - 1].altura) - op.peralte_losa, 100.0)
        lim = min(2 * h_libre, 5000.0)
        nc = max(2, math.ceil(ms.longitud / lim - 1e-9) + 1)
        m = Muro(nombre=ms.pier, piso=i, direccion=d, L=ms.longitud, t=ms.espesor, h=h_libre, x=ms.x, y=ms.y,
                 Nc=nc, grupo=ms.pier, material=ms.material,
                 Pg=(-fg.P if fg else None), Pm=(-fsv.P if fsv else None),
                 Ve=(abs(fs.V2) * fm if fs else 0.0), Me=(abs(fs.M3) * fm if fs else 0.0))
        if fg is None:
            av.append(f"{ms.pier} piso {i}: sin fuerzas para el combo de gravedad '{cfg.combo_gravedad}'.")
        if fs is None:
            av.append(f"{ms.pier} piso {i}: sin fuerzas para el caso de sismo '{caso}'.")
        ed.muros.append(m)
    if placas:
        av.append("Piers que no son de albañilería (placas de concreto u otros): " + ", ".join(placas)
                  + " — se listan pero no se verifican (FUERA DE ALCANCE; programa de concreto).")
    # cortantes de piso del sismo severo
    VE = {}
    for d, comp in (("X", "VX"), ("Y", "VY")):
        caso = cfg.casos_sismo.get(d, "")
        vals = []
        for p in snap.pisos:
            fps = [f for f in snap.fuerzas_piso if f.piso == p.nombre and f.caso == caso
                   and f.ubicacion.lower().startswith("bot")]
            if fps:
                vals.append(max(abs(getattr(f, comp)) for f in fps) * 2 * fm)
            else:
                vals.append(sum(m.Ve for m in ed.muros if m.direccion == d and m.piso == snap.indice_piso(p.nombre))
                            * 2 if caso else 0.0)
        if caso and any(vals):
            VE[d] = vals
    ed.VE_override = VE or None
    # derivas elásticas
    der = {}
    for d in ("X", "Y"):
        caso = cfg.casos_sismo.get(d, "")
        vals = []
        for p in snap.pisos:
            ds = [x.deriva for x in snap.derivas if x.piso == p.nombre and x.caso == caso and x.direccion == d]
            vals.append(max(ds) * 2 * fm if ds else 0.0)
        if any(vals):
            der[d] = vals
    ed.derivas = der or None
    # pesos
    if snap.masas:
        pes = [snap.masas.get(p.nombre, 0.0) * G_MM_S2 for p in snap.pisos]
        if any(pes):
            ed.pesos_override = pes
    # área en planta
    if cfg.Ap:
        ed.Ap = cfg.Ap
    else:
        m1 = [m for m in ed.muros if m.piso == 1] or ed.muros
        if m1:
            xs = [m.x + s * (m.L / 2 if m.direccion == "X" else 0) for m in m1 for s in (-1, 1)]
            ys = [m.y + s * (m.L / 2 if m.direccion == "Y" else 0) for m in m1 for s in (-1, 1)]
            ed.Lx, ed.Ly = max(xs) - min(xs), max(ys) - min(ys)
            ed.Ap = max(ed.Lx * ed.Ly, 1e6)
            av.append(f"Área en planta estimada con el contorno de los muros: Ap = {ed.Ap / 1e6:.1f} m² "
                      "(ingrésela en la pestaña Edificio si es distinta).")
    return ed, av


def verificar_snapshot(snap: SnapshotEtabs, cfg: ConfigEtabs) -> tuple[Edificio, ResultadoEdificio]:
    ed, av = edificio_desde_snapshot(snap, cfg)
    res = verificar_edificio(ed)
    res.advertencias = av + res.advertencias
    return ed, res


def procesar_edb(ruta_edb: str | Path, cfg: ConfigEtabs, memoria: bool = True,
                 avisar: Callable[[str], None] = print) -> dict:
    """Flujo completo: copia el .EDB, abre ETABS oculto, lee, guarda snapshot, verifica y genera la memoria."""
    from ..io_csv import exportar_tabla
    from ..reports.word import generar_memoria
    from .client import ClienteEtabs
    from .reader import leer_modelo

    with ClienteEtabs() as cli:
        avisar("Iniciando ETABS (instancia oculta)…")
        cli.iniciar()
        copia = cli.abrir(ruta_edb)
        avisar(f"Trabajando sobre la copia: {copia}")
        if not cfg.combo_gravedad:
            cfg.combo_gravedad = cli.crear_combo_gravedad()
            avisar(f"Se creó el combo {cfg.combo_gravedad} en la copia.")
        snap = leer_modelo(cli, cfg.casos_sismo, cfg.combo_gravedad, cfg.combo_servicio, cfg.lista_usuario, avisar)
    base = copia.with_suffix("")
    ruta_snap = snap.guardar(str(base) + "_snapshot.json")
    avisar(f"Snapshot guardado: {ruta_snap}")
    ed, res = verificar_snapshot(snap, cfg)
    salida = {"copia": copia, "snapshot": ruta_snap, "edificio": ed, "resultado": res}
    filas = [[m.nombre, m.piso, m.direccion, f"{m.valores.get('Pg', 0) / 1000:.2f}", f"{m.valores.get('Ve', 0) / 1000:.2f}",
              f"{m.valores.get('Me', 0) / 1e6:.2f}", f"{m.valores.get('Vm', 0) / 1000:.2f}", f"{m.dcr_max:.3f}", m.estado]
             for m in res.muros]
    ruta_csv = Path(str(base) + "_resultados.csv")
    exportar_tabla(["Muro", "Piso", "Dir", "Pg kN", "Ve kN", "Me kN·m", "Vm kN", "DCR", "Estado"], filas, ruta_csv)
    salida["csv"] = ruta_csv
    if memoria:
        salida["memoria"] = generar_memoria(ed, res, str(base) + "_memoria.docx")
        avisar(f"Memoria generada: {salida['memoria']}")
    return salida
