"""Alternativas para muros que no cumplen y "copiloto" de comandos de texto.

Las alternativas se evalúan con el motor E.070 sobre una copia del Edificio. En modo ETABS las fuerzas
(Ve, Me) se mantienen hasta reanalizar: el resultado es una ESTIMACIÓN hasta ejecutar "reanalizar".
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Callable, Optional

from ..e070 import verificar_edificio
from ..e070.materiales import unidades_tabla9
from ..models import Edificio, Material, Muro, ResultadoEdificio, mpa


@dataclass
class Alternativa:
    descripcion: str
    tipo: str                      # "espesor" | "unidad" | "agregar_muro" | "longitud"
    muro: str = ""
    piso: int = 0
    parametros: dict = field(default_factory=dict)
    dcr_antes: float = 0.0
    dcr_despues: float = 0.0
    estado_edificio: str = ""
    estimacion: bool = False

    def texto(self) -> str:
        est = " (estimación: fuerzas sin reanalizar)" if self.estimacion else ""
        return (f"{self.descripcion}: DCR {self.dcr_antes:.2f} → {self.dcr_despues:.2f}; "
                f"edificio {self.estado_edificio}{est}")


def _aplicar_en_edificio(ed: Edificio, alt: Alternativa) -> Edificio:
    e = copy.deepcopy(ed)
    if alt.tipo == "espesor":
        for m in e.muros:
            if m.id_grupo == alt.parametros.get("grupo", alt.muro) and m.piso in alt.parametros.get("pisos", [alt.piso]):
                m.t = alt.parametros["t"]
    elif alt.tipo == "unidad":
        u = alt.parametros["unidad"]
        e.material = Material(tipo_unidad=u["nombre"], familia=u["familia"], fb=mpa(u["fb"]), fm=mpa(u["fm"]),
                              vm=mpa(u["vm"]), mortero=ed.material.mortero, fc=ed.material.fc, fy=ed.material.fy,
                              ft=ed.material.ft, gamma=ed.material.gamma)
    elif alt.tipo == "longitud":
        for m in e.muros:
            if m.id_grupo == alt.parametros.get("grupo", alt.muro):
                m.L = alt.parametros["L"]
    elif alt.tipo == "agregar_muro":
        base: Muro = alt.parametros["muro"]
        refs = {m.piso: m for m in ed.muros if m.id_grupo == alt.parametros.get("ref_grupo") and m.direccion == base.direccion}
        # se redistribuye el cortante de piso por rigidez (estimación) en todos los muros
        e.opciones.modo_etabs = False
        for m in e.muros:
            m.Ve = m.Me = None
        for p in range(1, e.N + 1):
            n = copy.deepcopy(base)
            n.piso = p
            ref = refs.get(p)
            n.Pg, n.Pm = (ref.Pg, ref.Pm) if ref else (None, None)
            n.Ve = n.Me = None
            e.muros.append(n)
    return e


def proponer_alternativas(ed: Edificio, res: ResultadoEdificio, nombre: Optional[str] = None,
                          piso: Optional[int] = None, maximo: int = 6) -> list[Alternativa]:
    """Propone alternativas para los muros que no cumplen (o para el muro indicado)."""
    fallan = [r for r in res.muros if r.estado == "NO CUMPLE" and (nombre is None or r.nombre == nombre)
              and (piso is None or r.piso == piso)]
    alts: list[Alternativa] = []
    est = ed.opciones.modo_etabs
    for r in fallan[:3]:
        m = next(x for x in ed.muros if x.nombre == r.nombre and x.piso == r.piso)
        # 1) aumentar espesor (muro en cabeza ≈ 23 cm efectivos, o 15 cm con unidades mayores)
        for t in (150.0, 230.0):
            if t > m.t:
                alts.append(Alternativa(f"Muro {m.nombre}: t = {m.t / 10:.0f} → {t / 10:.0f} cm (todos los pisos)",
                                        "espesor", m.nombre, m.piso, {"t": t, "grupo": m.id_grupo,
                                                                      "pisos": list(range(1, ed.N + 1))}, estimacion=est))
        # 2) alargar el muro 20 %
        alts.append(Alternativa(f"Muro {m.nombre}: L = {m.L / 1000:.2f} → {m.L * 1.2 / 1000:.2f} m", "longitud",
                                m.nombre, m.piso, {"L": round(m.L * 1.2, -1), "grupo": m.id_grupo}, estimacion=est))
    # 3) cambiar el tipo de unidad por una de mayor v'm
    vm_act = ed.material.vm
    mejores = sorted([u for u in unidades_tabla9() if mpa(u["vm"]) > vm_act * 1.001 and not u["nombre"].startswith("Personal")],
                     key=lambda u: u["vm"])
    if mejores and fallan:
        u = mejores[0]
        alts.append(Alternativa(f"Cambiar unidad a {u['nombre']} (v'm = {u['vm']} kgf/cm²)", "unidad",
                                parametros={"unidad": u}, estimacion=est))
    # 4) agregar un muro paralelo en la dirección deficiente
    dirs = {r.direccion for r in fallan} | {p.direccion for p in res.pisos if p.relacion < 1 or p.dcr_densidad > 1}
    for d in sorted(dirs):
        ref = max((x for x in ed.muros if x.direccion == d and x.piso == 1), key=lambda x: x.L, default=None)
        if ref is not None:
            nuevo = copy.deepcopy(ref)
            nuevo.nombre = f"{d}N{len(ed.muros) + 1}"
            nuevo.grupo = nuevo.nombre
            alts.append(Alternativa(f"Agregar muro {nuevo.nombre} (L = {ref.L / 1000:.2f} m) en dirección {d}",
                                    "agregar_muro", nuevo.nombre, 1, {"muro": nuevo, "ref_grupo": ref.id_grupo},
                                    estimacion=True))
    # evaluar
    for a in alts:
        try:
            e2 = _aplicar_en_edificio(ed, a)
            r2 = verificar_edificio(e2)
        except Exception as ex:     # noqa: BLE001
            a.estado_edificio = f"error: {ex}"
            continue
        if a.muro and a.tipo != "agregar_muro":
            r1 = res.muro(a.muro, a.piso)
            rr = r2.muro(a.muro, a.piso)
            a.dcr_antes = r1.dcr_max if r1 else 0.0
            a.dcr_despues = rr.dcr_max if rr else 0.0
        else:
            a.dcr_antes = res.dcr_max
            a.dcr_despues = r2.dcr_max
        a.estado_edificio = r2.estado
    alts.sort(key=lambda a: a.dcr_despues)
    return alts[:maximo]


def aplicar_en_etabs(cli, ed: Edificio, alt: Alternativa, snap) -> list[tuple[str, str]]:
    """Aplica una alternativa de espesor en el modelo COPIA de ETABS. Devuelve [(área, propiedad anterior)]
    para poder deshacer. Sólo se implementa el cambio de espesor (lo demás: FUERA DE ALCANCE en ETABS)."""
    if alt.tipo != "espesor":
        raise NotImplementedError("FUERA DE ALCANCE: sólo el cambio de espesor se aplica automáticamente en ETABS; "
                                  "las demás alternativas se modelan a mano.")
    m = cli.m
    t = alt.parametros["t"]
    historial = []
    pier = alt.parametros.get("grupo", alt.muro)
    areas = list(m.AreaObj.GetNameList())[1]
    for a in areas:
        try:
            pr = list(m.AreaObj.GetPier(a))[0]
        except Exception:           # noqa: BLE001
            continue
        if pr != pier:
            continue
        prop = list(m.AreaObj.GetProperty(a))[0]
        r = list(m.PropArea.GetWall(prop))
        # [WallPropType, ShellType, MatProp, Thickness, color, notes, GUID, ret]
        mat, tipo_shell = r[2], r[1]
        nueva = f"{prop}_t{t / 10:.0f}"
        m.PropArea.SetWall(nueva, 1, tipo_shell, mat, t)
        m.AreaObj.SetProperty(a, nueva)
        historial.append((a, prop))
    return historial


def deshacer_en_etabs(cli, historial: list[tuple[str, str]]) -> None:
    for a, prop in historial:
        cli.m.AreaObj.SetProperty(a, prop)


class Copiloto:
    """Intérprete de comandos de texto. Las acciones con ETABS se delegan en callbacks."""

    AYUDA = ("Comandos: revisar edificio | muro <nombre> [piso] | piso <n> | alternativas [muro] | "
             "aplicar <n> | deshacer | reanalizar | ayuda")

    def __init__(self, obtener: Callable[[], tuple[Edificio, ResultadoEdificio]],
                 aplicar: Optional[Callable[[Alternativa], str]] = None,
                 deshacer: Optional[Callable[[], str]] = None,
                 reanalizar: Optional[Callable[[], str]] = None):
        self.obtener = obtener
        self.cb_aplicar, self.cb_deshacer, self.cb_reanalizar = aplicar, deshacer, reanalizar
        self.ultimas: list[Alternativa] = []

    def ejecutar(self, texto: str) -> str:
        p = texto.strip().split()
        if not p:
            return self.AYUDA
        cmd = p[0].lower()
        if cmd in ("ayuda", "help", "?"):
            return self.AYUDA
        ed, res = self.obtener()
        if res is None:
            return "Primero lea el modelo y verifique el edificio."
        if cmd == "revisar":
            return self._revisar(res)
        if cmd == "muro" and len(p) > 1:
            return self._muro(res, p[1], int(p[2]) if len(p) > 2 and p[2].isdigit() else None)
        if cmd == "piso" and len(p) > 1 and p[1].isdigit():
            return self._piso(res, int(p[1]))
        if cmd == "alternativas":
            self.ultimas = proponer_alternativas(ed, res, p[1] if len(p) > 1 else None)
            if not self.ultimas:
                return "No hay muros que no cumplan: no se proponen alternativas."
            return "\n".join(f"[{i}] {a.texto()}" for i, a in enumerate(self.ultimas, start=1)) + \
                "\nUse 'aplicar <n>' para crear la alternativa en el modelo copia."
        if cmd == "aplicar":
            if not self.ultimas:
                return "Primero ejecute 'alternativas'."
            k = int(p[1]) if len(p) > 1 and p[1].isdigit() else 1
            if not 1 <= k <= len(self.ultimas):
                return f"Elija entre 1 y {len(self.ultimas)}."
            if self.cb_aplicar is None:
                return "Aplicar requiere conexión con ETABS."
            return self.cb_aplicar(self.ultimas[k - 1])
        if cmd == "deshacer":
            return self.cb_deshacer() if self.cb_deshacer else "Deshacer requiere conexión con ETABS."
        if cmd == "reanalizar":
            return self.cb_reanalizar() if self.cb_reanalizar else "Reanalizar requiere conexión con ETABS."
        return "Comando no reconocido. " + self.AYUDA

    @staticmethod
    def _revisar(res: ResultadoEdificio) -> str:
        lin = [f"Edificio: {res.estado} (DCR máx {res.dcr_max:.3f})"]
        for pd in res.pisos:
            lin.append(f"  Piso {pd.piso} {pd.direccion}: densidad {pd.densidad_existente:.4f}/{pd.densidad_requerida:.4f}, "
                       f"ΣVm/VE = {pd.relacion:.2f}" + (" (elástico)" if pd.elastico else "")
                       + (f", no cumplen: {', '.join(pd.muros_no_cumplen)}" if pd.muros_no_cumplen else ""))
        nc = [n.nombre for n in res.no_portantes if n.estado == "NO CUMPLE"]
        if nc:
            lin.append("  No portantes que requieren arriostre: " + ", ".join(nc))
        return "\n".join(lin)

    @staticmethod
    def _muro(res: ResultadoEdificio, nombre: str, piso: Optional[int]) -> str:
        ms = [m for m in res.muros if m.nombre.lower() == nombre.lower() and (piso is None or m.piso == piso)]
        if not ms:
            return f"No existe el muro {nombre}."
        lin = []
        for m in ms:
            v = m.valores
            lin.append(f"{m.nombre} piso {m.piso} ({m.direccion}): {m.estado}, DCR {m.dcr_max:.3f}")
            if v:
                lin.append(f"  Pg {v['Pg'] / 1000:.1f} kN, Ve {v['Ve'] / 1000:.1f} kN, Me {v['Me'] / 1e6:.1f} kN·m, "
                           f"α {v['alfa']:.2f}, Vm {v['Vm'] / 1000:.1f} kN, Ve/0.55Vm = {v['dcr_fis']:.2f}, "
                           f"Vu {v.get('Vu', 0) / 1000:.1f} kN ({v.get('tipo_diseno', '')})")
            g = m.check_gobernante
            if g:
                lin.append(f"  Gobierna: {g.nombre} ({g.referencia})")
            if m.confinamiento:
                rc = m.confinamiento
                lin.append(f"  Columnas {rc.seccion_ext}: {rc.refuerzo_ext}; solera {rc.refuerzo_solera}")
        return "\n".join(lin)

    @staticmethod
    def _piso(res: ResultadoEdificio, n: int) -> str:
        lin = []
        for pd in res.pisos:
            if pd.piso == n:
                lin.append(f"Piso {n} {pd.direccion}: VE = {pd.VE / 1000:.1f} kN, ΣVm = {pd.suma_Vm / 1000:.1f} kN "
                           f"(ΣVm/VE = {pd.relacion:.2f}); agrietados: {', '.join(pd.muros_agrietados) or '—'}; "
                           f"no cumplen: {', '.join(pd.muros_no_cumplen) or '—'}")
        return "\n".join(lin) or f"No existe el piso {n}."
