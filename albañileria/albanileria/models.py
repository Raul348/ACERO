"""Modelo de datos del programa ALBAÑILERÍA E.070.

Unidades internas SIEMPRE: fuerza N, longitud mm, esfuerzo MPa (N/mm²),
momento N·mm. Las conversiones a kN, kN·m y kgf/cm² se hacen sólo al mostrar.
"""
from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------- unidades
KGF_CM2_POR_MPA = 10.197          # 1 MPa = 10.197 kgf/cm²
MPA_POR_KGF_CM2 = 1.0 / KGF_CM2_POR_MPA
KPA = 1e-3                        # 1 kN/m² = 1e-3 N/mm² (MPa)
KN_M3 = 1e-6                      # 1 kN/m³ = 1e-6 N/mm³

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

ESTADO_CUMPLE = "CUMPLE"
ESTADO_NO_CUMPLE = "NO CUMPLE"
ESTADO_FUERA = "FUERA DE ALCANCE"


def kgf_cm2(mpa: float) -> float:
    return mpa * KGF_CM2_POR_MPA


def mpa(kgfcm2: float) -> float:
    return kgfcm2 * MPA_POR_KGF_CM2


def cargar_json(nombre: str) -> dict:
    with open(DATA_DIR / nombre, encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------- formato de números
def fmt_kN(n: float, dec: int = 2) -> str:
    return f"{n / 1000:.{dec}f} kN"


def fmt_kNm(nmm: float, dec: int = 2) -> str:
    return f"{nmm / 1e6:.{dec}f} kN·m"


def fmt_MPa(v: float, dec: int = 3) -> str:
    return f"{v:.{dec}f} MPa ({kgf_cm2(v):.2f} kgf/cm²)"


def fmt_mm(v: float, dec: int = 0) -> str:
    return f"{v:.{dec}f} mm"


def fmt_mm2(v: float, dec: int = 0) -> str:
    return f"{v:.{dec}f} mm² ({v / 100:.2f} cm²)"


def fmt(v: float, dec: int = 3) -> str:
    return f"{v:.{dec}f}"


# ------------------------------------------------------------- resultados
@dataclass
class Step:              # un renglón de la memoria de cálculo manual
    simbolo: str         # p.ej. "Vm"
    valor: str           # p.ej. "152.3 kN"
    descripcion: str = ""
    formula: str = ""    # simbólica, p.ej. "Vm = 0.5·v'm·α·t·L + 0.23·Pg"
    sustitucion: str = ""  # con números, p.ej. "0.5·0.81·1.0·130·3200 + 0.23·85000"


@dataclass
class Check:             # un estado límite
    nombre: str          # p.ej. "Control de fisuración ante sismo moderado"
    referencia: str      # p.ej. "E.070 Art. 26.2"
    demanda: float
    capacidad: float
    unidad: str          # "kN", "kN·m", "MPa" o "" (adimensional)
    pasos: list[Step] = field(default_factory=list)
    notas: list[str] = field(default_factory=list)
    aplica: bool = True

    @property
    def dcr(self) -> float:
        if not self.aplica:
            return 0.0
        if self.capacidad == 0:
            return math.inf if self.demanda > 0 else 0.0
        return self.demanda / self.capacidad

    @property
    def cumple(self) -> bool:
        return self.dcr <= 1.0


def check_fuera_de_alcance(nombre: str, referencia: str, motivo: str) -> Check:
    return Check(nombre, referencia, 0.0, 0.0, "", notas=[f"FUERA DE ALCANCE: {motivo}"], aplica=False)


@dataclass
class ResultadoConfinamiento:
    """Diseño de columnas de confinamiento y solera de un muro."""
    tipo: str = ""                  # "agrietado", "no agrietado", "mínimo"
    Pc_ext: float = 0.0
    Pc_int: float = 0.0
    Vc_ext: float = 0.0
    T_ext: float = 0.0
    C_ext: float = 0.0
    Vc_int: float = 0.0
    T_int: float = 0.0
    C_int: float = 0.0
    M: float = 0.0
    F: float = 0.0
    Acf: float = 0.0
    An_req: float = 0.0
    As_req_ext: float = 0.0
    As_req_int: float = 0.0
    seccion_ext: str = ""           # "13 x 25 cm"
    seccion_int: str = ""
    b_col: float = 0.0
    d_col: float = 0.0
    d_col_int: float = 0.0
    refuerzo_ext: str = ""          # "4 Ø 3/8\""
    refuerzo_int: str = ""
    As_prov_ext: float = 0.0
    As_prov_int: float = 0.0
    estribos: str = ""
    Ts: float = 0.0
    As_solera: float = 0.0
    refuerzo_solera: str = ""
    seccion_solera: str = ""
    estribos_solera: str = ""
    checks: list[Check] = field(default_factory=list)
    notas: list[str] = field(default_factory=list)


@dataclass
class ResultadoMuro:
    nombre: str
    piso: int
    direccion: str
    checks: list[Check] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)
    valores: dict = field(default_factory=dict)        # σm, Ve, Me, α, Vm, Vu, Mu, ...
    confinamiento: Optional[ResultadoConfinamiento] = None
    fuera_de_alcance: bool = False

    @property
    def checks_aplicables(self) -> list[Check]:
        return [c for c in self.checks if c.aplica]

    @property
    def dcr_max(self) -> float:
        ap = self.checks_aplicables
        return max((c.dcr for c in ap), default=0.0)

    @property
    def dcr_verificacion(self) -> float:
        """DCR máximo de las verificaciones (sin los chequeos de la sección propuesta de columnas y soleras)."""
        return max((c.dcr for c in self.checks_aplicables if not c.nombre.startswith(("Columna", "Solera"))),
                   default=0.0)

    @property
    def check_gobernante(self) -> Optional[Check]:
        ap = self.checks_aplicables
        return max(ap, key=lambda c: c.dcr) if ap else None

    @property
    def estado(self) -> str:
        if self.fuera_de_alcance:
            return ESTADO_FUERA
        return ESTADO_CUMPLE if self.dcr_max <= 1.0 else ESTADO_NO_CUMPLE


@dataclass
class ResultadoPisoDir:
    piso: int
    direccion: str
    densidad_existente: float = 0.0
    densidad_requerida: float = 0.0
    VE: float = 0.0                 # cortante del sismo severo (N)
    Ve_total: float = 0.0           # cortante del sismo moderado (N)
    suma_Vm: float = 0.0
    relacion: float = 0.0           # ΣVm/VE
    elastico: bool = False          # ΣVm ≥ 3·VE
    muros: list[str] = field(default_factory=list)
    muros_no_cumplen: list[str] = field(default_factory=list)
    muros_agrietados: list[str] = field(default_factory=list)
    deriva: Optional[float] = None
    checks: list[Check] = field(default_factory=list)

    @property
    def dcr_densidad(self) -> float:
        if self.densidad_existente <= 0:
            return math.inf if self.densidad_requerida > 0 else 0.0
        return self.densidad_requerida / self.densidad_existente


@dataclass
class ResultadoNoPortante:
    nombre: str
    tipo: str
    valores: dict = field(default_factory=dict)
    checks: list[Check] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)

    @property
    def dcr_max(self) -> float:
        return max((c.dcr for c in self.checks if c.aplica), default=0.0)

    @property
    def estado(self) -> str:
        return ESTADO_CUMPLE if self.dcr_max <= 1.0 else ESTADO_NO_CUMPLE


@dataclass
class ResultadoSismo:
    direccion: str = "X"
    Z: float = 0.0
    U: float = 0.0
    S: float = 0.0
    Tp: float = 0.0
    Tl: float = 0.0
    T: float = 0.0
    C: float = 0.0
    R: float = 3.0
    k: float = 1.0
    P: float = 0.0                       # peso sísmico total (N)
    V: float = 0.0                       # cortante basal sismo severo (N)
    pesos: list[float] = field(default_factory=list)       # por nivel (N)
    alturas: list[float] = field(default_factory=list)     # acumuladas hi (mm)
    F: list[float] = field(default_factory=list)           # fuerzas por nivel (N)
    VE: list[float] = field(default_factory=list)          # cortante por piso, sismo severo
    Ve: list[float] = field(default_factory=list)          # cortante por piso, sismo moderado
    origen: str = "E.030 estático"
    pasos: list[Step] = field(default_factory=list)


@dataclass
class ResultadoEdificio:
    nombre: str
    sismo: dict = field(default_factory=dict)             # "X"/"Y" -> ResultadoSismo
    muros: list[ResultadoMuro] = field(default_factory=list)
    pisos: list[ResultadoPisoDir] = field(default_factory=list)
    no_portantes: list[ResultadoNoPortante] = field(default_factory=list)
    checks_globales: list[Check] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)
    fuera_de_alcance: list[str] = field(default_factory=list)

    def todos_los_checks(self) -> list[Check]:
        cs = list(self.checks_globales)
        for p in self.pisos:
            cs += p.checks
        for m in self.muros:
            cs += m.checks
        for n in self.no_portantes:
            cs += n.checks
        return cs

    @property
    def dcr_max(self) -> float:
        return max((c.dcr for c in self.todos_los_checks() if c.aplica), default=0.0)

    @property
    def estado(self) -> str:
        return ESTADO_CUMPLE if self.dcr_max <= 1.0 else ESTADO_NO_CUMPLE

    def muro(self, nombre: str, piso: Optional[int] = None) -> Optional[ResultadoMuro]:
        for m in self.muros:
            if m.nombre == nombre and (piso is None or m.piso == piso):
                return m
        return None

    def piso_dir(self, piso: int, direccion: str) -> Optional[ResultadoPisoDir]:
        for p in self.pisos:
            if p.piso == piso and p.direccion == direccion:
                return p
        return None


# ------------------------------------------------------------- datos
@dataclass
class Material:
    tipo_unidad: str = "King Kong industrial (arcilla)"
    familia: str = "arcilla"            # arcilla | silico | concreto
    fb: float = mpa(145)                # MPa
    fm: float = mpa(65)
    vm: float = mpa(8.1)
    Em: Optional[float] = None          # None = automático según familia
    Gm: Optional[float] = None
    mortero: str = "P2"                 # P1, P2, NP
    fc: float = mpa(175)                # concreto de confinamiento
    fy: float = mpa(4200)               # 412 MPa
    ft: float = mpa(1.5)                # tracción por flexión (cargas perpendiculares)
    gamma: float = 18.0 * KN_M3         # peso específico de la albañilería (N/mm³)

    @property
    def coef_Em(self) -> float:
        return {"arcilla": 500.0, "silico": 600.0, "concreto": 700.0}.get(self.familia, 500.0)

    @property
    def Em_calc(self) -> float:
        return self.Em if self.Em else self.coef_Em * self.fm

    @property
    def Gm_calc(self) -> float:
        return self.Gm if self.Gm else 0.4 * self.Em_calc

    @property
    def coef_vm(self) -> float:
        return 0.35 if self.familia == "silico" else 0.5

    @staticmethod
    def desde_tabla9(nombre: str) -> "Material":
        t = cargar_json("tabla9.json")
        for u in t["unidades"]:
            if u["nombre"] == nombre:
                return Material(tipo_unidad=nombre, familia=u["familia"], fb=mpa(u["fb"]),
                                fm=mpa(u["fm"]), vm=mpa(u["vm"]))
        raise KeyError(nombre)


@dataclass
class Cargas:
    """Cargas por unidad de área en MPa (N/mm²); 1 kN/m² = 0.001 MPa."""
    peso_losa: float = 3.0 * KPA        # aligerado h = 20 cm ≈ 3.0 kN/m²
    acabados: float = 1.0 * KPA
    tabiqueria: float = 1.0 * KPA
    sobrecarga: float = 2.0 * KPA       # vivienda E.020
    sobrecarga_azotea: float = 1.0 * KPA
    peralte_losa: float = 200.0         # mm (para soleras)

    @property
    def muerta(self) -> float:
        return self.peso_losa + self.acabados + self.tabiqueria


@dataclass
class Muro:
    nombre: str
    piso: int = 1
    direccion: str = "X"                # "X" o "Y"
    L: float = 3000.0                   # mm
    t: float = 130.0                    # espesor efectivo (mm)
    h: float = 2400.0                   # altura libre (mm)
    portante: bool = True
    confinado: bool = True
    x: float = 0.0                      # centro del muro en planta (mm)
    y: float = 0.0
    Pg: Optional[float] = None          # D + 25 % L (N); None = por área tributaria
    Pm: Optional[float] = None          # D + L (N) para el esfuerzo axial
    area_trib: Optional[float] = None   # mm² de losa por nivel
    Lm: Optional[float] = None          # paño mayor (mm); None = L si Nc = 2
    Nc: int = 2                         # número de columnas de confinamiento
    tabique: bool = False               # tiene tabique/alféizar no aislado
    Pt: float = 0.0                     # carga del muro transversal sobre la columna (N)
    Pc_ext: Optional[float] = None      # carga de gravedad de la columna extrema (N)
    Pc_int: Optional[float] = None
    muro_transversal: bool = False      # columna extrema confinada por muro transversal (δ = 1)
    junta_rugosa: bool = False          # μ = 1.0
    aberturas: str = ""
    carga_perpendicular: bool = False   # verificar ante carga perpendicular (Art. 29)
    grupo: str = ""                     # identifica el mismo muro en todos los pisos
    Ve: Optional[float] = None          # sismo moderado (N), si viene de ETABS
    Me: Optional[float] = None          # sismo moderado (N·mm)
    material: str = "albañilería"       # en ETABS: nombre del material

    @property
    def id_grupo(self) -> str:
        return self.grupo or self.nombre

    @property
    def area(self) -> float:
        return self.L * self.t

    @property
    def inercia(self) -> float:
        return self.t * self.L ** 3 / 12.0


@dataclass
class ElementoNoPortante:
    nombre: str
    tipo: str = "tabique"               # tabique | parapeto | cerco | voladizo
    caso: int = 1                       # 1: 4 bordes, 2: 3 bordes, 3: 2 bordes horiz., 4: voladizo
    a: float = 2400.0                   # mm (dimensión crítica según Tabla 12)
    b: float = 3000.0                   # mm (otra dimensión)
    t: float = 130.0                    # espesor efectivo (mm)
    C1: Optional[float] = None          # None = por tipo
    gamma: float = 18.0 * KN_M3         # N/mm³
    e: Optional[float] = None           # espesor bruto para el peso (mm); None = t
    piso: int = 1


@dataclass
class OpcionesAnalisis:
    modo_momento: str = "Ve·h"          # "Ve·h", "voladizo", "factor"
    factor_brazo: float = 1.0           # Me = Ve·factor·h (modo "factor")
    rigidez: str = "voladizo"           # "voladizo" | "empotrado"
    factor_forma: float = 1.2
    torsion: bool = False
    exc_accidental: float = 0.05
    dmin_columna: float = 250.0         # peralte mínimo constructivo de columnas (mm)
    recubrimiento_nucleo: float = 20.0  # mm, al borde del estribo
    peralte_losa: float = 200.0         # mm (sección de la solera = t × peralte)
    modo_etabs: bool = False            # True: Ve, Me, Pg vienen del modelo
    factor_moderado: float = 0.5


@dataclass
class DatosMemoria:
    proyecto: str = "Edificio de albañilería confinada"
    ubicacion: str = ""
    ingeniero: str = ""
    fecha: str = ""
    observaciones: str = ""


@dataclass
class Edificio:
    nombre: str = "Edificio"
    zona: int = 4
    categoria: str = "C"
    suelo: str = "S1"
    N: int = 3
    h_entrepiso: float = 2600.0         # altura de piso a piso (mm)
    Ap: float = 100.0e6                 # área en planta típica (mm²)
    Lx: Optional[float] = None          # dimensiones en planta (mm) para excentricidad/dibujo
    Ly: Optional[float] = None
    Ia: float = 1.0
    Ip: float = 1.0
    Tp: Optional[float] = None          # None = por tabla
    Tl: Optional[float] = None
    U: Optional[float] = None           # None = por categoría
    material: Material = field(default_factory=Material)
    cargas: Cargas = field(default_factory=Cargas)
    muros: list[Muro] = field(default_factory=list)
    no_portantes: list[ElementoNoPortante] = field(default_factory=list)
    opciones: OpcionesAnalisis = field(default_factory=OpcionesAnalisis)
    memoria: DatosMemoria = field(default_factory=DatosMemoria)
    alturas: Optional[list[float]] = None            # alturas de entrepiso por piso (mm); None = h_entrepiso
    pesos_override: Optional[list[float]] = None     # peso sísmico por nivel (N)
    VE_override: Optional[dict] = None               # {"X": [VE piso1..N], "Y": [...]} sismo severo (N)
    derivas: Optional[dict] = None                   # {"X": [..], "Y": [..]} derivas inelásticas
    origen: str = "manual"

    def muros_piso(self, piso: int, direccion: Optional[str] = None) -> list[Muro]:
        return [m for m in self.muros
                if m.piso == piso and (direccion is None or m.direccion == direccion)]

    @property
    def alturas_entrepiso(self) -> list[float]:
        if self.alturas and len(self.alturas) == self.N:
            return list(self.alturas)
        return [self.h_entrepiso] * self.N

    # ------------------------------------------------ guardar / abrir
    def a_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def desde_dict(d: dict) -> "Edificio":
        d = dict(d)
        mat = Material(**_filtrar(Material, d.pop("material", {}) or {}))
        car = Cargas(**_filtrar(Cargas, d.pop("cargas", {}) or {}))
        muros = [Muro(**_filtrar(Muro, m)) for m in d.pop("muros", []) or []]
        nps = [ElementoNoPortante(**_filtrar(ElementoNoPortante, n)) for n in d.pop("no_portantes", []) or []]
        op = OpcionesAnalisis(**_filtrar(OpcionesAnalisis, d.pop("opciones", {}) or {}))
        mem = DatosMemoria(**_filtrar(DatosMemoria, d.pop("memoria", {}) or {}))
        return Edificio(material=mat, cargas=car, muros=muros, no_portantes=nps, opciones=op,
                        memoria=mem, **_filtrar(Edificio, d))

    def guardar(self, ruta: str | Path) -> None:
        Path(ruta).write_text(json.dumps(self.a_dict(), indent=2, ensure_ascii=False), encoding="utf-8")

    @staticmethod
    def abrir(ruta: str | Path) -> "Edificio":
        return Edificio.desde_dict(json.loads(Path(ruta).read_text(encoding="utf-8")))


def _filtrar(cls, d: dict) -> dict:
    nombres = {f.name for f in fields(cls)}
    return {k: v for k, v in d.items() if k in nombres}


def en_unidad(v: float, unidad: str) -> float:
    """Convierte un valor interno (N, N·mm, MPa, mm, mm²) a la unidad de presentación del Check."""
    if unidad == "kN":
        return v / 1000.0
    if unidad == "kN·m":
        return v / 1e6
    return v


def texto_valor(v: float, unidad: str) -> str:
    if v == float("inf"):
        return "∞"
    x = en_unidad(v, unidad)
    if unidad == "MPa":
        return f"{x:.3f} MPa ({kgf_cm2(x):.2f} kgf/cm²)"
    if unidad in ("mm", "mm²"):
        return f"{x:.0f} {unidad}"
    if unidad == "":
        return f"{x:.4f}"
    return f"{x:.2f} {unidad}"
