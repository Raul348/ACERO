"""Conexión COM con ETABS 22 (instancia oculta propia; nunca GetObject).

Reglas:
- Se crea SIEMPRE una instancia nueva: helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject").
- Nunca se abre el .EDB original: se copia a tests/salida/real/ y se trabaja sobre la copia.
- Unidades N, mm, C: SetPresentUnits(9).
- Al terminar: ApplicationExit(False).
Sólo funciona en Windows con ETABS 22 instalado y comtypes.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Any

SALIDA_REAL = Path(__file__).resolve().parents[2] / "tests" / "salida" / "real"
UNIDADES_N_MM_C = 9


class ErrorEtabs(RuntimeError):
    pass


def copiar_modelo(ruta_edb: str | Path, destino: Path = SALIDA_REAL) -> Path:
    """Copia el .EDB del usuario (y archivos asociados con el mismo nombre) a tests/salida/real."""
    src = Path(ruta_edb).resolve()
    if not src.exists():
        raise ErrorEtabs(f"No existe el archivo: {src}")
    destino.mkdir(parents=True, exist_ok=True)
    dst = destino / src.name
    if dst.resolve() == src:
        return dst
    shutil.copy2(src, dst)
    return dst


def filas_tabla(resultado: Any) -> list[dict]:
    """Convierte el resultado de GetTableForDisplayArray (comtypes) en lista de dicts.

    Orden de salida en comtypes: [FieldKeyList, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret].
    TableData viene aplanado (NumberRecords × len(FieldsKeysIncluded)).
    """
    if resultado is None:
        return []
    r = list(resultado)
    if len(r) < 6:
        raise ErrorEtabs(f"Resultado de tabla inesperado ({len(r)} elementos)")
    ret = r[5]
    if ret not in (0, None):
        return []
    campos = [str(c) for c in (r[2] or [])]
    n = int(r[3] or 0)
    datos = list(r[4] or [])
    nc = len(campos)
    if nc == 0 or n == 0:
        return []
    return [dict(zip(campos, datos[i * nc:(i + 1) * nc])) for i in range(n)]


class ClienteEtabs:
    """Instancia oculta de ETABS creada por este programa."""

    def __init__(self):
        self.etabs = None
        self.m = None
        self.ruta_modelo: Path | None = None

    # ------------------------------------------------------------ ciclo de vida
    def iniciar(self, visible: bool = False) -> None:
        if sys.platform != "win32":
            raise ErrorEtabs("La conexión con ETABS requiere Windows con ETABS 22 instalado.")
        try:
            import comtypes.client  # type: ignore
        except ImportError as e:
            raise ErrorEtabs("Falta el paquete comtypes (pip install comtypes).") from e
        try:
            helper = comtypes.client.CreateObject("ETABSv1.Helper")
            import comtypes.gen.ETABSv1 as ETABSv1  # type: ignore
            helper = helper.QueryInterface(ETABSv1.cHelper)
            self.etabs = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
            self.etabs.ApplicationStart()
        except Exception as e:      # noqa: BLE001
            raise ErrorEtabs(f"No se pudo iniciar ETABS: {e}") from e
        if not visible:
            try:
                self.etabs.Hide()
            except Exception:       # noqa: BLE001
                pass
        self.m = self.etabs.SapModel
        self.m.InitializeNewModel(UNIDADES_N_MM_C)

    def abrir(self, ruta_edb: str | Path, copiar: bool = True) -> Path:
        if self.m is None:
            self.iniciar()
        ruta = copiar_modelo(ruta_edb) if copiar else Path(ruta_edb).resolve()
        ret = self.m.File.OpenFile(str(ruta.resolve()))
        if ret not in (0, None):
            raise ErrorEtabs(f"ETABS no pudo abrir {ruta} (ret = {ret})")
        self.m.SetPresentUnits(UNIDADES_N_MM_C)
        self.ruta_modelo = ruta
        return ruta

    def nuevo_modelo(self) -> None:
        if self.m is None:
            self.iniciar()
        self.m.InitializeNewModel(UNIDADES_N_MM_C)
        self.m.File.NewBlank()
        self.m.SetPresentUnits(UNIDADES_N_MM_C)

    def cerrar(self) -> None:
        if self.etabs is not None:
            try:
                self.etabs.ApplicationExit(False)
            except Exception:       # noqa: BLE001
                pass
        self.etabs = None
        self.m = None

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.cerrar()

    # ------------------------------------------------------------ utilidades
    def guardar(self, ruta: str | Path | None = None) -> None:
        r = Path(ruta) if ruta else self.ruta_modelo
        if r is None:
            raise ErrorEtabs("No hay ruta para guardar.")
        if self.ruta_modelo is not None and r.resolve() != self.ruta_modelo.resolve() \
                and SALIDA_REAL not in r.resolve().parents:
            raise ErrorEtabs("Por seguridad sólo se guardan copias dentro de tests/salida.")
        self.m.File.Save(str(r.resolve()))

    def bloqueado(self) -> bool:
        return bool(self.m.GetModelIsLocked())

    def analizar(self) -> None:
        ret = self.m.Analyze.RunAnalysis()
        if ret not in (0, None):
            raise ErrorEtabs(f"RunAnalysis devolvió {ret}")

    @staticmethod
    def _lista(res, idx_lista: int = 1) -> list[str]:
        r = list(res)
        for x in r:
            if isinstance(x, (list, tuple)) and x and isinstance(x[0], str):
                return [str(v) for v in x]
        return [str(v) for v in (r[idx_lista] or [])] if len(r) > idx_lista else []

    def casos(self) -> list[str]:
        return self._lista(self.m.LoadCases.GetNameList())

    def combos(self) -> list[str]:
        return self._lista(self.m.RespCombo.GetNameList())

    def patrones(self) -> list[tuple[str, int]]:
        nombres = self._lista(self.m.LoadPatterns.GetNameList())
        out = []
        for n in nombres:
            try:
                r = list(self.m.LoadPatterns.GetLoadType(n))
                out.append((n, int(r[0])))
            except Exception:       # noqa: BLE001
                out.append((n, 0))
        return out

    def pisos(self) -> list[tuple[str, float, float]]:
        """[(nombre, elevación, altura)] de abajo hacia arriba (sin la base)."""
        r = list(self.m.Story.GetStories())
        # [NumberStories, StoryNames, StoryElevations, StoryHeights, IsMasterStory, SimilarToStory,
        #  SpliceAbove, SpliceHeight, ret]
        n = int(r[0])
        nombres, elev, alt = list(r[1]), list(r[2]), list(r[3])
        out = [(str(nombres[i]), float(elev[i]), float(alt[i])) for i in range(n)]
        out = [p for p in out if p[2] > 0]
        out.sort(key=lambda p: p[1])
        return out

    def tabla(self, nombre: str, casos: list[str] | None = None, combos: list[str] | None = None) -> list[dict]:
        dt = self.m.DatabaseTables
        if casos is not None:
            dt.SetLoadCasesSelectedForDisplay(casos)
        if combos is not None:
            dt.SetLoadCombinationsSelectedForDisplay(combos)
        res = dt.GetTableForDisplayArray(nombre, [], "All", 0, [], 0, [])
        return filas_tabla(res)

    def tablas_disponibles(self) -> list[str]:
        r = list(self.m.DatabaseTables.GetAvailableTables())
        for x in r:
            if isinstance(x, (list, tuple)) and x and isinstance(x[0], str):
                return [str(v) for v in x]
        return []

    def material(self, nombre: str) -> dict:
        tipos = {1: "Steel", 2: "Concrete", 3: "NoDesign", 4: "Aluminum", 5: "ColdFormed", 6: "Rebar",
                 7: "Tendon", 8: "Masonry"}
        try:
            r = list(self.m.PropMaterial.GetMaterial(nombre))
            t = int(r[0])
        except Exception:           # noqa: BLE001
            t = 0
        d = {"tipo": tipos.get(t, str(t))}
        try:
            r = list(self.m.PropMaterial.GetMPIsotropic(nombre))
            d["E"] = float(r[0])
        except Exception:           # noqa: BLE001
            pass
        return d

    def materiales(self) -> dict:
        return {n: self.material(n) for n in self._lista(self.m.PropMaterial.GetNameList())}

    def crear_combo(self, nombre: str, factores: dict[str, float]) -> None:
        rc = self.m.RespCombo
        if nombre not in self.combos():
            rc.Add(nombre, 0)
        for caso, f in factores.items():
            rc.SetCaseList(nombre, 0, caso, f)

    def crear_combo_gravedad(self, nombre: str = "ALB_D+0.25L", factor_vivo: float = 0.25) -> str:
        """Crea D + 0.25 L con los patrones de tipo muerta (1, 2) y viva (3, 4, 11)."""
        f = {}
        for n, t in self.patrones():
            if t in (1, 2):
                f[n] = 1.0
            elif t in (3, 4, 11):
                f[n] = factor_vivo
        if not f:
            raise ErrorEtabs("No se encontraron patrones de carga muerta/viva para crear el combo.")
        self.crear_combo(nombre, f)
        return nombre
