# ALBAÑILERÍA E.070

Programa de escritorio para Windows, de uso personal, para la **verificación y el diseño de edificaciones de
albañilería confinada según la NTE E.070 Albañilería** (Reglamento Nacional de Edificaciones del Perú), con
análisis sísmico **E.030 (2018)**, conexión opcional a **ETABS 22** y **memoria de cálculo en Word** en estilo de
cálculo manual (ecuaciones de Word OMML, sustitución numérica y resultado).

Es el tercero de la familia (acero E.090/AISC con SAP2000 y concreto E.030/E.060/ACI 318 con ETABS) y usa las
mismas estructuras `Step` / `Check`, el mismo tema oscuro y el mismo formato de memoria.

> **Programa de apoyo.** El ingeniero responsable debe revisar los datos y los resultados antes de usarlos en
> planos o expedientes. Lo que no está implementado se marca como **FUERA DE ALCANCE** en los resultados y en la
> memoria; nunca se omite en silencio.

## Instalación y uso

1. Python 3.12 (64 bits) y las dependencias:

```bash
pip install -r requirements.txt
```

2. Doble clic en `ejecutar.bat` (o `python main.py`).
3. Se abre con un **edificio de ejemplo** (3 pisos, Lima). Presione **Verificar (F5)** y luego **Memoria Word**.

Proyectos: **Guardar/Abrir proyecto** escribe un JSON con todo lo ingresado.
Muros: **Importar/Exportar CSV** (separador `;`, unidades de pantalla: m, cm, kN, m²).

## Pestañas

| Pestaña | Contenido |
|---|---|
| 1 Edificio | Z, U, S, Tp/Tl (automáticos por tabla o editables), N, altura de entrepiso, Ap, Lx/Ly, irregularidad (Ia, Ip; R0 = 3 fijo), unidad de albañilería (Tabla 9) y materiales (f'b, f'm, v'm, Em, Gm, mortero, f'c, fy, f't, γ), cargas y opciones de análisis |
| 2 Muros | Tabla editable (agregar, quitar, duplicar, copiar piso 1 a todos, CSV), planta esquemática por piso coloreada por estado y esquema del muro seleccionado con columnas y solera |
| 3 Resultados | Por muro: σm, Ve, Me, α, Vm, 0.55·Vm, DCR de fisuración, Vu, Mu, tipo de diseño, DCR y estado; resumen por piso y dirección (densidad, ΣVm vs VE, deriva, muros agrietados); al seleccionar, los pasos de cada verificación (símbolo, fórmula, sustitución, valor, descripción) |
| 4 Columnas y soleras | Pc, Vc, T, C, Acf, An, sección adoptada, refuerzo vertical propuesto (barras comerciales), estribos y solera |
| 5 Muros no portantes | Tabiques, parapetos y cercos (Tabla 12): w, m, Ms, fm vs f't y arriostres propuestos |
| 6 ETABS | Conectar (instancia oculta, copia del .EDB), leer modelo, elegir casos de sismo y combos de gravedad, clasificar muros, verificar, CSV, memoria automática y **copiloto** de comandos |
| 7 Datos de la memoria | Proyecto, ubicación, ingeniero, fecha, observaciones |

A la derecha, un panel permanente muestra el estado del edificio, el resumen por piso, los muros que no
cumplen, las advertencias y la lista de alcances FUERA DE ALCANCE.

## Motor de cálculo (independiente de la interfaz y de ETABS)

```python
from albanileria.e070 import verificar_edificio
res = verificar_edificio(edificio)      # -> ResultadoEdificio
```

Unidades internas: **N, mm, MPa, N·mm**. En pantalla y memoria: kN, kN·m, MPa y kgf/cm² entre paréntesis
(1 MPa = 10.197 kgf/cm²).

| Módulo | Qué verifica | Referencia |
|---|---|---|
| `e070/materiales.py` | Tabla 9, Em = 500/600/700·f'm, Gm = 0.4·Em, f'c ≥ 175 kgf/cm², mortero | E.070 Cap. 5, Art. 24.7 |
| `e070/requisitos.py` | Espesor t ≥ h/20 (h/25 zona 1), densidad ΣL·t/Ap ≥ ZUSN/56, σm ≤ 0.2·f'm·[1−(h/35t)²] ≤ 0.15·f'm, separación de columnas ≤ 2h ≤ 5 m, avisos de alféizares/tabiques | Art. 19, 20 |
| `e030/` | Z, U, S, Tp, Tl, C, R = 3·Ia·Ip, T = hn/60, C/R ≥ 0.11, P = D + %L, V = ZUCS·P/R, Fi = αi·V, derivas (0.75·R) ≤ 0.005 | E.030 2018 |
| `e070/rigidez.py` | K = 1/(h³/(3EmI) + f·h/(GmA)) (o doble empotramiento), reparto por rigidez, torsión opcional (excentricidad real + 0.05·B) | — |
| `e070/corte.py` | α = Ve·L/Me (1/3…1), Vm = 0.5·v'm·α·t·L + 0.23·Pg (0.35 sílico-calcáreo), Ve ≤ 0.55·Vm, ΣVm ≥ VE, ΣVm ≥ 3·VE elástico | Art. 26 |
| `e070/severo.py` | 2 ≤ Vm1/Ve1 ≤ 3, Vu, Mu, muros agrietados, refuerzo horizontal (Vu ≥ Vm, σm ≥ 0.05 f'm, N > 3) | Art. 27 |
| `e070/confinamiento.py` | Tabla 11 (columna extrema e interior), Acf = Vc/(0.2·f'c·φ), Asf + Ast, An, estribos s1…s4, sección y barras comerciales; soleras Ts = Vm·Lm/(2L); pisos no agrietados (F = Mu/L) y refuerzo mínimo | Art. 27, Tabla 11 |
| `e070/perpendicular.py` | w = 0.8·Z·U·C1·γ·e, Ms = m·w·a² (Tabla 12), fm = 6Ms/t² ≤ f't, arriostres; muros portantes ante carga perpendicular (fm − fa ≤ f't, fa + fm ≤ 0.25 f'm) | Art. 29, Tabla 12 |
| `e070/verificar.py` | Orquesta todo | — |

Tablas editables en `data/` (JSON): `tabla9.json`, `tabla11.json` (coeficientes), `tabla12.json`,
`barras.json`, `e030.json` (incluye los C1 de elementos no estructurales).

### Criterios a revisar por el ingeniero

- **C1 de elementos no estructurales**: por defecto 3.0 parapetos, 2.0 tabiques, 1.3 cercos (editables). La
  fórmula w = 0.8·Z·U·C1·γ·e de la E.070 se calibró con la E.030 de 2003 (C1 = 1.3 / 0.9); con los C1 de la
  E.030 2018 la carga sale bastante mayor.
- **Tabla 9**: los valores por defecto deben confirmarse con la norma y, de preferencia, reemplazarse por
  ensayos de prismas y muretes. El bloque de concreto tipo P está separado por f'b.
- **Pc** de las columnas: por defecto, la mitad de la carga del paño adyacente; agregue la carga de los muros
  transversales (campo Pc ext en el proyecto JSON).
- **Me en modo manual**: por defecto Me = Ve·h; opciones "voladizo acumulado" y "factor de brazo".
- **Deriva en modo manual**: aproximada con muros en voladizo por piso (no incluye la flexión global). La
  deriva confiable es la del modelo ETABS.
- Factor de carga 1.25 en el diseño de arriostres de tabiques (editable en `perpendicular.py`).

## Conexión con ETABS 22

- Instancia **oculta y propia**: `ETABSv1.Helper` → `CreateObjectProgID("CSI.ETABS.API.ETABSObject")` →
  `ApplicationStart()`; nunca `GetObject`. Unidades `SetPresentUnits(9)` (N, mm, C). `ApplicationExit(False)` al cerrar.
- **Nunca** se abre ni se modifica el .EDB original: se copia a `tests/salida/real/` y se trabaja sobre la copia.
- Tablas: `Pier Section Properties`, `Pier Forces`, `Story Forces`, `Story Drifts`, `Mass Summary by Story`,
  `Modal Participating Mass Ratios` con `GetTableForDisplayArray` (orden comtypes:
  `[FieldKeyList, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret]`).
- Signos: Pg = −P (ETABS P > 0 = tracción). Ve = |V2|·0.5, Me = |M3|·0.5 del caso de sismo (severo) en la base
  del pier; casilla para indicar que el caso ya es el sismo moderado. Pg del combo D + 0.25 L (se puede crear
  en la copia).
- Clasificación: material Masonry, o nombre con "ALB", "MURO", "KK"…, o lista marcada por el usuario. Las placas
  de concreto se listan pero no se verifican.
- Todo lo leído se guarda en un **snapshot JSON** para trabajar y probar sin ETABS.
- El COM corre en un **QThread** dedicado (`TrabajadorEtabs`); la interfaz sólo se actualiza en slots del hilo
  principal mediante señales.
- Copiloto: `revisar edificio`, `muro <nombre> [piso]`, `piso <n>`, `alternativas [muro]` (aumentar t,
  alargar, cambiar unidad, agregar muro, evaluadas con el motor), `aplicar <n>` (cambio de espesor en la copia),
  `deshacer`, `reanalizar`.

Herramientas: `tools/explorar_tablas.py` (lista tablas y campos de un .EDB) y `tools/test_etabs_live.py`
(crea un edificio pequeño en una instancia oculta, lo analiza y compara Pg, Ve, Me).

## Estructura

```
albañileria/
  main.py, ejecutar.bat, requirements.txt
  albanileria/
    models.py                 Edificio, Material, Muro, Cargas, Step, Check, ResultadoMuro, ResultadoEdificio…
    e030/                     norma.py, sismo.py
    e070/                     materiales, requisitos, rigidez, corte, severo, confinamiento, perpendicular, verificar
    etabs/                    client.py, reader.py, clasificar.py, batch.py, alternativas.py, model.py
    reports/                  omml.py, word.py
    ui/                       theme.py, main_window.py, widgets.py, etabs_window.py
    ejemplos/                 edificio de ejemplo
    io_csv.py
  data/                       tablas de la norma (JSON)
  tools/                      test_etabs_live.py, explorar_tablas.py, generar_snapshot_ejemplo.py
  tests/                      pytest; tests/data (snapshot de ejemplo); tests/salida (ignorado en git)
```

## Pruebas

```bash
python -m pytest -q tests
```

Incluye números calculados a mano para densidad, σm, α, Vm, 0.55·Vm, factor Vm1/Ve1, Tabla 11 (T, C, Vc),
Acf, As, An, Ts, w, m, Ms, fm, cortante basal y distribución E.030; el edificio de ejemplo completo; la
lectura desde un snapshot ETABS (reproduce exactamente el cálculo manual) y la memoria Word.
Ver la sección de validación en `VALIDACION.md`.
