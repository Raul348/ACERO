# ACERO Refuerzo para Revit 2024 – 2027

Complemento (add-in) en **C#** para colocar acero de refuerzo en Revit desde una sola pestaña llamada **ACERO**:

| Botón | Qué arma |
|---|---|
| **Vigas** | Acero corrido superior e inferior con patas en los apoyos, bastones superiores en apoyos e inferiores al centro, estribos cerrados con ganchos de 135° por zonas |
| **Columnas** | Rectangulares (barras por cara, estribo y grapas) **e irregulares**: L, T, cruz, circulares y poligonales (sigue el contorno real) |
| **Zapatas** | Parrilla inferior X/Y con patas y parrilla superior opcional |
| **Losas** | Malla inferior y superior (temperatura) que sigue el contorno real de la losa (Refuerzo de Área) |
| **Muros** | Malla simple o doble, acero vertical con anclaje y empalme, acero horizontal con ganchos |
| **Acero integral** | Menú con las 5 figuras para elegir el elemento |
| **Diámetros** | Crea los diámetros estándar que falten: 6 mm, 8 mm, 3/8", 12 mm, 1/2", 5/8", 3/4", 1", 1 3/8" |

Cada botón tiene su **figurita** en la cinta, y cada ventana muestra la **figura del elemento a escala**.
La figura se redibuja al cambiar cualquier dato: recubrimiento, diámetros, número de barras, separaciones o distribución de estribos.

![Íconos](docs/img/iconos.png)

## Figuras

| Viga | Columna rectangular |
|---|---|
| ![Viga](docs/img/viga.png) | ![Columna](docs/img/columna_rect.png) |

| Columna irregular (L) | Columna circular |
|---|---|
| ![Columna L](docs/img/columna_L.png) | ![Columna circular](docs/img/columna_circ.png) |

| Zapata | Losa de techo |
|---|---|
| ![Zapata](docs/img/zapata.png) | ![Losa](docs/img/losa.png) |

| Muro estructural |
|---|
| ![Muro](docs/img/muro.png) |

## Instalación

### Versiones de Revit

El mismo código se compila para cada versión de Revit, porque Revit cambió de plataforma .NET:

| Revit | Plataforma |
|---|---|
| 2024 | .NET Framework 4.8 |
| 2025 y 2026 | .NET 8 |
| 2027 | .NET 10 |

Las diferencias de la API entre versiones ya están resueltas en el código.
Por ejemplo, Revit 2026 y 2027 reemplazan `RebarHookOrientation` por `BarTerminationsData`.

### Opción A: descargar el complemento ya compilado
1. En GitHub, abra **Actions → Compilar complemento → la última ejecución**.
2. Descargue el artefacto de su versión, por ejemplo `AceroRefuerzo-Revit2025`.
3. Copie su contenido en `%AppData%\Autodesk\Revit\Addins\2025\` para que quede así:
   ```
   Addins\2025\AceroRefuerzo.addin
   Addins\2025\AceroRefuerzo\AceroRefuerzo.dll
   ```
4. Abra Revit y acepte **Cargar siempre** cuando aparezca el aviso de seguridad.

### Opción B: compilar usted mismo
Necesita el [SDK de .NET 8](https://dotnet.microsoft.com/download), y además el SDK de .NET 10 para Revit 2027.
No hace falta tener Revit instalado para compilar: las referencias de la API se descargan de NuGet.

```powershell
.\build.ps1                    # compila 2024, 2025, 2026 y 2027 en .\dist
.\install.ps1                  # copia a las versiones de Revit instaladas
```

Para una sola versión: `.\build.ps1 -Versions 2025` y `.\install.ps1 -Versions 2025`.
También puede abrir `AceroRefuerzo.sln` en Visual Studio 2022. La versión de Revit se elige con la propiedad `RevitVersion`, que por defecto es 2025.

## Uso

1. Seleccione uno o varios elementos, o pulse el botón y selecciónelos; termine con **Finalizar**.
2. Complete los datos. La figura de la izquierda muestra el armado.
3. Pulse **Colocar acero**. Al final verá un resumen con el número de barras, la longitud total y el peso aproximado en kg.

El programa recuerda los últimos valores de cada ventana en `%AppData%\ACERORefuerzo`.
Cada barra queda marcada en *Comentarios* como `ACERO - …`, para filtrarla o tabularla.

### Distribución de estribos
Se escribe como en los planos y se aplica **desde cada extremo** del elemento:

```
1@5, 10@10, R@20          (centímetros)
1@0.05, 10@0.10, R@0.20   (metros: si todas las separaciones son < 1)
```

Significa 1 estribo a 5 cm, 10 a 10 cm y el **R**esto a 20 cm como máximo. Esto se repite desde ambos extremos.
En columnas, los dos extremos son abajo y arriba, así que se generan las zonas de confinamiento.

### Columnas irregulares
Con **Tipo de armado = Automático**, el programa lee la cara inferior de la columna:
- Si es un rectángulo, usa *barras por cara*.
- Si no (L, T, cruz, círculo, polígono), usa *contorno*:
  - coloca una barra en cada esquina y barras intermedias en cada lado, a la **separación máxima** indicada;
  - en secciones circulares reparte uniformemente, con un **número mínimo** de barras;
  - el estribo es cerrado, sigue el contorno (con arcos reales en los círculos) y lleva ganchos de 135°.

## Requisitos en el modelo
- Vigas y muros **rectos**. Las columnas deben ser **verticales**.
- Muros y losas con la opción **Estructural** activada; Revit solo permite armadura en anfitriones estructurales.
- Losas: categoría *Suelos*. Zapatas: categoría *Cimentación estructural*, como ejemplar de familia o losa de cimentación.

## Estructura del código

```
src/Comun/                  Ventana de datos con figura, primitivas de dibujo, unidades y geometría local
src/AceroRefuerzo/
  App.cs                    Pestaña ACERO y botones con sus íconos
  Commands/                 Comandos de Revit y flujo común (selección → ventana → transacción → resumen)
  Modules/                  Un módulo por elemento: Viga, Columna, Zapata, Losa, Muro
  Core/                     Polígonos, distribución de estribos, creación de barras
  UI/                       Figuras del acero, íconos y menú integral
```

Para agregar otro elemento estructural, cree una clase que implemente `IRebarModule` y regístrela en `ModuleRegistry`.

## Limitaciones conocidas
- Vigas curvas, muros curvos y columnas inclinadas no están soportados.
- En columnas irregulares el estribo es uno perimetral; no se generan estribos interiores ni grapas.
- Losas aligeradas (con viguetas): se arma la losa como maciza.
- Las barras que salen del elemento (anclajes y empalmes) quedan alojadas en el mismo elemento.
