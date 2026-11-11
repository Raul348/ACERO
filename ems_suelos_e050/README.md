# EMS PRO · NTE E.050

Aplicación de escritorio (un solo archivo, 100 % offline) para automatizar Estudios de Mecánica de Suelos según la **NTE E.050**, la **E.030** y la **E.060** del RNE (Perú).

**Archivo principal:** `app_ems_suelos_e050_pro.py`

## Qué hace
- **Capacidad portante:** Terzaghi, Meyerhof y Vesic, con falla general o local (tan φ* = 2/3·tan φ ; c* = 2/3·c), factores de forma y profundidad, corrección por nivel freático y FS = 3.0.
- **Asentamiento elástico:** Se = q·B·(1-μ²)·If/Es. Controla S ≤ 2.54 cm y la distorsión angular α ≤ 1/500.
- **Matriz de diseño:** Df = 1.20, 1.50, 1.80 y 2.00 m × B = 0.60 m (cimiento corrido) y 1.20, 1.50, 2.00 y 2.50 m. Toma la envolvente de la calicata más desfavorable.
- **Clasificación automática:** SUCS (NTP 339.134) y AASHTO (NTP 339.135).
- **Parámetros sísmicos E.030-2018:** Z, U, S, Tp, TL y C.
- **Agresividad química:** sulfatos, cloruros, sales y pH. Recomienda el tipo de cemento y la relación a/c.
- **Empujes:** Ka, Kp, K0 y μ.

## Exportaciones
| Botón | Archivo |
|---|---|
| 1 (azul) | Informe EMS en Word (.docx, ~22 páginas, con perfiles y anexo fotográfico) |
| 2 (naranja) | Lámina EMS-01 en AutoCAD (.dxf): registros por calicata, correlación, cuadro resumen, leyenda SUCS y membrete |
| 3 (verde) | Excel (.xlsx) con 4 hojas: laboratorio y granulometría, capacidad portante, asentamientos, química/sismo/empujes |

## Instalación y uso
```bash
pip install customtkinter openpyxl pandas ezdxf python-docx Pillow
```
```bash
python app_ems_suelos_e050_pro.py
```
Genera el ejemplo completo sin abrir la interfaz:
```bash
python app_ems_suelos_e050_pro.py --demo carpeta_salida
```

## Ejecutable para Windows
```bash
pip install pyinstaller
```
```bash
pyinstaller --onefile --windowed --name EMS_PRO_E050 --collect-all customtkinter --collect-all ezdxf app_ems_suelos_e050_pro.py
```

## Formato del Excel de importación
Usa el botón **Plantilla Excel**. Hoja `Estratigrafia` con las columnas: Calicata, Estrato, Prof_Ini, Prof_Fin, Muestra, SUCS, AASHTO, Grava, Arena, Finos, Humedad, LL, LP, IP, Gamma, Phi, Cohesion, Cu, Cc y Descripcion. Escribe `NP` en los suelos no plásticos. La hoja `Calicatas` es opcional y lleva Este, Norte, Cota y NF.

> El Es y el μ automáticos son valores típicos por tipo SUCS (Bowles). El profesional responsable debe validar todos los resultados.
