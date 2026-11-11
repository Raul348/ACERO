# Generador Pro PSST – Norma G.050 + IPERC + EPI/EPC + DXF

Aplicación de escritorio **100 % offline** (un solo archivo: `app_psst_g050_iperc_pro.py`) que, a partir del
presupuesto de obra en Excel, genera:

| Botón | Archivo | Contenido |
|---|---|---|
| 1 (verde) | `.xlsx` | `1_Matriz_IPERC_Obra` (fórmulas IP = A+B+C+D, P×S, nivel y colores condicionales), `2_Presupuesto_Seguridad_G050` (Cantidad × P.U. y factor de reposición vivos), `3_Programa_Capacitaciones`, `4_Formatos_ATS_y_PETAR` (ATS, PETAR y check-lists) |
| 2 (azul) | `.docx` | PSST completo (≈ 30 páginas): 11 capítulos, PETS-01 a PETS-05, plan de emergencias, botiquín G.050, presupuesto |
| 3 (naranja) | `.dxf` | Lámina **SS-01** (A1): layout de instalaciones provisionales, cachaco, entibado, andamio, almacén de materiales peligrosos, señalética NTP 399.010-1, colores RR.SS. y membrete |

Marco normativo: Ley N° 29783, D.S. N° 005-2012-TR, D.S. N° 011-2019-TR, Norma G.050, R.M. N° 050-2013-TR,
R.M. N° 375-2008-TR.

## Cómo funciona

- **Motor IPERC por palabras clave**: 18 familias de actividades (obras provisionales, demoliciones, trazo,
  excavaciones, relleno, concreto, encofrado, acero, albañilería en altura, estructuras metálicas, sanitarias,
  eléctricas, pavimentos, pintura, pisos, carpintería, izaje, señalización/limpieza). Cada partida recibe peligro,
  riesgo, consecuencia, índices A-B-C-D-S, controles por jerarquía, riesgo residual y PETAR/PETS.
  En la pestaña 2, doble clic en una partida para ver el detalle o reasignar su familia.
- **Lector de presupuestos**: detecta solo las columnas Ítem / Descripción / Und. / Metrado / Precio y descarta
  títulos y totales.
- **Desagregado G.050**: 6 subpartidas (EPI, EPC, señalización, capacitación, emergencias, EMO/SCTR) calculadas con
  N° de obreros, staff, plazo, perímetro y trabajos de alto riesgo marcados.

## Instalación y uso

```bash
pip install -r requirements.txt
```

```bash
python app_psst_g050_iperc_pro.py
```

Sin interfaz gráfica (genera los 3 archivos de la obra de ejemplo):

```bash
python app_psst_g050_iperc_pro.py --demo salida_demo
```

## Ejecutable para Windows

```bash
pyinstaller --noconfirm --onefile --windowed --name "Generador_PSST_G050" --collect-all customtkinter app_psst_g050_iperc_pro.py
```

> Los precios unitarios son referenciales (S/ sin IGV) y los teléfonos locales del directorio de emergencia
> deben completarse en obra. El documento debe ser revisado y firmado por el profesional responsable.
