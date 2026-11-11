# Generador Pro de TDR y Requisitos Técnicos Mínimos (OSCE / OECE - SEACE)

Aplicación de escritorio offline (un solo archivo: `app_tdr_osce_pro.py`) que genera:

- **TDR oficial en Word (.docx)** de 20 a 23 páginas: capítulos 1.0 a 8.0, membrete, tablas, anexos.
- **Cuadro de evaluación en Excel (.xlsx)** con fórmulas: puntaje de precio, total, orden de prelación,
  verificación CUMPLE / NO CUMPLE, pagos, penalidades y garantías.

## Instalación y uso

```bash
pip install -r requirements.txt
python app_tdr_osce_pro.py
```

Prueba sin interfaz (genera archivos de ejemplo en la carpeta indicada):

```bash
python app_tdr_osce_pro.py --prueba salida
```

## Crear el ejecutable (Windows)

```bash
pyinstaller --onefile --windowed --name "Generador_TDR_OSCE" --collect-all customtkinter app_tdr_osce_pro.py
```

## Notas normativas

- Marco legal seleccionable: **Ley N° 32069 y D.S. N° 009-2025-EF** (vigente) o **TUO Ley N° 30225 y
  D.S. N° 344-2018-EF** (procesos anteriores).
- Penalidad por mora: `0.10 × Monto / (F × Plazo)`, con F = 0.40 para plazos ≤ 60 días; para plazos > 60 días,
  F = 0.25 (bienes, servicios y consultorías) o F = 0.15 (obras).
- La UIT, los umbrales de procedimiento y los porcentajes de adelantos son **referenciales y editables**
  (pestaña 3). Verifíquelos cada año con la Ley de Presupuesto y las Bases Estándar vigentes.
