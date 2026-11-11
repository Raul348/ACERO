# Diseñador Pro de Cimentaciones – Equipos Vibratorios y Postes/Torres

Aplicación de escritorio (Python, 100 % offline) en un solo archivo: `app_cimentacion_equipos_postes_pro.py`.

- **Módulo A – Bancada antivibratoria** (ACI 351.3R-04, NTE E.050/E.060, ACI 318-19 Cap. 17): regla de masa, borde libre, presiones (≤ 50 % q adm), excentricidad, rigidez Kz (Lysmer/Richart), amortiguamiento, frecuencias natural vertical/horizontal/cabeceo, resonancia (0.70 ≤ r ≤ 1.40), amplitud, aisladores, acero y pernos J.
- **Módulo B – Macizo de poste/torre** (Sulzberger + viento NTE E.020): viento por tramos, Mv, M1, M2, FS ≥ 1.50, giro tan α ≤ 0.01, presiones, pedestal y canastilla.
- **Exporta**: memoria Word (.docx, ~17 págs.), lámina CE-01 AutoCAD (.dxf), hojas Excel (.xlsx) y lotes desde Excel.

```bash
pip install customtkinter openpyxl pandas ezdxf python-docx
```

```bash
python app_cimentacion_equipos_postes_pro.py
```

Sin interfaz (genera los 3 archivos de ejemplo):

```bash
python app_cimentacion_equipos_postes_pro.py --demo salida_demo
```

Ejecutable para Windows:

```bash
pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter --name "DisenadorProCimentaciones" app_cimentacion_equipos_postes_pro.py
```
