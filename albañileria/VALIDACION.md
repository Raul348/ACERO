# Validación — ALBAÑILERÍA E.070

Todas las pruebas: `python -m pytest -q tests` (39 pruebas, todas en verde).

## 1. Números calculados a mano (diferencia 0 %)

| Verificación | Dato de la prueba | Resultado a mano | Programa | Dif. |
|---|---|---|---|---|
| Densidad (Art. 19.2.b) | ΣL·t = 1.56 m², Ap = 60 m², Z·U·S·N = 0.45·1·1·3 | 0.02600 / 0.02411 | igual | 0 % |
| σm (Art. 19.1.b) | Pm = 200 kN, L = 3 m, t = 13 cm, h = 2.4 m, f'm = 65 kgf/cm² | σm = 0.5128 MPa; Fa = 0.9201 MPa | igual | 0 % |
| α (Art. 26.3) | Ve = 50 kN, L = 3 m, Me = 150 / 300 / 600 kN·m | 1.0 / 0.5 / 0.25 → 1/3 | igual | 0 % |
| Vm (Art. 26.3) | v'm = 8.1 kgf/cm², α = 1, t = 130, L = 3000, Pg = 100 kN | 177.90 kN | 177.90 kN | 0 % |
| 0.55·Vm (Art. 26.2) | — | 97.84 kN | 97.84 kN | 0 % |
| Vm1/Ve1 (Art. 27) | 200/60, 90/60, 150/60 | 3.0 / 2.0 / 2.5 | igual | 0 % |
| Tabla 11, columna extrema | Vm = 180 kN, Mu = 375 kN·m, Pg = 100 kN, L = 3 m, h = 2.4 m | M = 159 kN·m, F = 53 kN, T = 3 kN, C = 103 kN, Vc = 90 kN | igual | 0 % |
| Tabla 11, columna interior | L = 6 m, Nc = 3, Vm = 300 kN, Pg = 200 kN | Vc = 37.5 kN, T = 20 kN, C = 40 kN | igual | 0 % |
| Acf, Asf + Ast, An (Art. 27.3) | Vc = 90 kN, T = 3 kN, C = 103 kN | Acf = 308.5 cm², As = 3.30 cm², An = 12.95 cm² | igual | 0 % |
| Solera (Art. 27) | Ts = 90 kN | As = 2.43 cm² → 4 Ø 3/8" | igual | 0 % |
| w (Art. 29) | Z = 0.45, U = 1, C1 = 2, γ = 18 kN/m³, e = 0.13 m | 1.6848 kN/m² | igual | 0 % |
| m (Tabla 12) | 4 bordes, b/a = 1.25 | 0.0659 (interpolado) | igual | 0 % |
| Ms, fm | a = 2.4 m | 0.6395 kN·m/m; 0.2270 MPa | igual | 0 % |
| E.030 | P = 3000 kN, 3 pisos de 2.6 m, Z4-S1-C | T = 0.13 s, C = 2.5, V = 1125 kN, Fi = V/6, 2V/6, 3V/6 | igual | 0 % |
| Rigidez | L = 3 m, t = 13 cm, h = 2.4 m, Em = 500 f'm | K = 93 160 N/mm | igual | 0 % |

## 2. Coherencia modo manual ↔ modo ETABS

`tests/data/snapshot_ejemplo.json` tiene el formato de lo que se lee de ETABS (P con signo de ETABS, V2 y M3
del sismo severo, cortantes de piso, derivas, masas, más una placa de concreto). Al verificarlo se obtienen
exactamente los mismos Pg, Ve, Me, Vm, VE y derivas que en el cálculo manual del mismo edificio
(diferencia 0 %, tolerancia 1e-9), y la placa se lista pero no se verifica.

## 3. Prueba en vivo con ETABS

`tools/test_etabs_live.py` crea un edificio de 2 pisos con 4 muros, lo analiza en una instancia oculta y
compara Pg (con el área tributaria a mano), ΣV2 de los piers con el cortante de piso y α.
**No se pudo ejecutar en este entorno (Linux, sin ETABS)**: debe correrse en Windows con ETABS 22.

## 4. Ejemplo clásico (San Bartolomé / comentarios de la E.070)

**PENDIENTE.** No se pudieron obtener de una fuente pública verificable los números del ejemplo del libro de
Á. San Bartolomé (PUCP) ni de los comentarios de la E.070, y no se inventaron cifras. Para completar esta
validación basta con los datos del ejemplo (geometría y cargas de los muros, Pg, Ve, Me, Vm, ΣVm/VE, fuerzas y
refuerzo de columnas y soleras): se cargan como edificio en el programa y se agrega una prueba que reporte las
diferencias en porcentaje por magnitud.
