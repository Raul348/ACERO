# Validación — ALBAÑILERÍA E.070

Todas las pruebas: `python -m pytest -q tests` (61 pruebas, todas en verde).

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

## 5. Memoria externa: muro MX12, piso 1 (hoja de cálculo E.070 entregada por el usuario)

Datos: f'm = 65, v'm = 8.05, E = 32 500 kgf/cm²; L = 2.60 m, t = 0.23 m, h = 2.40 m; Pm = 16 t, Pg = 14 t,
Ve = 8.72 t, Me = 42.98 t·m; f'c = 210, fy = 4200 kgf/cm²; μ = 1.0; δ = 1.0. Prueba:
`tests/test_validacion_mx12.py` (tabla: `PYTHONPATH=. python tests/test_validacion_mx12.py`).

| Magnitud | Memoria | Programa | Dif. % |
|---|---|---|---|
| σm (t/m²) | 26.76 | 26.756 | −0.02 |
| α | 0.53 | 0.528 | −0.47 (redondeo de la memoria) |
| Vm (t) | 15.92 | 15.917 | −0.02 |
| 0.55·Vm (t) | 8.75 | 8.754 | +0.05 |
| Vm1/Ve1 | 2.00 | 2.00 | 0.00 |
| Vu / Mu | 17.44 t / 85.96 t·m | 17.44 / 85.96 | 0.00 |
| M = Mu − ½·Vm·h (t·m) | 66.85 | 66.86 | +0.01 |
| F / Pc (t) | 25.71 / 7.00 | 25.715 / 7.00 | +0.02 / 0.00 |
| Vc / T / C (t) | 7.96 / 18.713 / 32.71 | 7.958 / 18.715 / 32.715 | −0.02 / +0.01 / +0.02 |
| Acf (cm²) | 222.93 | 222.92 | 0.00 |
| As = Asf + Ast (cm²) | 7.47 | 7.472 | +0.02 |
| Ac / An núcleo (cm²) | 575 / 399 | 575 / 399 | 0.00 |
| s1 / s2 con Ø1/4" (cm) | 5.09 / 5.61 | 5.091 / 5.614 | +0.02 / +0.07 |
| Ts (t) / As solera (cm²) | 7.96 / 2.11 | 7.958 / 2.105 | −0.02 / −0.22 |

Decisiones iguales: no hay fisuración (8.72 ≤ 8.75 t), muro agrietado, requiere refuerzo horizontal
(Vu = 17.44 > Vm = 15.92 t), columna 25 × 23 cm (gobierna 15·t = 345 cm²), solera 4 Ø 3/8".

### Observaciones a la memoria (argumentadas)

1. **Acero vertical insuficiente**: se requieren As = 7.47 cm² y se adopta 10 Ø 3/8" = 7.10 cm² (−5 %). Además
   10 barras en un núcleo de 19 × 21 cm es poco constructivo. El programa propone 4 Ø 5/8" (7.96 cm²);
   6 Ø 1/2" (7.74 cm²) también cumpliría.
2. **Estribos de la zona confinada**: la propia memoria calcula s = mín(s1…s4) = 5.09 cm, pero en el detalle
   final pone 1 @ 0.05, 4 @ 0.10, resto @ 0.25 (10 cm en la zona de 45 cm). Eso no cumple el s calculado. Lo
   coherente es [] Ø 1/4": 1 @ 5, 8 @ 5 cm en los 45 cm de cada extremo, resto @ 25 (es lo que da el programa).
3. **An calculado con el acero adoptado**: 101.85 cm² resulta de usar As = 7.10 cm² (el insuficiente). Con el
   As requerido (7.47 cm²) An = 93.5 cm². Como An disponible = 399 cm², no gobierna.
4. **δ = 1.0 y T sin muro transversal no son coherentes**: el diseño usa T1 (columna sin muro transversal) pero
   δ = 1.0 (columna confinada por muros transversales). Si no hay muro transversal, δ = 0.8 y An = 115 cm²,
   que sigue siendo menor que 399 cm²: no cambia el resultado.
5. **Pt = 0.25·Pg = 4 t** es una suposición. En la E.070, Pt es la carga tributaria del muro transversal
   (≈ Lt·Pg/L). Sólo afecta a T2, que no se usó para diseñar.
6. **Errata en la fórmula**: dice M = Mu1 − ½·Mm1·h, debe decir ½·Vm1·h. El número (66.85) sí usa Vm1.
7. **Vu y Mu del cuadro inicial** (26.16 t, 128.93 t·m = 3·Ve, 3·Me) no son los que se usan en el diseño
   (17.44 t, 85.96 t·m con el factor 2.0). Conviene corregir el cuadro para no confundir.
8. **s3 con d = 22.5 cm** en vez del peralte de 25 cm (d/4 = 6.25 cm). No gobierna.
9. **Solera de 25 cm de ancho** en un muro de 23 cm. Lo usual es t × peralte (23 × 20). As mín = 2.50 cm² frente
   a 2.30 cm²: 4 Ø 3/8" cumple en ambos casos.
10. **Fisuración al límite**: DCR = 8.72/8.75 = 0.996. Cumple, pero sin margen (Vm/Ve = 1.83).

### Ajustes hechos en el programa por esta validación
- Refuerzo horizontal en el piso 1 cuando Vu > Vm porque el factor Vm1/Ve1 se llevó al mínimo de 2 (antes sólo
  por σm ≥ 0.05·f'm o N > 3). Ahora coincide con la memoria.
- Se agregó la barra Ø 1/4" (0.32 cm²) a los estribos comerciales (6 mm, 1/4", 8 mm, 3/8").
