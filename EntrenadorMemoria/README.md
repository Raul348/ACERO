# 🧠 Entrenador de Memoria

Entrenamiento de memoria basado en **métodos con eficacia demostrada en estudios científicos**, con una **evaluación objetiva** para comprobar si mejoras.
No necesita instalación ni internet: abre `index.html` en cualquier navegador (ordenador, tablet o móvil).

## Cómo usarlo

1. **Evaluación inicial (≈ 12 min)**: mide tu memoria antes de entrenar (tu línea base).
2. **Entrena a diario (10–20 min)**: repaso espaciado + palacio de la memoria.
3. **Repite la evaluación cada 2 semanas** y compara la tendencia.

## 📋 Evaluación (pre / post)

Protocolo fijo inspirado en pruebas clásicas de la psicología:

| Prueba | Mide |
|---|---|
| Aprendizaje de 15 palabras (3 ensayos) | Memoria verbal y capacidad de aprendizaje |
| Span de dígitos directo e inverso | Memoria a corto plazo y memoria de trabajo |
| Span espacial tipo Corsi | Memoria de trabajo visoespacial |
| Recuerdo diferido | Retención tras varios minutos |

- 6 **listas paralelas** de palabras (una palabra de las mismas 15 categorías en cada lista), para que las evaluaciones sucesivas sean comparables sin repetir palabras.
- Avisa si repites la evaluación demasiado pronto (efecto de práctica).
- Es orientativa: **no es un diagnóstico clínico**.

## ✅ Métodos con evidencia sólida

| Método | Evidencia |
|---|---|
| 🔁 **Repaso espaciado con recuerdo activo** (algoritmo SM-2, tarjetas propias) | Roediger y Karpicke (2006); Karpicke y Roediger (2008); Cepeda et al. (2006); Dunlosky et al. (2013) |
| 🏛️ **Palacio de la memoria** (método de loci con tu propio recorrido) | Dresler et al. (2017), *Neuron* |

## ⚠️ Ejercicios complementarios (evidencia limitada)

Parejas, Secuencia, Dígitos, Cuadrícula, Lista de palabras y N-Back. Mejoran la tarea que se practica, pero los metaanálisis encuentran poca transferencia a la memoria cotidiana (Melby-Lervåg y Hulme, 2013; Simons et al., 2016; Soveri et al., 2017). Se incluyen como calentamiento y para la concentración, y así se indica en el programa.

La página **Evidencia** del programa resume cada estudio con su referencia completa.

## Archivos

- `index.html` — página principal
- `estilos.css` — diseño (adaptable a móvil, modo oscuro)
- `app.js` — núcleo, ejercicios complementarios, progreso y página de evidencia
- `evaluacion.js` — evaluación estandarizada y comparación de resultados
- `metodos.js` — palacio de la memoria y repaso espaciado (SM-2)

Los datos se guardan en el navegador (`localStorage`) y se pueden exportar a JSON desde «Progreso».
