# 🧠 Entrenador de Memoria

Aplicación web interactiva para entrenar y mejorar la memoria. No necesita instalación ni conexión a internet.

## Cómo usarla

Abre `index.html` con doble clic (Chrome, Edge, Firefox o Safari). Funciona en ordenador, tablet y móvil.

## Ejercicios

| Ejercicio | Qué entrena | Cómo se juega |
|---|---|---|
| 🃏 **Parejas** | Memoria visual-espacial | Encuentra las parejas de cartas con el menor número de movimientos. |
| 🎵 **Secuencia** | Memoria secuencial y auditiva | Repite una secuencia de luces y sonidos que crece cada ronda (teclas 1–9). |
| 🔢 **Dígitos** | Memoria de trabajo verbal | Recuerda números cada vez más largos. Modo inverso y lectura en voz alta. |
| 🟪 **Cuadrícula** | Memoria espacial | Recuerda qué casillas se iluminaron; la cuadrícula crece. |
| 📝 **Lista de palabras** | Memoria verbal | Estudia una lista y escribe todas las palabras que recuerdes. |
| 🎯 **N-Back** | Memoria de trabajo y concentración | Pulsa (o barra espaciadora) cuando la posición coincida con la de N pasos atrás. |

## Funciones

- **Dificultad adaptativa**: el nivel sube o baja según tu rendimiento (también puedes elegirlo).
- **Entrenamiento diario**: 3 ejercicios aleatorios seguidos.
- **Progreso**: racha de días, récords, gráfico de actividad y evolución de puntuación.
- **Técnicas de memoria**: palacio de la memoria, agrupación, historias, repetición espaciada…
- Sonidos, confeti en los récords, tema claro/oscuro y control por teclado.
- Los datos se guardan en el navegador (`localStorage`) y se pueden exportar a JSON.

## Archivos

- `index.html` — estructura de la página
- `estilos.css` — diseño (adaptable a móvil, modo oscuro)
- `app.js` — lógica de los ejercicios, progreso y estadísticas
