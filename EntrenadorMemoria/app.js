'use strict';

/* =====================================================================
   Entrenador de Memoria
   Aplicación web sin dependencias. Abre index.html en cualquier navegador.
   ===================================================================== */

/* ---------- Utilidades ---------- */
const $ = (sel, raiz = document) => raiz.querySelector(sel);
const $$ = (sel, raiz = document) => [...raiz.querySelectorAll(sel)];
const azar = (n) => Math.floor(Math.random() * n);
const limitar = (v, min, max) => Math.max(min, Math.min(max, v));
const barajar = (arr) => {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = azar(i + 1);
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
};
const normalizar = (t) => t.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
const fechaISO = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const fmtTiempo = (s) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
const escapar = (t) => String(t).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function nodo(html) {
  const t = document.createElement('template');
  t.innerHTML = html.trim();
  return t.content.firstElementChild;
}

/* ---------- Control de la sesión de juego (temporizadores y teclado) ----------
   Todo temporizador pasa por aquí para poder cancelarlo al salir de un juego. */
const Sesion = {
  timers: new Set(),
  tecla: null,
  limpiar() {
    this.timers.forEach((t) => { clearTimeout(t); clearInterval(t); });
    this.timers.clear();
    this.tecla = null;
    if ('speechSynthesis' in window) speechSynthesis.cancel();
  },
  luego(fn, ms) {
    const id = setTimeout(() => { this.timers.delete(id); fn(); }, ms);
    this.timers.add(id);
    return id;
  },
  cada(fn, ms) {
    const id = setInterval(fn, ms);
    this.timers.add(id);
    return id;
  },
  parar(id) {
    clearTimeout(id);
    clearInterval(id);
    this.timers.delete(id);
  },
  esperar(ms) {
    return new Promise((r) => this.luego(r, ms));
  },
};
document.addEventListener('keydown', (e) => { if (Sesion.tecla) Sesion.tecla(e); });

/* ---------- Almacenamiento local ---------- */
const Datos = {
  clave: 'entrenadorMemoria.v1',
  d: null,
  base() {
    return { historial: [], mejores: {}, niveles: {}, racha: { ultimo: null, dias: 0, maxima: 0 }, sonido: true, tema: null };
  },
  cargar() {
    try { this.d = JSON.parse(localStorage.getItem(this.clave)); } catch (e) { this.d = null; }
    this.d = Object.assign(this.base(), this.d || {});
  },
  guardar() {
    try { localStorage.setItem(this.clave, JSON.stringify(this.d)); } catch (e) { /* almacenamiento no disponible */ }
  },
  nivel(juego) {
    return this.d.niveles[juego] || 1;
  },
  registrar(juego, puntos, nivel, nivelSig, segundos) {
    const ahora = new Date();
    const dia = fechaISO(ahora);
    this.d.historial.push({ juego, puntos, nivel, fecha: ahora.toISOString(), dia, segundos });
    if (this.d.historial.length > 1000) this.d.historial.splice(0, this.d.historial.length - 1000);

    const anterior = this.d.mejores[juego] || 0;
    const record = puntos > anterior;
    if (record) this.d.mejores[juego] = puntos;
    this.d.niveles[juego] = nivelSig;

    const r = this.d.racha;
    if (r.ultimo !== dia) {
      const ayer = new Date(ahora);
      ayer.setDate(ayer.getDate() - 1);
      r.dias = r.ultimo === fechaISO(ayer) ? r.dias + 1 : 1;
      r.ultimo = dia;
      r.maxima = Math.max(r.maxima || 0, r.dias);
    }
    this.guardar();
    return { record: record && anterior > 0, primero: anterior === 0 };
  },
  rachaActual() {
    const r = this.d.racha;
    if (!r.ultimo) return 0;
    const ayer = new Date();
    ayer.setDate(ayer.getDate() - 1);
    return r.ultimo === fechaISO() || r.ultimo === fechaISO(ayer) ? r.dias : 0;
  },
};

/* ---------- Sonido (Web Audio) y voz ---------- */
const Sonido = {
  ctx: null,
  tono(freq = 440, dur = 0.15, tipo = 'sine', vol = 0.15) {
    if (!Datos.d.sonido) return;
    try {
      this.ctx = this.ctx || new (window.AudioContext || window.webkitAudioContext)();
      const t = this.ctx.currentTime;
      const o = this.ctx.createOscillator();
      const g = this.ctx.createGain();
      o.type = tipo;
      o.frequency.value = freq;
      g.gain.setValueAtTime(vol, t);
      g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
      o.connect(g).connect(this.ctx.destination);
      o.start(t);
      o.stop(t + dur + 0.02);
    } catch (e) { /* sin audio */ }
  },
  clic() { this.tono(520, 0.06, 'triangle', 0.08); },
  ok() { this.tono(660, 0.12); setTimeout(() => this.tono(880, 0.16), 90); },
  suave() { this.tono(300, 0.12, 'triangle', 0.07); },
  error() { this.tono(160, 0.35, 'square', 0.06); },
  nota(i, dur = 0.3) {
    const escala = [261.6, 329.6, 392.0, 523.3, 440.0, 349.2, 293.7, 587.3, 659.3];
    this.tono(escala[i % escala.length], dur, 'sine', 0.18);
  },
  fanfarria() { [523, 659, 784, 1047].forEach((f, i) => setTimeout(() => this.tono(f, 0.18), i * 110)); },
};

function hablar(texto) {
  if (!Datos.d.sonido || !('speechSynthesis' in window)) return;
  const u = new SpeechSynthesisUtterance(texto);
  u.lang = 'es-ES';
  u.rate = 1.1;
  speechSynthesis.speak(u);
}

function confeti() {
  const capa = $('#confeti');
  const colores = ['#5b5bd6', '#9b5de5', '#f15bb5', '#fee440', '#00bbf9', '#00f5d4'];
  for (let i = 0; i < 90; i++) {
    const s = document.createElement('span');
    s.style.left = `${Math.random() * 100}%`;
    s.style.background = colores[azar(colores.length)];
    s.style.animationDuration = `${1.8 + Math.random() * 1.8}s`;
    s.style.animationDelay = `${Math.random() * 0.5}s`;
    capa.append(s);
  }
  setTimeout(() => { capa.innerHTML = ''; }, 4500);
}

/* ---------- Banco de contenidos ---------- */
const EMOJIS = ['🍎', '🚀', '🐶', '🌵', '🎸', '⚽', '🦋', '🍕', '🌙', '🐙', '🎈', '🔑', '🍩', '🦊', '🌻', '🐢', '🎲', '🍇', '⛵', '🦉', '🍓', '🐝', '🎩', '🌈'];
const COLORES_PAD = ['#e63946', '#2a9d8f', '#e9c46a', '#457b9d', '#9b5de5', '#f4a261', '#06d6a0', '#ef476f', '#118ab2'];
const PALABRAS = [
  'casa', 'perro', 'árbol', 'mesa', 'nube', 'libro', 'reloj', 'barco', 'luna', 'manzana', 'puente', 'guitarra',
  'zapato', 'espejo', 'montaña', 'lápiz', 'ventana', 'caballo', 'cuchara', 'silla', 'tren', 'flor', 'llave', 'pelota',
  'sombrero', 'faro', 'castillo', 'botella', 'tigre', 'camisa', 'piano', 'avión', 'naranja', 'paraguas', 'escalera', 'vela',
  'tortuga', 'cometa', 'martillo', 'isla', 'campana', 'ballena', 'tambor', 'cohete', 'pájaro', 'mapa', 'cereza', 'abanico',
  'iglesia', 'cámara', 'violín', 'bosque', 'corona', 'dragón', 'globo', 'hormiga', 'jirafa', 'limón', 'mariposa', 'nieve',
  'oso', 'pirata', 'queso', 'río', 'serpiente', 'tijeras', 'uva', 'volcán', 'yate', 'zorro', 'anillo', 'bicicleta',
  'caracol', 'diamante', 'elefante', 'fuego', 'gato', 'helado', 'imán', 'jardín', 'kiwi', 'león', 'mochila', 'nido',
  'ojo', 'pan', 'rana', 'sol', 'taza', 'vaca', 'brújula', 'cactus', 'delfín', 'estrella', 'fresa', 'gorila',
  'hoja', 'tiburón', 'lámpara', 'molino', 'pingüino', 'robot', 'semáforo', 'tenedor', 'puerta', 'almohada', 'búho', 'carta',
];

const TECNICAS = [
  { icono: '🏛️', titulo: 'Palacio de la memoria', texto: 'Imagina un lugar que conoces muy bien (tu casa) y coloca cada elemento que quieres recordar en un punto concreto del recorrido. Para recordar, vuelve a recorrerlo mentalmente.', ejemplo: 'Lista: pan, llaves, reloj → el pan está en la puerta, las llaves cuelgan del sofá, el reloj suena dentro de la nevera.' },
  { icono: '🧩', titulo: 'Agrupación (chunking)', texto: 'La memoria de trabajo retiene unos 4–7 elementos. Agrupa los datos en bloques con sentido para recordar mucho más.', ejemplo: '6 1 8 2 4 7 9 3 → 618 · 247 · 93' },
  { icono: '🎬', titulo: 'Historias absurdas', texto: 'Une los elementos en una historia exagerada, divertida y con movimiento. Lo raro y emocional se recuerda mejor.', ejemplo: 'Un tigre con sombrero toca el piano encima de un barco.' },
  { icono: '🔁', titulo: 'Repetición espaciada', texto: 'Repasa justo antes de olvidar: hoy, mañana, en 3 días, en una semana… Cada repaso alarga la duración del recuerdo.', ejemplo: 'Entrena aquí unos minutos cada día en lugar de una hora un solo día.' },
  { icono: '🗣️', titulo: 'Recuerdo activo', texto: 'Intentar recordar sin mirar es mucho más eficaz que releer. Ponte a prueba constantemente.', ejemplo: 'Cierra el libro y explica en voz alta lo que acabas de leer.' },
  { icono: '🔤', titulo: 'Acrónimos y rimas', texto: 'Usa la primera letra de cada palabra o crea una rima o canción para listas ordenadas.', ejemplo: 'Colores del arcoíris: «RNAVAAV» (rojo, naranja, amarillo, verde, azul, añil, violeta).' },
  { icono: '😴', titulo: 'Sueño y ejercicio', texto: 'El cerebro consolida los recuerdos mientras duermes. El ejercicio aeróbico y una buena hidratación mejoran la memoria.', ejemplo: 'Duerme 7–9 horas y camina 30 minutos al día.' },
  { icono: '🎯', titulo: 'Atención plena', texto: 'No se puede recordar lo que no se ha atendido. Elimina distracciones y concéntrate en una sola cosa.', ejemplo: 'Al conocer a alguien, repite su nombre en voz alta: «Encantado, Laura».' },
];

/* ---------- Ejercicios ---------- */
const JUEGOS = {
  /* ===== Parejas ===== */
  parejas: {
    nombre: 'Parejas',
    icono: '🃏',
    habilidad: 'Memoria visual-espacial',
    descripcion: 'Da la vuelta a las cartas y encuentra todas las parejas con el menor número de movimientos.',
    instrucciones: [
      'Toca una carta para darle la vuelta; después toca otra.',
      'Si coinciden, se quedan descubiertas. Si no, se vuelven a ocultar.',
      'Recuerda la posición de cada símbolo: menos movimientos y menos tiempo dan más puntos.',
    ],
    maxNivel: 8,
    iniciar({ area, hud, nivel, fin }) {
      const pares = [4, 6, 8, 10, 12, 15, 18, 21][nivel - 1];
      const simbolos = barajar(EMOJIS).slice(0, pares);
      const cartas = barajar([...simbolos, ...simbolos]);
      const cols = { 8: 4, 12: 4, 16: 4, 20: 5, 24: 6, 30: 6, 36: 6, 42: 7 }[cartas.length];

      let abiertas = [];
      let movimientos = 0;
      let encontradas = 0;
      let bloqueo = false;
      let inicio = 0;
      let seg = 0;
      let reloj = null;

      const tablero = nodo(`<div class="tablero-cartas" style="--cols:${cols}"></div>`);
      cartas.forEach((s) => {
        const b = nodo(`<button class="carta" aria-label="Carta oculta"><span class="cara frente">?</span><span class="cara dorso">${s}</span></button>`);
        b.dataset.s = s;
        b.addEventListener('click', () => voltear(b));
        tablero.append(b);
      });
      area.append(tablero);

      const actualizar = () => hud([['Movimientos', movimientos], ['Parejas', `${encontradas}/${pares}`], ['Tiempo', fmtTiempo(seg)]]);
      actualizar();

      function voltear(b) {
        if (bloqueo || b.classList.contains('abierta') || b.classList.contains('hecha')) return;
        if (!inicio) {
          inicio = Date.now();
          reloj = Sesion.cada(() => { seg = Math.floor((Date.now() - inicio) / 1000); actualizar(); }, 500);
        }
        Sonido.clic();
        b.classList.add('abierta');
        b.setAttribute('aria-label', `Carta ${b.dataset.s}`);
        abiertas.push(b);
        if (abiertas.length < 2) return;

        movimientos++;
        const [a, c] = abiertas;
        abiertas = [];
        if (a.dataset.s === c.dataset.s) {
          encontradas++;
          Sonido.ok();
          a.classList.replace('abierta', 'hecha');
          c.classList.replace('abierta', 'hecha');
          if (encontradas === pares) {
            Sesion.parar(reloj);
            seg = Math.max(1, Math.round((Date.now() - inicio) / 1000));
            Sesion.luego(terminar, 800);
          }
        } else {
          bloqueo = true;
          Sesion.luego(() => {
            Sonido.suave();
            [a, c].forEach((x) => { x.classList.remove('abierta'); x.setAttribute('aria-label', 'Carta oculta'); });
            bloqueo = false;
          }, 850);
        }
        actualizar();
      }

      function terminar() {
        const sobrantes = Math.max(0, movimientos - pares);
        const bonusTiempo = Math.max(0, pares * 5 - seg) * 4;
        const puntos = Math.max(10, pares * 50 - sobrantes * 10 + bonusTiempo);
        const eficiencia = Math.round((pares / movimientos) * 100);
        let nivelSig = nivel;
        if (movimientos <= pares * 1.75) nivelSig++;
        else if (movimientos > pares * 3) nivelSig--;
        fin({
          puntos,
          nivelSig,
          detalle: [['Movimientos', movimientos], ['Tiempo', fmtTiempo(seg)], ['Eficiencia', `${eficiencia}%`]],
          mensaje: eficiencia >= 60 ? '¡Excelente memoria visual!' : eficiencia >= 40 ? 'Buen trabajo. Intenta fijarte en la posición de cada carta.' : 'Consejo: nombra en voz baja cada símbolo y su posición («perro, esquina»).',
        });
      }
    },
  },

  /* ===== Secuencia de colores y sonidos ===== */
  secuencia: {
    nombre: 'Secuencia',
    icono: '🎵',
    habilidad: 'Memoria secuencial y auditiva',
    descripcion: 'Observa y escucha la secuencia de luces y repítela. Cada ronda añade un paso más.',
    instrucciones: [
      'Los botones se iluminan y suenan en un orden.',
      'Repite la secuencia tocándolos (o con las teclas 1–9).',
      'Cada acierto añade un paso. Un error termina la partida.',
    ],
    maxNivel: 9,
    iniciar({ area, hud, nivel, fin }) {
      const nPads = nivel <= 3 ? 4 : nivel <= 6 ? 6 : 9;
      const vel = Math.max(220, 620 - nivel * 45);
      const estado = nodo('<p class="estado">Preparado…</p>');
      const tablero = nodo(`<div class="pads pads-${nPads} bloqueado"></div>`);
      const pads = COLORES_PAD.slice(0, nPads).map((c, i) => {
        const b = nodo(`<button class="pad" style="--c:${c}" aria-label="Botón ${i + 1}"><span>${i + 1}</span></button>`);
        b.addEventListener('pointerdown', (e) => { e.preventDefault(); pulsar(i); });
        tablero.append(b);
        return b;
      });
      area.append(estado, tablero);

      const seq = [];
      let pos = 0;
      let turno = false;
      const mejor = Datos.d.mejores.secuenciaLargo || 0;
      const actualizar = () => hud([['Longitud', seq.length], ['Botones', nPads], ['Tu récord', Math.max(mejor, seq.length - 1)]]);

      Sesion.tecla = (e) => {
        const n = parseInt(e.key, 10);
        if (n >= 1 && n <= nPads) pulsar(n - 1);
      };

      function encender(i, ms) {
        pads[i].classList.add('activo');
        Sonido.nota(i, ms / 1000);
        Sesion.luego(() => pads[i].classList.remove('activo'), ms);
      }

      async function ronda() {
        turno = false;
        seq.push(azar(nPads));
        pos = 0;
        actualizar();
        estado.className = 'estado';
        estado.textContent = 'Observa la secuencia…';
        tablero.classList.add('bloqueado');
        await Sesion.esperar(700);
        for (const i of seq) {
          encender(i, vel);
          await Sesion.esperar(vel + 140);
        }
        estado.textContent = 'Tu turno: repite la secuencia';
        tablero.classList.remove('bloqueado');
        turno = true;
      }

      function pulsar(i) {
        if (!turno) return;
        if (i !== seq[pos]) {
          turno = false;
          Sonido.error();
          pads[i].classList.add('activo');
          Sesion.luego(() => pads[i].classList.remove('activo'), 300);
          pads[seq[pos]].classList.add('correcto');
          estado.className = 'estado mal';
          estado.textContent = `¡Ups! Era el botón ${seq[pos] + 1}`;
          Sesion.luego(terminar, 1500);
          return;
        }
        encender(i, 220);
        pos++;
        if (pos === seq.length) {
          turno = false;
          estado.className = 'estado ok';
          estado.textContent = '¡Correcto!';
          Sesion.luego(ronda, 900);
        }
      }

      function terminar() {
        const logrado = seq.length - 1;
        if (logrado > mejor) { Datos.d.mejores.secuenciaLargo = logrado; Datos.guardar(); }
        let nivelSig = nivel;
        if (logrado >= 8) nivelSig++;
        else if (logrado < 4) nivelSig--;
        fin({
          puntos: Math.round(logrado * 10 * (1 + (nivel - 1) * 0.25)),
          nivelSig,
          detalle: [['Secuencia lograda', logrado], ['Botones', nPads], ['Velocidad', `${vel} ms`]],
          mensaje: logrado >= 8 ? '¡Impresionante capacidad secuencial!' : 'Consejo: agrupa la secuencia en bloques de 3 o tararea la melodía.',
        });
      }

      ronda();
    },
  },

  /* ===== Dígitos ===== */
  digitos: {
    nombre: 'Dígitos',
    icono: '🔢',
    habilidad: 'Memoria de trabajo verbal',
    descripcion: 'Memoriza números cada vez más largos y escríbelos. Incluye modo inverso y lectura en voz alta.',
    instrucciones: [
      'Aparecerán dígitos de uno en uno.',
      'Cuando terminen, escríbelos en el mismo orden (o al revés en modo inverso).',
      'Cada acierto alarga el número. Tienes 3 vidas.',
    ],
    opciones: [
      { id: 'inverso', texto: 'Modo inverso: escribe los dígitos al revés (más difícil, ×1,5 puntos)' },
      { id: 'voz', texto: 'Escuchar los dígitos en voz alta' },
    ],
    maxNivel: 12,
    iniciar({ area, hud, nivel, opc, fin }) {
      let largo = 2 + nivel;
      let vidas = 3;
      let maxOk = 0;
      let puntos = 0;
      let aciertos = 0;
      let numero = '';
      let esperandoRespuesta = false;
      const pausa = Math.max(450, 950 - nivel * 35);

      const pantalla = nodo('<div class="digito-display" aria-live="polite"></div>');
      const estado = nodo('<p class="estado"></p>');
      const form = nodo('<form class="entrada-digitos oculto"><input inputmode="numeric" autocomplete="off" aria-label="Tu respuesta" maxlength="24"><button class="btn primario">OK</button></form>');
      const input = $('input', form);
      const teclado = nodo('<div class="teclado oculto"></div>');
      ['1', '2', '3', '4', '5', '6', '7', '8', '9', '⌫', '0', '✔'].forEach((k) => {
        const b = nodo(`<button type="button">${k}</button>`);
        b.addEventListener('click', () => {
          if (k === '⌫') input.value = input.value.slice(0, -1);
          else if (k === '✔') form.requestSubmit();
          else input.value += k;
          Sonido.clic();
        });
        teclado.append(b);
      });
      area.append(estado, pantalla, form, teclado);

      const actualizar = () => hud([['Dígitos', largo], ['Vidas', '❤️'.repeat(vidas) || '—'], ['Puntos', puntos]]);

      async function ronda() {
        esperandoRespuesta = false;
        numero = Array.from({ length: largo }, () => azar(10)).join('');
        form.classList.add('oculto');
        teclado.classList.add('oculto');
        pantalla.classList.remove('chico');
        estado.className = 'estado';
        estado.textContent = 'Memoriza…';
        pantalla.textContent = '';
        actualizar();
        await Sesion.esperar(700);
        for (const d of numero) {
          pantalla.textContent = d;
          pantalla.classList.remove('pulso');
          void pantalla.offsetWidth;
          pantalla.classList.add('pulso');
          if (opc.voz) hablar(d);
          Sonido.tono(400, 0.05, 'sine', 0.05);
          await Sesion.esperar(pausa);
          pantalla.textContent = '';
          await Sesion.esperar(160);
        }
        pantalla.textContent = '?';
        estado.textContent = opc.inverso ? 'Escríbelos AL REVÉS' : 'Escribe los dígitos en orden';
        input.value = '';
        form.classList.remove('oculto');
        teclado.classList.remove('oculto');
        input.focus();
        esperandoRespuesta = true;
      }

      form.addEventListener('submit', (e) => {
        e.preventDefault();
        if (!esperandoRespuesta) return;
        const resp = input.value.replace(/\D/g, '');
        if (!resp) return;
        esperandoRespuesta = false;
        const esperado = opc.inverso ? [...numero].reverse().join('') : numero;
        form.classList.add('oculto');
        teclado.classList.add('oculto');
        if (resp === esperado) {
          Sonido.ok();
          aciertos++;
          maxOk = Math.max(maxOk, largo);
          puntos += Math.round(largo * 10 * (opc.inverso ? 1.5 : 1));
          estado.className = 'estado ok';
          estado.textContent = '¡Correcto! Un dígito más…';
          pantalla.textContent = '✔';
          largo++;
          actualizar();
          Sesion.luego(ronda, 1100);
        } else {
          Sonido.error();
          vidas--;
          actualizar();
          estado.className = 'estado mal';
          estado.textContent = `Escribiste ${resp}`;
          pantalla.classList.add('chico');
          pantalla.innerHTML = `<div><small>La respuesta era</small>${esperado}</div>`;
          Sesion.luego(vidas > 0 ? ronda : terminar, 2200);
        }
      });

      function terminar() {
        fin({
          puntos,
          nivelSig: maxOk ? Math.max(1, maxOk - 3) : nivel - 1,
          detalle: [['Máximo recordado', maxOk ? `${maxOk} dígitos` : '—'], ['Aciertos', aciertos], ['Modo', opc.inverso ? 'Inverso' : 'Directo']],
          mensaje: maxOk >= 8 ? '¡Por encima de la media! (la mayoría recuerda 7 ± 2)' : 'Consejo: agrupa los dígitos de 3 en 3, como un número de teléfono.',
        });
      }

      ronda();
    },
  },

  /* ===== Cuadrícula ===== */
  cuadricula: {
    nombre: 'Cuadrícula',
    icono: '🟪',
    habilidad: 'Memoria espacial',
    descripcion: 'Unas casillas se iluminan un instante. Recuerda cuáles eran y tócalas. La cuadrícula crece.',
    instrucciones: [
      'Observa qué casillas se iluminan.',
      'Cuando se apaguen, toca todas las que estaban encendidas (el orden no importa).',
      'Un error te quita una vida (tienes 3). Cada acierto añade una casilla más.',
    ],
    maxNivel: 15,
    iniciar({ area, hud, nivel, fin }) {
      let cantidad = 2 + nivel;
      let vidas = 3;
      let puntos = 0;
      let maxOk = 0;
      let rondas = 0;
      const tam = (c) => (c <= 4 ? 3 : c <= 7 ? 4 : c <= 10 ? 5 : c <= 14 ? 6 : 7);
      const estado = nodo('<p class="estado"></p>');
      const zona = nodo('<div style="width:100%;display:flex;justify-content:center"></div>');
      area.append(estado, zona);

      const actualizar = () => hud([['Casillas', cantidad], ['Vidas', '❤️'.repeat(vidas) || '—'], ['Puntos', puntos]]);

      async function ronda() {
        rondas++;
        actualizar();
        const n = tam(cantidad);
        const objetivo = new Set(barajar([...Array(n * n).keys()]).slice(0, cantidad));
        const grid = nodo(`<div class="cuadricula bloqueado" style="--n:${n}"></div>`);
        const celdas = [...Array(n * n)].map((_, i) => {
          const c = nodo(`<button class="celda" aria-label="Casilla ${i + 1}"></button>`);
          c.addEventListener('click', () => tocar(i, c));
          grid.append(c);
          return c;
        });
        zona.replaceChildren(grid);
        estado.className = 'estado';
        estado.textContent = 'Memoriza las casillas…';

        let aciertos = 0;
        let activo = false;

        await Sesion.esperar(600);
        objetivo.forEach((i) => celdas[i].classList.add('iluminada'));
        Sonido.tono(500, 0.1, 'sine', 0.06);
        await Sesion.esperar(900 + cantidad * 200);
        objetivo.forEach((i) => celdas[i].classList.remove('iluminada'));
        estado.textContent = `Toca las ${cantidad} casillas`;
        grid.classList.remove('bloqueado');
        activo = true;

        function tocar(i, c) {
          if (!activo || c.classList.contains('acierto')) return;
          if (objetivo.has(i)) {
            c.classList.add('acierto');
            aciertos++;
            Sonido.nota(aciertos, 0.12);
            if (aciertos === cantidad) {
              activo = false;
              grid.classList.add('bloqueado');
              Sonido.ok();
              puntos += cantidad * 10;
              maxOk = Math.max(maxOk, cantidad);
              cantidad = Math.min(cantidad + 1, 25);
              estado.className = 'estado ok';
              estado.textContent = '¡Perfecto!';
              Sesion.luego(ronda, 1000);
            }
          } else {
            activo = false;
            grid.classList.add('bloqueado');
            c.classList.add('fallo');
            Sonido.error();
            vidas--;
            actualizar();
            objetivo.forEach((j) => { if (!celdas[j].classList.contains('acierto')) celdas[j].classList.add('faltante'); });
            estado.className = 'estado mal';
            estado.textContent = vidas > 0 ? 'Fallaste — mira dónde estaban' : 'Sin vidas';
            Sesion.luego(vidas > 0 ? ronda : terminar, 1800);
          }
        }
      }

      function terminar() {
        fin({
          puntos,
          nivelSig: maxOk ? Math.max(1, maxOk - 3) : nivel - 1,
          detalle: [['Máximo recordado', maxOk ? `${maxOk} casillas` : '—'], ['Rondas', rondas], ['Puntos', puntos]],
          mensaje: maxOk >= 9 ? '¡Gran memoria espacial!' : 'Consejo: busca formas o figuras que formen las casillas iluminadas.',
        });
      }

      ronda();
    },
  },

  /* ===== Palabras ===== */
  palabras: {
    nombre: 'Lista de palabras',
    icono: '📝',
    habilidad: 'Memoria verbal a corto plazo',
    descripcion: 'Estudia una lista de palabras durante un tiempo limitado y luego escribe todas las que recuerdes.',
    instrucciones: [
      'Tendrás unos segundos para memorizar una lista de palabras.',
      'Después escribe todas las que recuerdes, en cualquier orden (pulsa Enter tras cada una).',
      'Prueba las técnicas de la historia absurda o el palacio de la memoria.',
    ],
    maxNivel: 10,
    iniciar({ area, hud, nivel, fin }) {
      const n = Math.min(4 + nivel * 2, 24);
      const palabras = barajar(PALABRAS).slice(0, n);
      const mapa = new Map(palabras.map((p) => [normalizar(p), p]));
      const recordadas = new Set();
      const tEstudio = Math.round(n * 2.5);
      const tRecuerdo = n * 6 + 20;
      let restante = tEstudio;
      let reloj = null;

      // Fase 1: estudio
      const estado = nodo('<p class="estado">Memoriza estas palabras</p>');
      const barra = nodo('<div class="barra-tiempo"><div style="width:100%"></div></div>');
      const lista = nodo('<div class="lista-palabras"></div>');
      palabras.forEach((p, i) => lista.append(nodo(`<div class="palabra" style="animation-delay:${i * 60}ms">${escapar(p)}</div>`)));
      const listo = nodo('<button class="btn primario">¡Ya las tengo!</button>');
      area.append(estado, barra, lista, listo);

      const tic = (total, alAcabar, etiqueta) => {
        hud([[etiqueta, fmtTiempo(restante)], ['Palabras', etiqueta === 'Estudio' ? n : `${recordadas.size}/${n}`]]);
        barra.firstElementChild.style.width = `${(restante / total) * 100}%`;
        if (restante <= 0) { Sesion.parar(reloj); alAcabar(); return; }
        restante--;
      };
      tic(tEstudio, recordar, 'Estudio');
      reloj = Sesion.cada(() => tic(tEstudio, recordar, 'Estudio'), 1000);
      listo.addEventListener('click', recordar);

      // Fase 2: recuerdo
      let enRecuerdo = false;
      function recordar() {
        if (enRecuerdo) return;
        enRecuerdo = true;
        Sesion.parar(reloj);
        lista.remove();
        listo.remove();
        estado.textContent = 'Escribe las palabras que recuerdes';
        const form = nodo('<form class="entrada-texto"><input autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="Escribe una palabra y pulsa Enter" aria-label="Palabra"><button class="btn primario">Añadir</button></form>');
        const input = $('input', form);
        const chips = nodo('<div class="chips"></div>');
        const terminarBtn = nodo('<button class="btn">Terminar</button>');
        area.append(form, chips, terminarBtn);
        input.focus();

        form.addEventListener('submit', (e) => {
          e.preventDefault();
          const intentos = input.value.split(/[\s,;]+/).map(normalizar).filter(Boolean);
          input.value = '';
          let algunaNueva = false;
          let algunaMal = false;
          intentos.forEach((t) => {
            if (mapa.has(t) && !recordadas.has(t)) {
              recordadas.add(t);
              chips.append(nodo(`<span class="chip">${escapar(mapa.get(t))}</span>`));
              algunaNueva = true;
            } else if (!mapa.has(t)) {
              algunaMal = true;
            }
          });
          if (algunaNueva) Sonido.ok();
          if (algunaMal) {
            Sonido.suave();
            form.classList.remove('temblor');
            void form.offsetWidth;
            form.classList.add('temblor');
          }
          hud([['Tiempo', fmtTiempo(Math.max(0, restante + 1))], ['Palabras', `${recordadas.size}/${n}`]]);
          if (recordadas.size === n) Sesion.luego(terminar, 500);
        });
        terminarBtn.addEventListener('click', terminar);

        restante = tRecuerdo;
        tic(tRecuerdo, terminar, 'Tiempo');
        reloj = Sesion.cada(() => tic(tRecuerdo, terminar, 'Tiempo'), 1000);
      }

      let terminado = false;
      function terminar() {
        if (terminado) return;
        terminado = true;
        Sesion.parar(reloj);
        const ok = recordadas.size;
        const olvidadas = palabras.filter((p) => !recordadas.has(normalizar(p)));
        const pct = ok / n;
        let nivelSig = nivel;
        if (pct >= 0.8) nivelSig++;
        else if (pct < 0.4) nivelSig--;
        fin({
          puntos: ok * 20 + (ok === n ? 50 : 0),
          nivelSig,
          detalle: [['Recordadas', `${ok}/${n}`], ['Acierto', `${Math.round(pct * 100)}%`], ['Tiempo de estudio', `${tEstudio} s`]],
          mensaje: pct >= 0.8 ? '¡Memoria verbal excelente!' : 'Consejo: inventa una historia que conecte las palabras en orden.',
          extra: olvidadas.length
            ? `<p class="suave">Palabras olvidadas:</p><div class="chips">${olvidadas.map((p) => `<span class="chip falta">${escapar(p)}</span>`).join('')}</div>`
            : '<p>🎉 ¡Recordaste todas las palabras!</p>',
        });
      }
    },
  },

  /* ===== N-back ===== */
  nback: {
    nombre: 'N-Back',
    icono: '🎯',
    habilidad: 'Memoria de trabajo y concentración',
    descripcion: 'Un cuadro se ilumina en distintas posiciones. Pulsa cuando la posición coincida con la de N pasos atrás.',
    instrucciones: [
      'En cada paso se ilumina una casilla de la cuadrícula 3×3.',
      'Pulsa «¡Coincide!» (o la barra espaciadora) si la posición es la misma que hace N pasos.',
      'Si no coincide, no pulses nada. Es uno de los ejercicios más estudiados para la memoria de trabajo.',
    ],
    maxNivel: 10,
    iniciar({ area, hud, nivel, fin }) {
      const n = Math.min(5, Math.ceil(nivel / 2));
      const intervalo = nivel % 2 ? 2800 : 2300;
      const total = 20 + n;
      const seq = [];
      for (let i = 0; i < total; i++) {
        if (i >= n && Math.random() < 0.3) seq.push(seq[i - n]);
        else {
          let p;
          do { p = azar(9); } while (i >= n && p === seq[i - n]);
          seq.push(p);
        }
      }

      const estado = nodo(`<p class="estado">Modo <b>${n}-back</b>: ¿coincide con la posición de hace ${n} ${n === 1 ? 'paso' : 'pasos'}?</p>`);
      const grid = nodo('<div class="nback"></div>');
      const celdas = [...Array(9)].map(() => { const c = nodo('<div class="celda"></div>'); grid.append(c); return c; });
      const boton = nodo('<button class="btn primario grande btn-coincide" disabled>¡Coincide! (Espacio)</button>');
      area.append(estado, grid, boton);

      let i = -1;
      let respondido = false;
      let aciertos = 0;
      let omisiones = 0;
      let falsas = 0;
      let rechazos = 0;

      const actualizar = () => hud([['Paso', `${Math.max(0, i + 1)}/${total}`], ['Aciertos', aciertos], ['Errores', omisiones + falsas]]);
      const destello = (cls) => {
        boton.classList.remove('ok', 'mal');
        void boton.offsetWidth;
        boton.classList.add(cls);
        Sesion.luego(() => boton.classList.remove(cls), 400);
      };

      function responder() {
        if (i < n || respondido || i >= total) return;
        respondido = true;
        if (seq[i] === seq[i - n]) { aciertos++; Sonido.ok(); destello('ok'); }
        else { falsas++; Sonido.error(); destello('mal'); }
        actualizar();
      }
      boton.addEventListener('click', responder);
      Sesion.tecla = (e) => { if (e.code === 'Space') { e.preventDefault(); responder(); } };

      function paso() {
        if (i >= n && !respondido) {
          if (seq[i] === seq[i - n]) { omisiones++; destello('mal'); }
          else rechazos++;
        }
        i++;
        if (i >= total) { terminar(); return; }
        respondido = false;
        boton.disabled = i < n;
        celdas[seq[i]].classList.add('iluminada');
        Sonido.tono(420 + seq[i] * 40, 0.12, 'sine', 0.07);
        Sesion.luego(() => celdas[seq[i]].classList.remove('iluminada'), Math.min(1000, intervalo - 400));
        actualizar();
        Sesion.luego(paso, intervalo);
      }

      async function cuenta() {
        actualizar();
        const zona = nodo('<div class="cuenta-atras"></div>');
        grid.before(zona);
        for (const t of ['3', '2', '1']) {
          zona.textContent = t;
          zona.style.animation = 'none';
          void zona.offsetWidth;
          zona.style.animation = '';
          Sonido.clic();
          await Sesion.esperar(700);
        }
        zona.remove();
        paso();
      }

      function terminar() {
        boton.disabled = true;
        const evaluados = total - n;
        const precision = (aciertos + rechazos) / evaluados;
        let nivelSig = nivel;
        if (precision >= 0.85) nivelSig++;
        else if (precision < 0.6) nivelSig--;
        fin({
          puntos: Math.round(precision * 100 * n),
          nivelSig,
          detalle: [['Precisión', `${Math.round(precision * 100)}%`], ['Coincidencias acertadas', aciertos], ['Omisiones', omisiones], ['Falsas alarmas', falsas]],
          mensaje: precision >= 0.85 ? `¡Dominas el ${n}-back! Sube de nivel.` : 'Consejo: repite mentalmente las últimas posiciones («arriba, centro…»).',
        });
      }

      cuenta();
    },
  },
};

/* ---------- Navegación y vistas ---------- */
const app = $('#app');
let diario = null; // { cola: [ids], idx, total }

function ir(vista, ...args) {
  Sesion.limpiar();
  if (vista !== 'jugar' && vista !== 'resultado' && vista !== 'intro') diario = null;
  $$('.barra nav [data-ir]').forEach((b) => b.classList.toggle('activo', b.dataset.ir === vista));
  window.scrollTo(0, 0);
  Vistas[vista](...args);
  app.focus({ preventScroll: true });
}

function pintarHud(contenedor) {
  return (items) => {
    contenedor.innerHTML = items.map(([k, v]) => `<div class="dato"><small>${escapar(k)}</small><b>${escapar(v)}</b></div>`).join('');
  };
}

const Vistas = {
  inicio() {
    const racha = Datos.rachaActual();
    const hoyN = Datos.d.historial.filter((h) => h.dia === fechaISO()).length;
    const saludo = (() => {
      const h = new Date().getHours();
      return h < 12 ? 'Buenos días' : h < 20 ? 'Buenas tardes' : 'Buenas noches';
    })();
    app.innerHTML = `
      <section class="bienvenida">
        <div>
          <h1>${saludo} 👋</h1>
          <p>Unos minutos al día de entrenamiento constante marcan la diferencia.</p>
          <div class="racha">
            <div>🔥 Racha: ${racha} ${racha === 1 ? 'día' : 'días'}</div>
            <div>✅ Hoy: ${hoyN} ${hoyN === 1 ? 'ejercicio' : 'ejercicios'}</div>
            <div>🏆 Total: ${Datos.d.historial.length}</div>
          </div>
        </div>
        <button class="btn grande" id="btnDiario">▶ Entrenamiento diario</button>
      </section>
      <h2>Ejercicios</h2>
      <p class="suave">Elige un ejercicio. El nivel se ajusta automáticamente a tu rendimiento.</p>
      <div class="rejilla-juegos">
        ${Object.entries(JUEGOS).map(([id, j]) => `
          <button class="juego-card" data-juego="${id}">
            <span class="icono-juego">${j.icono}</span>
            <h3>${j.nombre}</h3>
            <p>${j.descripcion}</p>
            <div class="etiquetas">
              <span class="etiqueta destacada">${j.habilidad}</span>
              <span class="etiqueta">Nivel ${Datos.nivel(id)}</span>
              ${Datos.d.mejores[id] ? `<span class="etiqueta">🏆 ${Datos.d.mejores[id]}</span>` : ''}
            </div>
          </button>`).join('')}
      </div>`;
    $$('[data-juego]', app).forEach((b) => b.addEventListener('click', () => ir('intro', b.dataset.juego)));
    $('#btnDiario').addEventListener('click', () => {
      const cola = barajar(Object.keys(JUEGOS)).slice(0, 3);
      diario = { cola, idx: 0, puntos: 0 };
      ir('intro', cola[0]);
    });
  },

  intro(id) {
    const j = JUEGOS[id];
    let nivel = limitar(Datos.nivel(id), 1, j.maxNivel);
    const etiquetaDiario = diario ? `<p class="etiqueta destacada" style="display:inline-block">Entrenamiento diario · ejercicio ${diario.idx + 1} de ${diario.cola.length}</p>` : '';
    app.innerHTML = `
      <section class="intro tarjeta">
        ${etiquetaDiario}
        <div class="icono-juego">${j.icono}</div>
        <h1>${j.nombre}</h1>
        <p class="suave">${j.habilidad}</p>
        <ul>${j.instrucciones.map((t) => `<li>${t}</li>`).join('')}</ul>
        <div>
          <p class="suave" style="margin:0">Nivel</p>
          <div class="selector-nivel">
            <button class="btn" id="menos" aria-label="Bajar nivel">−</button>
            <span class="valor" id="valorNivel">${nivel}</span>
            <button class="btn" id="mas" aria-label="Subir nivel">+</button>
          </div>
        </div>
        ${j.opciones ? `<div class="opciones">${j.opciones.map((o) => `<label><input type="checkbox" data-op="${o.id}"> ${o.texto}</label>`).join('')}</div>` : ''}
        <div class="fila centro">
          <button class="btn" id="volver">← Volver</button>
          <button class="btn primario grande" id="empezar">Comenzar</button>
        </div>
        <p class="suave" style="font-size:.85rem;margin-top:1rem">Pulsa Enter para empezar</p>
      </section>`;
    const pintar = () => { $('#valorNivel').textContent = nivel; };
    $('#menos').addEventListener('click', () => { nivel = Math.max(1, nivel - 1); pintar(); });
    $('#mas').addEventListener('click', () => { nivel = Math.min(j.maxNivel, nivel + 1); pintar(); });
    $('#volver').addEventListener('click', () => ir('inicio'));
    const empezar = () => {
      const opc = {};
      $$('[data-op]', app).forEach((c) => { opc[c.dataset.op] = c.checked; });
      ir('jugar', id, nivel, opc);
    };
    $('#empezar').addEventListener('click', empezar);
    $('#empezar').focus();
    Sesion.tecla = (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') empezar(); };
  },

  jugar(id, nivel, opc) {
    const j = JUEGOS[id];
    app.innerHTML = `
      <div class="cabecera-juego">
        <button class="btn" id="salir">✕ Salir</button>
        <h2>${j.icono} ${j.nombre} · Nivel ${nivel}</h2>
      </div>
      <div class="hud"></div>
      <div class="area"></div>`;
    $('#salir').addEventListener('click', () => ir('intro', id));
    const inicio = Date.now();
    j.iniciar({
      area: $('.area', app),
      hud: pintarHud($('.hud', app)),
      nivel,
      opc: opc || {},
      fin: (res) => {
        const nivelSig = limitar(res.nivelSig, 1, j.maxNivel);
        const seg = Math.round((Date.now() - inicio) / 1000);
        const r = Datos.registrar(id, res.puntos, nivel, nivelSig, seg);
        ir('resultado', id, { ...res, nivel, nivelSig, ...r });
      },
    });
  },

  resultado(id, res) {
    const j = JUEGOS[id];
    let cambio = '';
    if (res.nivelSig > res.nivel) cambio = `⬆️ ¡Subes al nivel ${res.nivelSig}!`;
    else if (res.nivelSig < res.nivel) cambio = `Próxima partida: nivel ${res.nivelSig}, para consolidar.`;
    else cambio = `Sigues en el nivel ${res.nivelSig}.`;

    let botones;
    if (diario) {
      diario.puntos += res.puntos;
      diario.idx++;
      botones = diario.idx < diario.cola.length
        ? `<button class="btn primario grande" id="siguiente">Siguiente ejercicio →</button>`
        : `<button class="btn primario grande" id="menu">Terminar entrenamiento</button>`;
    } else {
      botones = `<button class="btn" id="menu">Menú</button><button class="btn primario grande" id="repetir">↻ Jugar otra vez</button>`;
    }
    const finDiario = diario && diario.idx >= diario.cola.length
      ? `<div class="tarjeta" style="margin:1rem 0"><h3>🎉 ¡Entrenamiento diario completado!</h3><p class="suave" style="margin:0">Puntos totales de hoy: <b>${diario.puntos}</b></p></div>`
      : '';

    app.innerHTML = `
      <section class="resultado tarjeta">
        <div class="icono-juego">${j.icono}</div>
        <h1>${j.nombre}</h1>
        ${res.record ? '<div class="insignia">🏆 ¡Nuevo récord personal!</div>' : ''}
        <div class="puntos-grandes">${res.puntos}</div>
        <p class="suave" style="margin-top:0">puntos · mejor marca: ${Datos.d.mejores[id] || res.puntos}</p>
        <p><b>${res.mensaje || ''}</b></p>
        <div class="detalles">${res.detalle.map(([k, v]) => `<div><small>${escapar(k)}</small><b>${escapar(v)}</b></div>`).join('')}</div>
        ${res.extra ? `<div class="extra">${res.extra}</div>` : ''}
        <p>${cambio}</p>
        ${finDiario}
        <div class="fila centro">${botones}</div>
      </section>`;

    if (res.record || finDiario) { confeti(); Sonido.fanfarria(); }
    const repetir = $('#repetir');
    const siguiente = $('#siguiente');
    if (repetir) { repetir.addEventListener('click', () => ir('intro', id)); repetir.focus(); }
    if (siguiente) { siguiente.addEventListener('click', () => ir('intro', diario.cola[diario.idx])); siguiente.focus(); }
    $('#menu')?.addEventListener('click', () => ir('inicio'));
  },

  progreso(juegoSel) {
    const h = Datos.d.historial;
    const minutos = Math.round(h.reduce((s, x) => s + (x.segundos || 0), 0) / 60);
    const ids = Object.keys(JUEGOS);
    juegoSel = juegoSel || (h.length ? h[h.length - 1].juego : ids[0]);

    app.innerHTML = `
      <h1>Tu progreso</h1>
      <div class="resumen">
        <div class="tarjeta"><span class="suave">Sesiones</span><b>${h.length}</b></div>
        <div class="tarjeta"><span class="suave">Minutos entrenados</span><b>${minutos}</b></div>
        <div class="tarjeta"><span class="suave">Racha actual</span><b>🔥 ${Datos.rachaActual()}</b></div>
        <div class="tarjeta"><span class="suave">Racha máxima</span><b>${Datos.d.racha.maxima || 0}</b></div>
      </div>

      <section class="tarjeta seccion">
        <h3>Actividad de los últimos 14 días</h3>
        <div id="actividad"></div>
      </section>

      <section class="tarjeta seccion">
        <h3>Evolución de puntuación</h3>
        <div class="selector-juego">${ids.map((id) => `<button class="btn ${id === juegoSel ? 'activo' : ''}" data-sel="${id}">${JUEGOS[id].icono} ${JUEGOS[id].nombre}</button>`).join('')}</div>
        <div id="evolucion"></div>
      </section>

      <section class="tarjeta seccion">
        <h3>Por ejercicio</h3>
        <div class="tabla-scroll">
          <table class="tabla">
            <thead><tr><th>Ejercicio</th><th class="num">Partidas</th><th class="num">Nivel</th><th class="num">Récord</th><th class="num">Media (últ. 5)</th></tr></thead>
            <tbody>${ids.map((id) => {
              const partidas = h.filter((x) => x.juego === id);
              const ult = partidas.slice(-5);
              const media = ult.length ? Math.round(ult.reduce((s, x) => s + x.puntos, 0) / ult.length) : '—';
              return `<tr><td>${JUEGOS[id].icono} ${JUEGOS[id].nombre}</td><td class="num">${partidas.length}</td><td class="num">${Datos.nivel(id)}</td><td class="num">${Datos.d.mejores[id] || '—'}</td><td class="num">${media}</td></tr>`;
            }).join('')}</tbody>
          </table>
        </div>
      </section>

      <div class="fila">
        <button class="btn" id="exportar">⬇ Exportar datos</button>
        <button class="btn" id="borrar">🗑 Borrar progreso</button>
      </div>`;

    $('#actividad').innerHTML = graficoActividad(h);
    $('#evolucion').innerHTML = graficoEvolucion(h.filter((x) => x.juego === juegoSel).slice(-25));
    $$('[data-sel]', app).forEach((b) => b.addEventListener('click', () => Vistas.progreso(b.dataset.sel)));
    $('#borrar').addEventListener('click', () => {
      if (confirm('¿Seguro que quieres borrar todo tu progreso? No se puede deshacer.')) {
        const { sonido, tema } = Datos.d;
        Datos.d = Object.assign(Datos.base(), { sonido, tema });
        Datos.guardar();
        ir('progreso');
      }
    });
    $('#exportar').addEventListener('click', () => {
      const blob = new Blob([JSON.stringify(Datos.d, null, 2)], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `memoria-progreso-${fechaISO()}.json`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    });
  },

  tecnicas() {
    app.innerHTML = `
      <h1>Técnicas para mejorar la memoria</h1>
      <p class="suave">Combina los ejercicios con estas estrategias. Pruébalas en «Lista de palabras» y «Dígitos».</p>
      <div class="tecnicas">
        ${TECNICAS.map((t) => `
          <article class="tarjeta">
            <h3><span>${t.icono}</span>${t.titulo}</h3>
            <p>${t.texto}</p>
            <div class="ejemplo">💡 ${t.ejemplo}</div>
          </article>`).join('')}
      </div>`;
  },
};

/* ---------- Gráficos SVG ---------- */
function graficoActividad(h) {
  const dias = [];
  for (let i = 13; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const iso = fechaISO(d);
    dias.push({ iso, etiqueta: `${d.getDate()}/${d.getMonth() + 1}`, n: h.filter((x) => x.dia === iso).length });
  }
  const max = Math.max(1, ...dias.map((d) => d.n));
  const W = 560, H = 150, base = 125, ancho = W / 14;
  const barras = dias.map((d, i) => {
    const alto = (d.n / max) * 100;
    const x = i * ancho + ancho * 0.2;
    return `<rect class="barra-act" x="${x}" y="${base - alto}" width="${ancho * 0.6}" height="${alto}" rx="3"><title>${d.etiqueta}: ${d.n}</title></rect>
      ${d.n ? `<text x="${x + ancho * 0.3}" y="${base - alto - 4}" text-anchor="middle">${d.n}</text>` : ''}
      <text x="${x + ancho * 0.3}" y="${H - 8}" text-anchor="middle">${d.etiqueta}</text>`;
  }).join('');
  return `<svg class="grafico" viewBox="0 0 ${W} ${H}" role="img" aria-label="Ejercicios por día"><line class="eje" x1="0" x2="${W}" y1="${base}" y2="${base}"/>${barras}</svg>`;
}

function graficoEvolucion(partidas) {
  if (partidas.length < 2) {
    return `<p class="suave">${partidas.length ? 'Juega una partida más' : 'Todavía no hay partidas'} para ver tu evolución.</p>`;
  }
  const W = 560, H = 200, m = { i: 40, d: 12, s: 14, b: 26 };
  const max = Math.max(...partidas.map((p) => p.puntos), 1);
  const x = (i) => m.i + (i / (partidas.length - 1)) * (W - m.i - m.d);
  const y = (v) => H - m.b - (v / max) * (H - m.s - m.b);
  const puntos = partidas.map((p, i) => `${x(i).toFixed(1)},${y(p.puntos).toFixed(1)}`).join(' ');
  const guias = [0, 0.5, 1].map((f) => {
    const v = Math.round(max * f);
    return `<line class="eje" x1="${m.i}" x2="${W - m.d}" y1="${y(v)}" y2="${y(v)}"/><text x="${m.i - 6}" y="${y(v) + 4}" text-anchor="end">${v}</text>`;
  }).join('');
  const marcas = partidas.map((p, i) => {
    const f = new Date(p.fecha);
    return `<circle class="punto" cx="${x(i)}" cy="${y(p.puntos)}" r="4"><title>${f.toLocaleDateString('es')} · ${p.puntos} pts · nivel ${p.nivel}</title></circle>`;
  }).join('');
  return `<svg class="grafico" viewBox="0 0 ${W} ${H}" role="img" aria-label="Evolución de la puntuación">${guias}<polyline class="linea" points="${puntos}"/>${marcas}
    <text x="${m.i}" y="${H - 6}">Antiguas</text><text x="${W - m.d}" y="${H - 6}" text-anchor="end">Recientes</text></svg>`;
}

/* ---------- Arranque ---------- */
function aplicarTema() {
  if (Datos.d.tema) document.documentElement.dataset.theme = Datos.d.tema;
  else delete document.documentElement.dataset.theme;
}

function iniciarApp() {
  Datos.cargar();
  aplicarTema();
  const btnSonido = $('#btnSonido');
  btnSonido.textContent = Datos.d.sonido ? '🔊' : '🔇';
  btnSonido.addEventListener('click', () => {
    Datos.d.sonido = !Datos.d.sonido;
    btnSonido.textContent = Datos.d.sonido ? '🔊' : '🔇';
    Datos.guardar();
    Sonido.clic();
  });
  $('#btnTema').addEventListener('click', () => {
    const oscuroActual = Datos.d.tema ? Datos.d.tema === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
    Datos.d.tema = oscuroActual ? 'light' : 'dark';
    Datos.guardar();
    aplicarTema();
  });
  $$('[data-ir]').forEach((b) => b.addEventListener('click', () => ir(b.dataset.ir)));
  ir('inicio');
}

iniciarApp();
