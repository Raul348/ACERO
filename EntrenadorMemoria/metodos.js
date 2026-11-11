'use strict';

/* =====================================================================
   Métodos con evidencia sólida
     - Palacio de la memoria (método de loci) — Dresler et al., 2017
     - Repaso espaciado con recuerdo activo  — Cepeda et al., 2006;
       Karpicke y Roediger, 2008. Programación con el algoritmo SM-2.
   ===================================================================== */

/* ---------- Palacio de la memoria ---------- */

const LOCI_EJEMPLO = [
  'la puerta de entrada', 'el felpudo', 'el perchero', 'el sofá', 'la televisión',
  'la mesa del comedor', 'el fregadero', 'el frigorífico', 'la cama', 'la ducha',
];

const PISTAS_LOCI = [
  'Imagínala enorme, ocupando todo el lugar.',
  'Hazla moverse: que salte, gire o explote.',
  'Añade un sonido o un olor exagerado.',
  'Que interactúe con el lugar: que lo rompa, lo moje o lo pinte.',
  'Imagínala de un color absurdo o multiplicada por cien.',
  'Ponte tú en la escena: tócala, empújala, cómela.',
];

const lociActuales = () => (Datos.d.loci && Datos.d.loci.length >= 5 ? Datos.d.loci : LOCI_EJEMPLO);

JUEGOS.loci = {
  nombre: 'Palacio de la memoria',
  icono: '🏛️',
  categoria: 'evidencia',
  habilidad: 'Método de loci',
  evidencia: 'Alta: Dresler et al. (2017)',
  descripcion: 'La técnica de los campeones de memoria: coloca cada palabra en un lugar de un recorrido que conoces bien y recupérala en orden.',
  instrucciones: [
    'Usa un recorrido real que conozcas muy bien (tu casa, tu trabajo). Puedes escribir el tuyo abajo: funciona mucho mejor que el de ejemplo.',
    'Para cada palabra, crea una <b>imagen exagerada</b> de esa cosa en el lugar indicado. Tómate el tiempo que necesites.',
    'Después recorre mentalmente el camino y escribe la palabra de cada lugar.',
    'Entrénalo a diario: en el estudio, la mejora apareció tras varias semanas de práctica constante.',
  ],
  maxNivel: 26,
  maxNivelActual: () => Math.max(1, lociActuales().length - 4),
  etiquetaNivel: (n) => `${Math.min(n + 4, lociActuales().length)} palabras`,

  introExtra(cont, alCambiar) {
    const propio = Datos.d.loci && Datos.d.loci.length >= 5;
    cont.innerHTML = `
      <details class="editor-loci" ${propio ? '' : 'open'}>
        <summary>🗺️ Mi recorrido (${lociActuales().length} lugares${propio ? '' : ' · ejemplo'})</summary>
        <p class="suave">Escribe un lugar por línea, en el orden en que los recorrerías. Mínimo 5; cuantos más lugares, más palabras podrás memorizar.</p>
        <textarea rows="8" aria-label="Lugares del recorrido">${escapar(lociActuales().join('\n'))}</textarea>
        <div class="fila"><button class="btn" type="button" id="guardarLoci">Guardar recorrido</button><span class="suave" id="msgLoci"></span></div>
      </details>`;
    $('#guardarLoci', cont).addEventListener('click', () => {
      const lugares = $('textarea', cont).value.split('\n').map((l) => l.trim()).filter(Boolean);
      const msg = $('#msgLoci', cont);
      if (lugares.length < 5) { msg.textContent = 'Necesitas al menos 5 lugares.'; return; }
      Datos.d.loci = lugares.slice(0, 30);
      Datos.guardar();
      msg.textContent = `✔ Guardado (${Datos.d.loci.length} lugares)`;
      $('summary', cont).textContent = `🗺️ Mi recorrido (${Datos.d.loci.length} lugares)`;
      alCambiar();
    });
  },

  iniciar({ area, hud, nivel, fin }) {
    const loci = lociActuales();
    const n = Math.min(nivel + 4, loci.length);
    const palabras = barajar(PALABRAS).slice(0, n);
    const tiempos = [];
    let k = 0;
    let tPaso = Date.now();

    // Fase 1: codificación (a tu ritmo)
    const tarjeta = nodo(`
      <div class="tarjeta loci-paso">
        <p class="suave lugar-num"></p>
        <p class="lugar"></p>
        <div class="palabra-loci"></div>
        <p class="suave pista"></p>
        <button class="btn primario grande">Siguiente →</button>
        <p class="suave" style="font-size:.85rem;margin:.5rem 0 0">Pulsa cuando tengas la imagen clara (o Enter)</p>
      </div>`);
    area.append(tarjeta);
    const boton = $('button', tarjeta);

    const mostrar = () => {
      hud([['Fase', 'Memorizar'], ['Palabra', `${k + 1}/${n}`]]);
      $('.lugar-num', tarjeta).textContent = `Lugar ${k + 1} de ${n}`;
      $('.lugar', tarjeta).textContent = `📍 ${loci[k]}`;
      const p = $('.palabra-loci', tarjeta);
      p.textContent = palabras[k];
      p.classList.remove('pulso');
      void p.offsetWidth;
      p.classList.add('pulso');
      $('.pista', tarjeta).textContent = `💡 ${PISTAS_LOCI[k % PISTAS_LOCI.length]}`;
      tPaso = Date.now();
      boton.focus();
    };
    const avanzar = () => {
      tiempos.push((Date.now() - tPaso) / 1000);
      Sonido.clic();
      k++;
      if (k < n) mostrar();
      else transicion();
    };
    boton.addEventListener('click', avanzar);
    mostrar();

    // Breve pausa antes de recordar
    async function transicion() {
      area.innerHTML = '<p class="estado">Ahora recorre mentalmente tu palacio, lugar por lugar…</p><div class="cuenta-atras"></div>';
      const c = $('.cuenta-atras', area);
      hud([['Fase', 'Recordar'], ['Palabras', n]]);
      for (const t of ['3', '2', '1']) {
        c.textContent = t;
        c.style.animation = 'none';
        void c.offsetWidth;
        c.style.animation = '';
        await Sesion.esperar(900);
      }
      recordar();
    }

    // Fase 2: recuerdo en orden
    function recordar() {
      area.innerHTML = `
        <p class="estado">¿Qué palabra había en cada lugar?</p>
        <form class="recuerdo-loci">
          ${loci.slice(0, n).map((l, i) => `<label><span>${i + 1}. ${escapar(l)}</span><input autocomplete="off" autocapitalize="off" spellcheck="false" data-i="${i}"></label>`).join('')}
          <button class="btn primario grande">Comprobar</button>
        </form>`;
      const form = $('form', area);
      const inputs = $$('input', form);
      inputs[0].focus();
      inputs.forEach((inp, i) => inp.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && i < inputs.length - 1) { e.preventDefault(); inputs[i + 1].focus(); }
      }));
      form.addEventListener('submit', (e) => {
        e.preventDefault();
        corregir(inputs.map((x) => x.value));
      });
    }

    function corregir(respuestas) {
      const ok = respuestas.map((r, i) => {
        const t = normalizar(r);
        const p = normalizar(palabras[i]);
        return t === p || t === `${p}s` || t === `${p}es`;
      });
      const aciertos = ok.filter(Boolean).length;
      const pct = aciertos / n;
      const media = tiempos.reduce((a, b) => a + b, 0) / n;
      let nivelSig = nivel;
      if (pct >= 0.9) nivelSig += 2;
      else if (pct >= 0.7) nivelSig += 1;
      else if (pct < 0.5) nivelSig -= 1;
      fin({
        puntos: aciertos * 10 + (aciertos === n ? n * 5 : 0),
        nivelSig,
        detalle: [['Recordadas en orden', `${aciertos}/${n}`], ['Acierto', `${Math.round(pct * 100)}%`], ['Tiempo medio por imagen', `${media.toFixed(1)} s`]],
        mensaje: pct >= 0.9 ? '¡Excelente! La próxima vez, más palabras.' : pct >= 0.6 ? 'Buen trabajo. Haz las imágenes aún más absurdas y con movimiento.' : 'Tómate más tiempo en cada imagen y usa tu propio recorrido real.',
        extra: `<div class="tabla-scroll"><table class="tabla"><thead><tr><th>Lugar</th><th>Palabra</th><th>Tu respuesta</th></tr></thead><tbody>
          ${palabras.map((p, i) => `<tr><td>${escapar(loci[i])}</td><td>${escapar(p)}</td><td class="${ok[i] ? 'sube' : 'baja'}">${ok[i] ? '✔' : '✘'} ${escapar(respuestas[i] || '—')}</td></tr>`).join('')}
        </tbody></table></div>`,
      });
    }
  },
};

/* ---------- Repaso espaciado (SM-2) ---------- */

const MAZO_EJEMPLO = [
  ['Australia', 'Canberra'], ['Canadá', 'Ottawa'], ['Turquía', 'Ankara'], ['Brasil', 'Brasilia'], ['Marruecos', 'Rabat'],
  ['Suiza', 'Berna'], ['Nueva Zelanda', 'Wellington'], ['Kenia', 'Nairobi'], ['Vietnam', 'Hanói'], ['Noruega', 'Oslo'],
  ['Egipto', 'El Cairo'], ['Perú', 'Lima'], ['Corea del Sur', 'Seúl'], ['Polonia', 'Varsovia'], ['Hungría', 'Budapest'],
  ['Filipinas', 'Manila'], ['Nigeria', 'Abuya'], ['Pakistán', 'Islamabad'], ['Finlandia', 'Helsinki'], ['Colombia', 'Bogotá'],
].map(([pais, capital]) => [`Capital de ${pais}`, capital]);

const sumarDias = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return fechaISO(d); };

function nuevaTarjeta(frente, reverso) {
  return {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 7),
    frente, reverso, ef: 2.5, reps: 0, intervalo: 0, proxima: fechaISO(), creada: fechaISO(), revisiones: 0, aciertos: 0,
  };
}

function tarjetasPendientes() {
  const hoy = fechaISO();
  return Datos.d.tarjetas.filter((t) => t.proxima <= hoy);
}

/**
 * Algoritmo SM-2 (Wozniak, 1990). q: 1 = otra vez, 3 = difícil, 4 = bien, 5 = fácil.
 * Devuelve una copia programada sin modificar la original.
 */
function programarSM2(t, q) {
  const c = { ...t };
  if (q < 3) {
    c.reps = 0;
    c.intervalo = 1;
  } else {
    c.reps += 1;
    if (c.reps === 1) c.intervalo = q === 5 ? 3 : 1;
    else if (c.reps === 2) c.intervalo = 6;
    else c.intervalo = Math.round(c.intervalo * c.ef);
  }
  c.ef = Math.max(1.3, c.ef + (0.1 - (5 - q) * (0.08 + (5 - q) * 0.02)));
  c.proxima = sumarDias(c.intervalo);
  return c;
}

const textoIntervalo = (d) => (d <= 1 ? '1 día' : d < 30 ? `${d} días` : d < 365 ? `${Math.round(d / 30)} mes${Math.round(d / 30) > 1 ? 'es' : ''}` : `${(d / 365).toFixed(1)} años`);

function estadisticasRepaso() {
  const t = Datos.d.tarjetas;
  const repasos = t.reduce((s, x) => s + x.revisiones, 0);
  const aciertos = t.reduce((s, x) => s + x.aciertos, 0);
  return {
    repasos,
    retencion: repasos ? `${Math.round((aciertos / repasos) * 100)}%` : '—',
    dominadas: t.filter((x) => x.intervalo >= 21).length,
  };
}

Vistas.repaso = function () {
  const t = Datos.d.tarjetas;
  const pend = tarjetasPendientes().length;
  const est = estadisticasRepaso();
  app.innerHTML = `
    <div class="cabecera-juego"><button class="btn" id="volver">← Inicio</button><h2>🔁 Repaso espaciado</h2></div>
    <section class="tarjeta seccion">
      <p>Crea tarjetas con lo que <b>tú</b> quieras recordar: nombres de personas, vocabulario, fechas, datos del trabajo o de estudio…
      En cada repaso <b>intenta recordar antes de ver la respuesta</b> (recuerdo activo). El programa te volverá a preguntar cada tarjeta justo cuando estés a punto de olvidarla, con intervalos cada vez más largos.</p>
      <div class="detalles">
        <div><small>Pendientes hoy</small><b>${pend}</b></div>
        <div><small>Tarjetas</small><b>${t.length}</b></div>
        <div><small>Retención</small><b>${est.retencion}</b></div>
        <div><small>Dominadas (≥ 21 días)</small><b>${est.dominadas}</b></div>
      </div>
      <div class="fila centro">
        <button class="btn primario grande" id="repasar" ${pend ? '' : 'disabled'}>${pend ? `Repasar ahora (${pend})` : '✔ Todo al día por hoy'}</button>
      </div>
    </section>

    <section class="tarjeta seccion">
      <h3>Añadir tarjeta</h3>
      <form class="form-tarjeta" id="formTarjeta">
        <input id="frente" placeholder="Pregunta (p. ej.: Nombre de la vecina del 3.º)" autocomplete="off" required>
        <input id="reverso" placeholder="Respuesta (p. ej.: Laura)" autocomplete="off" required>
        <button class="btn primario">Añadir</button>
      </form>
      <details style="margin-top:.75rem">
        <summary>Añadir varias a la vez</summary>
        <p class="suave">Una tarjeta por línea: <code>pregunta ; respuesta</code></p>
        <textarea id="lote" rows="5" placeholder="Hello ; Hola&#10;Cumpleaños de mamá ; 12 de marzo"></textarea>
        <button class="btn" id="anadirLote">Añadir todas</button>
      </details>
      ${t.length ? '' : '<p style="margin-top:1rem"><button class="btn" id="ejemplo">Probar con un mazo de ejemplo (capitales del mundo)</button></p>'}
      <p class="suave" id="msgTarjeta"></p>
    </section>

    ${t.length ? `
    <section class="tarjeta seccion">
      <h3>Mis tarjetas (${t.length})</h3>
      <div class="tabla-scroll">
        <table class="tabla">
          <thead><tr><th>Pregunta</th><th>Respuesta</th><th class="num">Próximo repaso</th><th></th></tr></thead>
          <tbody>${[...t].sort((a, b) => a.proxima.localeCompare(b.proxima)).map((x) => `
            <tr><td>${escapar(x.frente)}</td><td>${escapar(x.reverso)}</td>
            <td class="num">${x.proxima <= fechaISO() ? '<b>hoy</b>' : new Date(`${x.proxima}T00:00`).toLocaleDateString('es')}</td>
            <td class="num"><button class="btn mini" data-borrar="${x.id}" aria-label="Borrar tarjeta">🗑</button></td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </section>` : ''}`;

  $('#volver').addEventListener('click', () => ir('inicio'));
  $('#repasar').addEventListener('click', () => ir('repasoSesion'));
  const msg = (txt) => { $('#msgTarjeta').textContent = txt; };
  $('#formTarjeta').addEventListener('submit', (e) => {
    e.preventDefault();
    const f = $('#frente').value.trim();
    const r = $('#reverso').value.trim();
    if (!f || !r) return;
    Datos.d.tarjetas.push(nuevaTarjeta(f, r));
    Datos.guardar();
    Vistas.repaso();
    $('#msgTarjeta').textContent = '✔ Tarjeta añadida';
    $('#frente').focus();
  });
  $('#anadirLote').addEventListener('click', () => {
    const lineas = $('#lote').value.split('\n').map((l) => l.split(/;|\t/)).filter((p) => p.length >= 2 && p[0].trim() && p[1].trim());
    if (!lineas.length) { msg('Formato: pregunta ; respuesta (una por línea)'); return; }
    lineas.forEach(([f, r]) => Datos.d.tarjetas.push(nuevaTarjeta(f.trim(), r.trim())));
    Datos.guardar();
    Vistas.repaso();
    $('#msgTarjeta').textContent = `✔ ${lineas.length} tarjetas añadidas`;
  });
  $('#ejemplo')?.addEventListener('click', () => {
    MAZO_EJEMPLO.forEach(([f, r]) => Datos.d.tarjetas.push(nuevaTarjeta(f, r)));
    Datos.guardar();
    Vistas.repaso();
  });
  $$('[data-borrar]', app).forEach((b) => b.addEventListener('click', () => {
    if (!confirm('¿Borrar esta tarjeta?')) return;
    Datos.d.tarjetas = Datos.d.tarjetas.filter((x) => x.id !== b.dataset.borrar);
    Datos.guardar();
    Vistas.repaso();
  }));
};

Vistas.repasoSesion = function () {
  const cola = barajar(tarjetasPendientes().map((t) => t.id));
  const vistasHoy = new Set();       // tarjetas ya calificadas en esta sesión
  const repeticiones = {};           // veces que una tarjeta fallada se ha vuelto a mostrar
  let revisadas = 0;
  let recordadas = 0;
  const inicio = Date.now();

  app.innerHTML = `
    <div class="cabecera-juego"><button class="btn" id="salir">✕ Salir</button><h2>🔁 Repaso espaciado</h2></div>
    <div class="hud"></div>
    <div class="area"></div>`;
  $('#salir').addEventListener('click', () => { if (revisadas) registrar(); ir('repaso'); });
  const hud = pintarHud($('.hud', app));
  const area = $('.area', app);

  function registrar() {
    Datos.d.historial.push({ juego: 'repaso', puntos: revisadas, nivel: 0, fecha: new Date().toISOString(), dia: fechaISO(), segundos: Math.round((Date.now() - inicio) / 1000), recordadas });
    Datos.marcarDia();
    Datos.guardar();
  }

  function siguiente() {
    if (!cola.length) { terminar(); return; }
    const id = cola.shift();
    const t = Datos.d.tarjetas.find((x) => x.id === id);
    if (!t) { siguiente(); return; }
    hud([['Pendientes', cola.length + 1], ['Revisadas', revisadas], ['Recordadas', recordadas]]);
    area.innerHTML = `
      <div class="tarjeta flashcard">
        <p class="suave">Pregunta</p>
        <div class="frente">${escapar(t.frente)}</div>
        <form class="entrada-texto"><input autocomplete="off" placeholder="Escribe tu respuesta (opcional)" aria-label="Tu respuesta"><button class="btn primario">Mostrar respuesta</button></form>
        <p class="suave" style="font-size:.85rem">Intenta recordarla de verdad antes de mirar: ese esfuerzo es lo que fija el recuerdo.</p>
      </div>`;
    const input = $('input', area);
    input.focus();
    $('form', area).addEventListener('submit', (e) => {
      e.preventDefault();
      mostrarRespuesta(t, input.value.trim());
    });
  }

  function mostrarRespuesta(t, escrita) {
    const coincide = escrita && normalizar(escrita) === normalizar(t.reverso);
    const opciones = [
      { q: 1, texto: 'Otra vez', tecla: '1', clase: 'mal' },
      { q: 3, texto: 'Difícil', tecla: '2', clase: '' },
      { q: 4, texto: 'Bien', tecla: '3', clase: '' },
      { q: 5, texto: 'Fácil', tecla: '4', clase: 'ok' },
    ];
    const primera = !vistasHoy.has(t.id);
    area.innerHTML = `
      <div class="tarjeta flashcard">
        <p class="suave">Pregunta</p>
        <div class="frente">${escapar(t.frente)}</div>
        <hr>
        <p class="suave">Respuesta</p>
        <div class="reverso">${escapar(t.reverso)}</div>
        ${escrita ? `<p class="${coincide ? 'sube' : 'suave'}">Tu respuesta: ${escapar(escrita)} ${coincide ? '✔' : ''}</p>` : ''}
        <p class="suave">¿Cómo de bien la recordaste?</p>
        <div class="botones-calif">
          ${opciones.map((o) => `<button class="btn ${o.clase}" data-q="${o.q}"><b>${o.texto}</b><small>${primera ? (o.q === 1 ? 'repetir hoy' : textoIntervalo(programarSM2(t, o.q).intervalo)) : (o.q === 1 ? 'repetir hoy' : 'hecho')}</small><kbd>${o.tecla}</kbd></button>`).join('')}
        </div>
      </div>`;
    const calificar = (q) => {
      Sesion.tecla = null;
      if (primera) {
        // Solo la primera calificación de la sesión cambia la programación (como en SM-2).
        vistasHoy.add(t.id);
        Object.assign(t, programarSM2(t, q));
        t.revisiones++;
        if (q >= 3) { t.aciertos++; recordadas++; }
        revisadas++;
        Datos.guardar();
      }
      if (q < 3 && (repeticiones[t.id] || 0) < 3) {
        repeticiones[t.id] = (repeticiones[t.id] || 0) + 1;
        cola.splice(Math.min(cola.length, 3), 0, t.id); // vuelve a salir en unas tarjetas
        Sonido.suave();
      } else if (q >= 3) {
        Sonido.ok();
      }
      siguiente();
    };
    $$('[data-q]', area).forEach((b) => b.addEventListener('click', () => calificar(Number(b.dataset.q))));
    Sesion.tecla = (e) => {
      const m = { 1: 1, 2: 3, 3: 4, 4: 5 }[e.key];
      if (m) calificar(m);
    };
    $('[data-q="4"]', area).focus();
  }

  function terminar() {
    if (revisadas) registrar();
    const proximas = Datos.d.tarjetas.map((x) => x.proxima).filter((f) => f > fechaISO()).sort();
    app.innerHTML = `
      <section class="resultado tarjeta">
        <div class="icono-juego">🔁</div>
        <h1>¡Repaso completado!</h1>
        <div class="detalles">
          <div><small>Tarjetas revisadas</small><b>${revisadas}</b></div>
          <div><small>Recordadas a la primera</small><b>${revisadas ? Math.round((recordadas / revisadas) * 100) : 0}%</b></div>
          <div><small>Próximo repaso</small><b>${proximas.length ? new Date(`${proximas[0]}T00:00`).toLocaleDateString('es') : '—'}</b></div>
        </div>
        <p>Vuelve mañana: repasar en los días indicados es lo que hace que el recuerdo dure meses.</p>
        <div class="fila centro">
          <button class="btn" id="tarjetas">Mis tarjetas</button>
          <button class="btn primario grande" id="inicio">Inicio</button>
        </div>
      </section>`;
    if (revisadas) Sonido.fanfarria();
    $('#tarjetas').addEventListener('click', () => ir('repaso'));
    $('#inicio').addEventListener('click', () => ir('inicio'));
  }

  siguiente();
};
