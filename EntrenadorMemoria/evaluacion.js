'use strict';

/* =====================================================================
   Evaluación de memoria (pre / post)
   Protocolo fijo para comparar tu memoria a lo largo del tiempo:
     1. Aprendizaje de 15 palabras (3 ensayos, recuerdo libre)
     2. Span de dígitos directo        (procedimiento clásico de span)
     3. Span de dígitos inverso
     4. Span espacial (tipo bloques de Corsi)
     5. Recuerdo diferido de las 15 palabras
   No es un diagnóstico clínico: sirve para seguir tu propia evolución.
   ===================================================================== */

/* Seis listas paralelas: cada una tiene una palabra de las mismas 15 categorías
   (mueble, recipiente, persona, ropa, verdura, edificio, parte de la casa, bebida,
   animal de granja, cuerpo, aseo, vajilla, paisaje, utensilio de cocina, vehículo),
   para que su dificultad sea equivalente. No se usan en los ejercicios. */
const LISTAS_EVAL = [
  ['cortina', 'maleta', 'abuelo', 'pantalón', 'cebolla', 'escuela', 'tejado', 'café', 'pavo', 'nariz', 'jabón', 'cuchillo', 'playa', 'sartén', 'coche'],
  ['alfombra', 'cesta', 'sobrino', 'falda', 'tomate', 'hospital', 'chimenea', 'leche', 'conejo', 'rodilla', 'cepillo', 'plato', 'desierto', 'olla', 'camión'],
  ['armario', 'bolso', 'vecino', 'chaqueta', 'zanahoria', 'museo', 'balcón', 'té', 'pato', 'codo', 'peine', 'vaso', 'selva', 'horno', 'autobús'],
  ['estante', 'cartera', 'primo', 'bufanda', 'patata', 'biblioteca', 'pasillo', 'zumo', 'oveja', 'hombro', 'toalla', 'jarra', 'lago', 'nevera', 'moto'],
  ['cajón', 'carpeta', 'tío', 'guante', 'pepino', 'farmacia', 'sótano', 'sopa', 'cerdo', 'tobillo', 'esponja', 'bandeja', 'valle', 'tostadora', 'tractor'],
  ['sofá', 'baúl', 'nieto', 'calcetín', 'lechuga', 'estadio', 'terraza', 'vino', 'gallina', 'muñeca', 'champú', 'cuenco', 'costa', 'batidora', 'metro'],
];

const PASOS_EVAL = ['Aprendizaje de palabras', 'Dígitos directos', 'Dígitos inversos', 'Secuencia espacial', 'Recuerdo diferido'];

/** Medidas que se guardan y comparan. */
const MEDIDAS = [
  { id: 'aprendizaje', nombre: 'Aprendizaje de palabras', max: 45, desc: 'Suma de los 3 ensayos (máx. 45)' },
  { id: 'diferido', nombre: 'Recuerdo diferido', max: 15, desc: 'Palabras recordadas tras varios minutos (máx. 15)' },
  { id: 'directo', nombre: 'Dígitos directos', max: 9, desc: 'Span: dígitos máximos repetidos en orden', ref: '≈ 6–7' },
  { id: 'inverso', nombre: 'Dígitos inversos', max: 8, desc: 'Span: dígitos máximos repetidos al revés', ref: '≈ 4–5' },
  { id: 'espacial', nombre: 'Secuencia espacial', max: 9, desc: 'Span espacial tipo Corsi', ref: '≈ 5–6' },
];

/* Posiciones de los 9 bloques del tablero espacial (en %, irregulares a propósito). */
const BLOQUES_CORSI = [[8, 62], [18, 14], [36, 40], [44, 76], [56, 8], [64, 48], [78, 22], [80, 72], [22, 84]];

/* ---------- Utilidades de la prueba ---------- */

/** Devuelve la clave de la lista que coincide con lo escrito (tolera tildes y plurales). */
function coincidir(texto, mapa) {
  if (mapa.has(texto)) return texto;
  if (texto.endsWith('es') && mapa.has(texto.slice(0, -2))) return texto.slice(0, -2);
  if (texto.endsWith('s') && mapa.has(texto.slice(0, -1))) return texto.slice(0, -1);
  return null;
}

/** Pantalla de instrucciones; se resuelve al pulsar el botón. */
function avisoPaso(area, titulo, html, textoBoton = 'Empezar') {
  return new Promise((resolve) => {
    area.innerHTML = `<section class="tarjeta intro"><h2>${titulo}</h2>${html}<button class="btn primario grande">${textoBoton}</button></section>`;
    const b = $('button', area);
    b.focus();
    b.addEventListener('click', resolve, { once: true });
  });
}

async function presentarPalabras(area, lista, texto) {
  area.innerHTML = `<p class="estado">${texto}</p><div class="digito-display chico"></div>`;
  const d = $('.digito-display', area);
  await Sesion.esperar(1200);
  for (const p of lista) {
    d.textContent = p;
    d.classList.remove('pulso');
    void d.offsetWidth;
    d.classList.add('pulso');
    await Sesion.esperar(1500);
    d.textContent = '';
    await Sesion.esperar(300);
  }
}

/** Recuerdo libre sin indicar qué es correcto (como en una prueba real). */
function recuerdoLibre(area, lista, segundos, titulo) {
  return new Promise((resolve) => {
    const mapa = new Map(lista.map((p) => [normalizar(p), p]));
    const correctas = new Set();
    const escritas = new Set();
    let intrusiones = 0;
    let restante = segundos;
    let hecho = false;

    area.innerHTML = `
      <p class="estado">${titulo}</p>
      <p class="suave centro" style="margin:-.5rem 0 0">Escribe cada palabra y pulsa Enter. El orden no importa.</p>
      <div class="barra-tiempo"><div style="width:100%"></div></div>
      <form class="entrada-texto"><input autocomplete="off" autocapitalize="off" spellcheck="false" aria-label="Palabra" placeholder="Escribe una palabra"><button class="btn primario">Añadir</button></form>
      <div class="chips"></div>
      <button class="btn" type="button" id="fin">He terminado</button>`;
    const input = $('input', area);
    const chips = $('.chips', area);
    const barra = $('.barra-tiempo div', area);
    input.focus();

    $('form', area).addEventListener('submit', (e) => {
      e.preventDefault();
      input.value.split(/[\s,;]+/).map(normalizar).filter(Boolean).forEach((t) => {
        if (escritas.has(t)) return;
        escritas.add(t);
        chips.append(nodo(`<span class="chip neutro">${escapar(t)}</span>`));
        const k = coincidir(t, mapa);
        if (k) correctas.add(k);
        else intrusiones++;
      });
      input.value = '';
      Sonido.clic();
    });

    const acabar = () => {
      if (hecho) return;
      hecho = true;
      Sesion.parar(reloj);
      resolve({ aciertos: correctas.size, intrusiones });
    };
    $('#fin', area).addEventListener('click', acabar);
    const reloj = Sesion.cada(() => {
      restante--;
      barra.style.width = `${(restante / segundos) * 100}%`;
      if (restante <= 0) acabar();
    }, 1000);
  });
}

async function mostrarDigitos(area, seq, texto) {
  area.innerHTML = `<p class="estado">${texto}</p><div class="digito-display"></div>`;
  const d = $('.digito-display', area);
  await Sesion.esperar(900);
  for (const c of seq) {
    d.textContent = c;
    d.classList.remove('pulso');
    void d.offsetWidth;
    d.classList.add('pulso');
    await Sesion.esperar(800); // 1 dígito por segundo
    d.textContent = '';
    await Sesion.esperar(200);
  }
}

function pedirDigitos(area, inverso) {
  return new Promise((resolve) => {
    area.innerHTML = `
      <p class="estado">${inverso ? 'Escríbelos AL REVÉS' : 'Escríbelos en el mismo orden'}</p>
      <form class="entrada-digitos"><input inputmode="numeric" autocomplete="off" aria-label="Tu respuesta" maxlength="12"><button class="btn primario">OK</button></form>
      <div class="teclado"></div>`;
    const input = $('input', area);
    const form = $('form', area);
    const teclado = $('.teclado', area);
    ['1', '2', '3', '4', '5', '6', '7', '8', '9', '⌫', '0', '✔'].forEach((k) => {
      const b = nodo(`<button type="button">${k}</button>`);
      b.addEventListener('click', () => {
        if (k === '⌫') input.value = input.value.slice(0, -1);
        else if (k === '✔') form.requestSubmit();
        else input.value += k;
      });
      teclado.append(b);
    });
    input.focus();
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const r = input.value.replace(/\D/g, '');
      if (r) resolve(r);
    });
  });
}

/** Span de dígitos: 2 intentos por longitud; se para cuando fallan ambos. */
async function spanDigitos(area, inverso) {
  let largo = inverso ? 2 : 3;
  const max = inverso ? 8 : 9;
  let span = 0;
  while (largo <= max) {
    let ok = 0;
    for (let intento = 1; intento <= 2; intento++) {
      const seq = barajar([...'0123456789']).slice(0, largo).join('');
      await mostrarDigitos(area, seq, `${largo} dígitos · intento ${intento} de 2`);
      const resp = await pedirDigitos(area, inverso);
      if (resp === (inverso ? [...seq].reverse().join('') : seq)) ok++;
    }
    if (ok === 0) break;
    span = largo;
    largo++;
  }
  return span;
}

/** Span espacial (tipo Corsi): mismo procedimiento con bloques que se iluminan. */
async function spanEspacial(area) {
  let largo = 2;
  let span = 0;
  while (largo <= 9) {
    let ok = 0;
    for (let intento = 1; intento <= 2; intento++) {
      const seq = barajar([...Array(9).keys()]).slice(0, largo);
      area.innerHTML = `<p class="estado">${largo} bloques · intento ${intento} de 2 · Observa</p><div class="corsi bloqueado"></div>`;
      const tablero = $('.corsi', area);
      const bloques = BLOQUES_CORSI.map(([x, y], i) => {
        const b = nodo(`<button class="bloque" style="left:${x}%;top:${y}%" aria-label="Bloque ${i + 1}"></button>`);
        tablero.append(b);
        return b;
      });
      await Sesion.esperar(900);
      for (const i of seq) {
        bloques[i].classList.add('iluminado');
        await Sesion.esperar(700);
        bloques[i].classList.remove('iluminado');
        await Sesion.esperar(300);
      }
      $('.estado', area).textContent = `Toca los ${largo} bloques en el mismo orden`;
      tablero.classList.remove('bloqueado');
      const toques = await new Promise((resolve) => {
        const r = [];
        bloques.forEach((b, i) => b.addEventListener('click', () => {
          if (r.length >= largo) return;
          r.push(i);
          b.classList.add('tocado');
          Sonido.clic();
          Sesion.luego(() => b.classList.remove('tocado'), 250);
          if (r.length === largo) Sesion.luego(() => resolve(r), 300);
        }));
      });
      if (toques.every((v, k) => v === seq[k])) ok++;
    }
    if (ok === 0) break;
    span = largo;
    largo++;
  }
  return span;
}

/* ---------- Vistas de evaluación ---------- */

Vistas.evaluacion = function () {
  const evals = Datos.d.evaluaciones;
  const ultima = evals[evals.length - 1];
  const dias = ultima ? diasEntre(ultima.fecha.slice(0, 10), fechaISO()) : null;
  const pronto = ultima && dias < 7;

  app.innerHTML = `
    <h1>📋 Evaluación de memoria</h1>
    <section class="tarjeta seccion">
      <p>Una prueba fija de unos <b>12 minutos</b> con 5 tareas basadas en pruebas clásicas de la psicología:
      aprendizaje de una lista de palabras, span de dígitos directo e inverso, span espacial tipo Corsi y recuerdo diferido.</p>
      <p class="suave">Hazla antes de empezar a entrenar y repítela cada 2 semanas. Cada evaluación usa una lista de palabras distinta, de dificultad equivalente. Es orientativa: <b>no es un diagnóstico médico</b>.</p>
      ${pronto ? `<p class="alerta">⚠️ Tu última evaluación fue hace ${dias} ${dias === 1 ? 'día' : 'días'}. Repetirla tan pronto infla los resultados por efecto de práctica; lo recomendable es esperar al menos 1–2 semanas.</p>` : ''}
      <p class="suave">Busca un momento tranquilo y sin distracciones, y hazla siempre en condiciones parecidas (misma hora, mismo dispositivo).</p>
      <button class="btn primario grande" id="empezarEval">${evals.length ? 'Nueva evaluación' : 'Empezar evaluación inicial'}</button>
    </section>
    ${evals.length ? `
      <section class="tarjeta seccion">
        <h3>Tus resultados</h3>
        <div class="tabla-scroll">
          <table class="tabla">
            <thead><tr><th>Medida</th>${evals.map((e, i) => `<th class="num">${i === 0 ? 'Inicial' : `#${i + 1}`}<br><small>${new Date(e.fecha).toLocaleDateString('es')}</small></th>`).join('')}${evals.length > 1 ? '<th class="num">Cambio</th>' : ''}</tr></thead>
            <tbody>${MEDIDAS.map((m) => {
              const d = evals[evals.length - 1][m.id] - evals[0][m.id];
              return `<tr><td>${m.nombre}<br><small class="suave">${m.desc}</small></td>${evals.map((e) => `<td class="num">${e[m.id]}</td>`).join('')}${evals.length > 1 ? `<td class="num ${d > 0 ? 'sube' : d < 0 ? 'baja' : ''}">${d > 0 ? '+' : ''}${d}</td>` : ''}</tr>`;
            }).join('')}</tbody>
          </table>
        </div>
        <p class="suave" style="font-size:.88rem">«Cambio» compara la última evaluación con la inicial. En los span, variaciones de ±1 son normales de un día a otro; lo que cuenta es la tendencia.</p>
      </section>
      <section class="tarjeta seccion">
        <h3>Evolución</h3>
        <div class="selector-juego">${MEDIDAS.map((m, i) => `<button class="btn ${i === 0 ? 'activo' : ''}" data-medida="${m.id}">${m.nombre}</button>`).join('')}</div>
        <div id="graficoEval"></div>
      </section>` : ''}`;

  $('#empezarEval').addEventListener('click', () => ir('evalPrueba'));
  const pintar = (id) => {
    $$('[data-medida]', app).forEach((b) => b.classList.toggle('activo', b.dataset.medida === id));
    $('#graficoEval').innerHTML = graficoLinea(evals.map((e) => ({ v: e[id], fecha: e.fecha, info: `${e[id]}` })));
  };
  if (evals.length) {
    $$('[data-medida]', app).forEach((b) => b.addEventListener('click', () => pintar(b.dataset.medida)));
    pintar(MEDIDAS[0].id);
  }
};

Vistas.evalPrueba = async function () {
  const n = Datos.d.evaluaciones.length;
  const indiceLista = n % LISTAS_EVAL.length;
  const lista = barajar(LISTAS_EVAL[indiceLista]);
  const inicio = Date.now();

  app.innerHTML = `
    <div class="cabecera-juego">
      <button class="btn" id="salir">✕ Salir</button>
      <h2>📋 Evaluación de memoria</h2>
    </div>
    <div class="barra-tiempo progreso-eval"><div style="width:0"></div></div>
    <p class="suave centro" id="pasoEval"></p>
    <div class="area"></div>`;
  $('#salir').addEventListener('click', () => {
    if (confirm('¿Salir de la evaluación? No se guardará.')) ir('evaluacion');
  });
  const area = $('.area', app);
  const paso = (i) => {
    $('#pasoEval').textContent = `Prueba ${i + 1} de ${PASOS_EVAL.length} · ${PASOS_EVAL[i]}`;
    $('.progreso-eval div').style.width = `${(i / PASOS_EVAL.length) * 100}%`;
  };

  // 1. Aprendizaje de palabras
  paso(0);
  await avisoPaso(area, '📝 Aprendizaje de palabras', `
    <p>Verás <b>15 palabras</b>, una tras otra. Cuando terminen, escribe <b>todas las que recuerdes</b>, en cualquier orden.</p>
    <p>La misma lista se mostrará <b>3 veces</b>. Intenta recordar más palabras cada vez.</p>
    <p class="suave">Al final de la evaluación se te pedirán otra vez, así que intenta retenerlas.</p>`);
  const ensayos = [];
  let intrusiones = 0;
  for (let e = 1; e <= 3; e++) {
    await presentarPalabras(area, lista, `Ensayo ${e} de 3 · Memoriza`);
    const r = await recuerdoLibre(area, lista, 90, `Ensayo ${e} de 3 · ¿Qué palabras recuerdas?`);
    ensayos.push(r.aciertos);
    intrusiones += r.intrusiones;
  }
  const finAprendizaje = Date.now();

  // 2. Dígitos directos
  paso(1);
  await avisoPaso(area, '🔢 Dígitos en orden', `
    <p>Aparecerán dígitos de uno en uno (uno por segundo). Después escríbelos <b>en el mismo orden</b>.</p>
    <p>Se empieza con 3 dígitos y se añade uno cada vez. Hay 2 intentos por longitud; la prueba termina cuando fallas los dos.</p>`);
  const directo = await spanDigitos(area, false);

  // 3. Dígitos inversos
  paso(2);
  await avisoPaso(area, '🔄 Dígitos al revés', `
    <p>Igual que antes, pero escribe los dígitos <b>en orden inverso</b>.</p>
    <p class="suave">Ejemplo: si ves 3 – 8 – 1, escribe 183.</p>`);
  const inverso = await spanDigitos(area, true);

  // 4. Span espacial
  paso(3);
  await avisoPaso(area, '🟦 Secuencia espacial', `
    <p>Unos bloques se iluminarán uno tras otro. Después tócalos <b>en el mismo orden</b>.</p>
    <p>Se empieza con 2 bloques; hay 2 intentos por longitud.</p>`);
  const espacial = await spanEspacial(area);

  // 5. Recuerdo diferido
  paso(4);
  await avisoPaso(area, '🔁 Recuerdo diferido', `
    <p>Escribe todas las palabras que recuerdes de la <b>lista del principio</b>.</p>
    <p class="suave">Esta vez no se volverá a mostrar.</p>`);
  const demoraSeg = Math.round((Date.now() - finAprendizaje) / 1000);
  const dif = await recuerdoLibre(area, lista, 90, 'Recuerdo diferido · ¿Qué palabras recuerdas?');

  const res = {
    fecha: new Date().toISOString(),
    lista: indiceLista + 1,
    ensayos,
    aprendizaje: ensayos.reduce((a, b) => a + b, 0),
    diferido: dif.aciertos,
    directo,
    inverso,
    espacial,
    intrusiones: intrusiones + dif.intrusiones,
    demoraSeg,
    duracionSeg: Math.round((Date.now() - inicio) / 1000),
  };
  Datos.d.evaluaciones.push(res);
  Datos.marcarDia();
  Datos.guardar();
  ir('evalResultado', Datos.d.evaluaciones.length - 1);
};

Vistas.evalResultado = function (i) {
  const evals = Datos.d.evaluaciones;
  const e = evals[i];
  const base = i > 0 ? evals[0] : null;
  const fila = (m) => {
    const d = base ? e[m.id] - base[m.id] : null;
    return `<tr>
      <td>${m.nombre}<br><small class="suave">${m.desc}${m.ref ? ` · Referencia orientativa en adultos: ${m.ref}` : ''}</small></td>
      <td class="num"><b>${e[m.id]}</b> / ${m.max}</td>
      ${base ? `<td class="num">${base[m.id]}</td><td class="num ${d > 0 ? 'sube' : d < 0 ? 'baja' : ''}">${d > 0 ? '+' : ''}${d}</td>` : ''}
    </tr>`;
  };
  app.innerHTML = `
    <section class="resultado tarjeta" style="max-width:720px">
      <div class="icono-juego">📋</div>
      <h1>${i === 0 ? 'Evaluación inicial completada' : `Evaluación n.º ${i + 1} completada`}</h1>
      <p class="suave">${new Date(e.fecha).toLocaleString('es')} · Lista ${e.lista} · Duración ${fmtTiempo(e.duracionSeg)}</p>
      <div class="tabla-scroll">
        <table class="tabla">
          <thead><tr><th>Medida</th><th class="num">Hoy</th>${base ? '<th class="num">Inicial</th><th class="num">Cambio</th>' : ''}</tr></thead>
          <tbody>${MEDIDAS.map(fila).join('')}</tbody>
        </table>
      </div>
      <div class="detalles">
        <div><small>Curva de aprendizaje</small><b>${e.ensayos.join(' → ')}</b></div>
        <div><small>Palabras no presentes</small><b>${e.intrusiones}</b></div>
        <div><small>Demora del recuerdo</small><b>${fmtTiempo(e.demoraSeg)}</b></div>
      </div>
      <p>${i === 0
        ? 'Esta es tu <b>línea base</b>. Entrena a diario con el repaso espaciado y el palacio de la memoria y repite la evaluación dentro de 2 semanas para comparar.'
        : 'Compara la tendencia con varias evaluaciones: un solo resultado puede variar por cansancio o concentración.'}</p>
      <p class="suave" style="font-size:.88rem">Las referencias son aproximadas, varían con la edad y no constituyen un diagnóstico. Si te preocupa tu memoria, consulta a un profesional sanitario.</p>
      <div class="fila centro">
        <button class="btn" id="verTodas">Ver todas las evaluaciones</button>
        <button class="btn primario grande" id="inicio">Empezar a entrenar</button>
      </div>
    </section>`;
  if (base && MEDIDAS.filter((m) => e[m.id] > base[m.id]).length >= 3) { confeti(); Sonido.fanfarria(); }
  $('#verTodas').addEventListener('click', () => ir('evaluacion'));
  $('#inicio').addEventListener('click', () => ir('inicio'));
};
