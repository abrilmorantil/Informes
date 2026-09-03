// La guía por pasos del Informe 04: un paso a la vez, con la posibilidad de volver.
//
// No reescribe el circuito. La app sigue mostrando y escondiendo sus tarjetas igual que
// siempre —`mostrar(id, v)`— y este módulo se entera de eso y decide cuál es la que se ve.
// Así el motor y su estado quedan intactos: lo único que cambia es la presentación.
//
// Reglas, decididas con la usuaria:
//   · se ve el paso del momento y se puede volver al anterior; no hay lista apilada
//   · el paso que no hace falta NO existe: si no hay cuentas nuevas, no se muestra vacío,
//     desaparece, y la numeración lo refleja ("Paso 2 de 3")
//   · avanzar es del circuito, no del usuario: el paso siguiente aparece cuando la app lo
//     habilita. Volver sí es libre.
//
// Lo que NO son pasos: cargar el maestro la primera vez, el mes anterior del EE RR, los
// paneles de configuración y el historial. Son mantenimiento o consulta, no el cierre del mes.

const PASOS = [
  { id: "cardExport", nombre: "Carga del balance" },
  { id: "cardNuevas", nombre: "Cuentas nuevas" },
  { id: "cardRevisionUsd", nombre: "Cuentas en dólares" },
  { id: "cardResultado", nombre: "Controles" },
  { id: "cardCierre", nombre: "Revisión y aprobación" },
];
const PASOS_IDS = PASOS.map(p => p.id);

// ------------------------------------------------------------------ el estado, sin DOM
//
// Se mantiene aparte de la pantalla para poder probarlo: es donde están las decisiones.
function pasosNuevoEstado() {
  return { disponibles: [], actual: null, nuevos: [] };
}

const pasosOrden = (a, b) => PASOS_IDS.indexOf(a) - PASOS_IDS.indexOf(b);

// La app habilitó (o deshabilitó) una tarjeta. Sólo anota; no decide adónde ir.
//
// La decisión no se puede tomar acá porque la app habilita las tarjetas de a una, y al terminar
// la corrida habilita DOS seguidas: los controles y la revisión. Mirando cada aviso por
// separado, la segunda pisaría a la primera y el paso de controles se saltearía sin que nadie
// lo vea. Por eso la tanda es explícita: se anota qué se habilitó y se decide al cerrarla.
function pasosAlMostrar(est, id, visible) {
  if (PASOS_IDS.indexOf(id) < 0) return est;
  const disponibles = est.disponibles.filter(x => x !== id);
  if (visible) disponibles.push(id);
  disponibles.sort(pasosOrden);
  const nuevos = est.nuevos.filter(x => x !== id);
  if (visible && est.disponibles.indexOf(id) < 0) nuevos.push(id);
  nuevos.sort(pasosOrden);
  return { disponibles, actual: est.actual, nuevos };
}

// Se cierra la tanda: recién acá se decide qué paso se ve.
//
//   · si todavía no había ninguno, el primero
//   · si se habilitaron pasos más adelante, el PRIMERO de ellos
//   · si el paso en el que estábamos se apagó —el circuito lo dio por terminado—, el siguiente
//     que haya quedado, o el último si no hay ninguno adelante
function pasosAsentar(est) {
  let actual = est.actual;
  const hay = est.disponibles;
  if (!hay.length) return { disponibles: hay, actual: null, nuevos: [] };

  if (actual === null) {
    actual = hay[0];
  } else {
    const adelante = est.nuevos.filter(x => pasosOrden(x, actual) > 0);
    if (adelante.length) actual = adelante[0];
    else if (hay.indexOf(actual) < 0) {
      const quedan = hay.filter(x => pasosOrden(x, actual) > 0);
      actual = quedan.length ? quedan[0] : hay[hay.length - 1];
    }
  }
  return { disponibles: hay, actual, nuevos: [] };
}

function pasosVolver(est) {
  const i = est.disponibles.indexOf(est.actual);
  return i > 0 ? { ...est, actual: est.disponibles[i - 1] } : est;
}

function pasosIrA(est, id) {
  return est.disponibles.indexOf(id) >= 0 ? { ...est, actual: id } : est;
}

// Dónde estoy, para dibujar. `indice` es 0-based sobre los pasos que HOY existen.
function pasosUbicacion(est) {
  const i = est.disponibles.indexOf(est.actual);
  const anterior = i > 0 ? est.disponibles[i - 1] : null;
  return {
    indice: i,
    total: est.disponibles.length,
    anterior,
    nombreAnterior: anterior ? (PASOS.find(p => p.id === anterior) || {}).nombre : null,
  };
}

// ------------------------------------------------------------------ la pantalla
let pasosEstado = pasosNuevoEstado();

// La llama `mostrar()` de app.js. Cuando la app habilita varias tarjetas seguidas, el redibujo
// se hace una sola vez al final del tick: si no, se vería pasar el paso intermedio.
let pasosRedibujoPedido = false;
function pasosAviso(id, visible) {
  if (PASOS_IDS.indexOf(id) < 0) return;
  pasosEstado = pasosAlMostrar(pasosEstado, id, visible);
  if (pasosRedibujoPedido) return;
  pasosRedibujoPedido = true;
  (typeof queueMicrotask === "function" ? queueMicrotask : setTimeout)(() => {
    pasosRedibujoPedido = false;
    pasosEstado = pasosAsentar(pasosEstado);   // se cierra la tanda y recién ahí se decide
    pasosPintar();
  }, 0);
}

function pasosPintar() {
  if (typeof document === "undefined") return;
  for (const p of PASOS) {
    const el = document.getElementById(p.id);
    if (!el) continue;
    // sólo la del paso del momento se ve; las demás siguen existiendo, con su estado intacto
    el.classList.toggle("hidden", p.id !== pasosEstado.actual);
  }
  const barra = document.getElementById("pasosBarra");
  if (!barra) return;
  const u = pasosUbicacion(pasosEstado);
  if (u.indice < 0) { barra.classList.add("hidden"); return; }
  barra.classList.remove("hidden");
  const tramos = [];
  for (let i = 0; i < u.total; i++) {
    tramos.push(`<i class="${i < u.indice ? "hecho" : (i === u.indice ? "aca" : "")}"></i>`);
  }
  barra.innerHTML =
    (u.anterior
      ? `<button class="pasos-volver" onclick="pasosAtras()">← ${u.nombreAnterior}</button>`
      : "") +
    `<span class="pasos-sp"></span>` +
    `<span class="pasos-cuantos">Paso ${u.indice + 1} de ${u.total}</span>` +
    `<span class="pasos-tramos">${tramos.join("")}</span>`;
}

function pasosAtras() {
  pasosEstado = pasosVolver(pasosEstado);
  pasosPintar();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// Avanzar a mano, para el paso de controles: ahí no hay nada que la app tenga que habilitar,
// simplemente se leyó y se sigue.
function pasosSeguir() {
  const i = pasosEstado.disponibles.indexOf(pasosEstado.actual);
  if (i >= 0 && i < pasosEstado.disponibles.length - 1) {
    pasosEstado = pasosIrA(pasosEstado, pasosEstado.disponibles[i + 1]);
    pasosPintar();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
}

if (typeof module !== "undefined") {
  module.exports = {
    PASOS, PASOS_IDS, pasosNuevoEstado, pasosAlMostrar, pasosAsentar, pasosVolver, pasosIrA,
    pasosUbicacion,
  };
}
