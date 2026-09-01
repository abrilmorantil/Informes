// Resolver un centro de costo desde la pantalla tiene que REHACER el balance.
//
// El error, con el export de agosto 2026: se declaraba la equivalencia de "Proyecto Lonco
// Vaca- Palenque", quedaba guardada en GitHub, la tarjeta desaparecía como si estuviera todo
// bien... y el archivo que se descargaba seguía 487,94 corto, porque el borrador ya estaba
// armado con el mapeo viejo y nadie lo rehacía.
//
// Y el cartel de "listo" se borraba solo justo después de escribirlo, así que no quedaba
// ninguna señal de que hubiera funcionado: se apretó tres veces y quedaron tres commits
// iguales en el repositorio.
//
//   node informe-a/test_resolver_centro.js
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const AQUI = __dirname;

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

// ---------------------------------------------------------------- navegador de mentira
const elementos = {};
function elemento(id) {
  if (!elementos[id]) {
    elementos[id] = {
      id, innerHTML: "", textContent: "", value: "", style: {}, files: [], disabled: false,
      dataset: {}, options: [], parentNode: null,
      classList: {
        _c: new Set(),
        add(...x) { x.forEach(c => this._c.add(c)); },
        remove(...x) { x.forEach(c => this._c.delete(c)); },
        toggle(c, on) { on ? this._c.add(c) : this._c.delete(c); },
        contains(c) { return this._c.has(c); },
      },
      addEventListener() {}, appendChild() {}, focus() {}, scrollIntoView() {},
      querySelectorAll() { return []; }, querySelector() { return null; },
    };
  }
  return elementos[id];
}

const guardados = [];
const ctx = {
  console,
  document: {
    getElementById: elemento,
    createElement: () => elemento("__n" + Math.random()),
    addEventListener() {},
    querySelectorAll: (sel) => (sel === "#centrosBody button" ? [elemento("__btn1"), elemento("__btn2")] : []),
    querySelector: () => null,
    body: elemento("__body"),
  },
  localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  fetch: () => Promise.resolve({ ok: true, json: () => Promise.resolve({}) }),
  alert() {}, confirm: () => true, setTimeout, clearTimeout, Blob: function () {},
  URL: { createObjectURL: () => "", revokeObjectURL() {} },
  TextEncoder, TextDecoder,
  btoa: (s) => Buffer.from(s, "binary").toString("base64"),
  atob: (s) => Buffer.from(s, "base64").toString("binary"),
};
ctx.window = ctx;
ctx.globalThis = ctx;
vm.createContext(ctx);

const html = fs.readFileSync(path.join(AQUI, "index.html"), "utf8");
for (const src of [...html.matchAll(/<script src="([^"]+)"/g)].map(m => m[1].split("?")[0])) {
  if (/vendor\//.test(src)) continue;
  const f = path.join(AQUI, src);
  if (!fs.existsSync(f)) continue;
  try { vm.runInContext(fs.readFileSync(f, "utf8"), ctx, { filename: src }); }
  catch (e) { check(false, `${src} no se pudo cargar: ${e.message}`); }
}

const dentro = (expr) => vm.runInContext(expr, ctx);

// El desplegable con buscador es decoración: necesita un DOM de verdad y no es lo que se
// prueba acá. Se neutraliza para que el <select> quede tal cual y se le pueda poner un valor.
dentro("conBuscador = function (sel) { return sel; };");

// ---------------------------------------------------------------- se arma la situación real
const mapeo = JSON.parse(fs.readFileSync(path.join(AQUI, "mapeo.json"), "utf8"));
// se saca la equivalencia, para reproducir el archivo tal como estaba antes de resolverla
delete mapeo.cc_nombres_onvio["PROYECTO LONCO VACA- PALENQUE"];

const CRUDO = "Proyecto Lonco Vaca- Palenque";
ctx._mapeo = mapeo;
ctx._lineas = [
  { cc_codigo: "19", cc_nombre_onvio: CRUDO, cuenta_codigo: "426120000",
    cuenta_label: "Movilidad y Viaticos Gs Campo", debe: 487.94, haber: 0, saldo: 487.94 },
  { cc_codigo: "1", cc_nombre_onvio: "Oficina General", cuenta_codigo: "426120000",
    cuenta_label: "Movilidad y Viaticos Gs Campo", debe: 100, haber: 0, saldo: 100 },
];
// `guardarTodo` y `correrMotor` se interceptan: no se toca GitHub ni se arma un Excel.
dentro(`
  currentMapping = null;
  mapeoGuardado = _mapeo;
  lineas = _lineas;
  bufferBase = new ArrayBuffer(8);
  estado = { ultimo_mes_cerrado: "2026-07", historial: [] };
  guardarTodo = async function (args) { _guardados.push(args.mensaje); };
  correrMotor = async function () { _corridas.push(mapeoGuardado.cc_nombres_onvio["${"PROYECTO LONCO VACA- PALENQUE"}"] || null); };
`);
ctx._guardados = guardados;
const corridas = [];
ctx._corridas = corridas;
dentro("guardarTodo = async function (a) { _guardados.push(a.mensaje); };");
dentro("correrMotor = async function () { _corridas.push(mapeoGuardado.cc_nombres_onvio['PROYECTO LONCO VACA- PALENQUE'] || null); };");

(async () => {
  // --- la tarjeta aparece con el centro sin resolver
  const det = ctx.detectarPendientes(ctx._lineas, mapeo);
  check(det.sinCc.length === 1 && det.sinCc[0].nombre === CRUDO,
    `la tarjeta muestra "${CRUDO}" con ${det.sinCc[0] && det.sinCc[0].saldo} sin cargar`);
  ctx.ccMostrar(det.sinCc);
  ctx.ccElegir(ctx.ccId(CRUDO), "existente");
  check(!elemento("cardCentros").classList.contains("hidden"), "y la tarjeta queda visible");

  // --- se elige el proyecto y se declara
  elemento("sel_" + ctx.ccId(CRUDO)).value = "LONCO VACA - PELENQUE";
  await ctx.ccDeclarar(ctx.ccId(CRUDO));

  check(mapeo.cc_nombres_onvio["PROYECTO LONCO VACA- PALENQUE"] === "LONCO VACA - PELENQUE",
    "la equivalencia queda declarada en el mapeo");
  check(guardados.length === 1, `se guardó una sola vez (${guardados.length}): "${guardados[0]}"`);

  // --- LO QUE FALLABA: el balance tiene que rehacerse, y con el mapeo YA actualizado
  check(corridas.length === 1, `el balance se rehizo (${corridas.length} corrida(s))`);
  check(corridas[0] === "LONCO VACA - PELENQUE",
    "y se rehizo con la equivalencia ya aplicada, no con el mapeo viejo");

  // --- el cartel de "listo" sigue en pantalla
  const st = elemento("centrosStatus").innerHTML;
  check(/status-msg ok/.test(st) && /Declarado/.test(st),
    "el cartel de que quedó declarado NO se borra");
  check(/rehizo/i.test(st), "y avisa que el balance se rehizo");

  // --- la tarjeta desaparece porque ya no queda nada sin resolver
  check(elemento("cardCentros").classList.contains("hidden"), "la tarjeta se cierra sola");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
