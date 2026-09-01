// El botón "Guardar cambios" del panel Configurar cuentas.
//
// Existe porque ese botón NUNCA pudo guardar. Pedía un campo `owner` en la configuración de
// GitHub, y la configuración no tiene `owner`: se guarda {token, repo, branch, path}, con el
// repositorio entero en `repo` ("usuaria/Informes"). La condición fallaba siempre, con GitHub
// configurado o sin configurar, y lo único que se veía era "Configurá GitHub antes de
// guardar" — un cartel que apunta al lugar equivocado.
//
// La forma de que no vuelva a pasar es que el panel use el MISMO control que el resto de la
// app (`hasGhSettings`), no una copia con sus propias reglas. Eso es lo que se prueba acá.
//
//   node informe-b/test_guardar_panel.js
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const AQUI = __dirname;

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

// ---------------------------------------------------------------- DOM y localStorage falsos
const elementos = {};
function elemento(id) {
  if (!elementos[id]) {
    elementos[id] = {
      id, innerHTML: "", textContent: "", value: "", style: {}, files: [], disabled: false,
      options: [],
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

const almacen = {};
let pedidos = [];
let contenidoRemoto = [];      // lo que "hay" en GitHub cuando el panel relee el sha
const ctx = {
  console,
  document: {
    getElementById: elemento,
    createElement: () => elemento("__nuevo__" + Math.random()),
    addEventListener() {}, querySelectorAll: () => [], querySelector: () => null,
    body: elemento("__body"),
  },
  localStorage: {
    getItem: (k) => (k in almacen ? almacen[k] : null),
    setItem: (k, v) => { almacen[k] = String(v); },
    removeItem: (k) => { delete almacen[k]; },
  },
  // Se intercepta la red: no se toca GitHub de verdad. El GET tiene que devolver lo mismo
  // que devuelve la Contents API —el JSON en base64— porque ghGetFile lo decodifica y lo
  // parsea; si no, el guardado muere ahí y nunca se llega al PUT.
  fetch: (url, opciones) => {
    const metodo = (opciones && opciones.method) || "GET";
    pedidos.push({ url: String(url), metodo });
    const contenido = Buffer.from(JSON.stringify(contenidoRemoto), "utf8").toString("base64");
    return Promise.resolve({
      ok: true, status: 200,
      json: () => Promise.resolve({ sha: "sha-de-mentira", content: contenido, encoding: "base64" }),
      text: () => Promise.resolve(""),
    });
  },
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
  if (/vendor\//.test(src) || /xlsx|exceljs/.test(src)) continue;
  const f = path.join(AQUI, src);
  if (!fs.existsSync(f)) continue;
  try { vm.runInContext(fs.readFileSync(f, "utf8"), ctx, { filename: src }); }
  catch (e) { check(false, `${src} no se pudo cargar: ${e.message}`); }
}

const dentro = (expr) => vm.runInContext(expr, ctx);
const estado = () => elemento("cbStatus").innerHTML;

// ---------------------------------------------------------------- la forma real de guardarla
// Es la que escribe saveSettings() en app.js: sin `owner`.
const AJUSTES = { token: "ghp_deMentira", repo: "abrilmorantil/Informes", branch: "main", path: "mapping.json" };
ctx.saveGhSettings(AJUSTES);
check(!("owner" in AJUSTES), "la configuración de GitHub no tiene campo `owner` — nunca lo tuvo");
check(ctx.hasGhSettings() === true, "y con token + repo, la app la da por configurada");

// ---------------------------------------------------------------- sin configurar: frena
ctx.localStorage.removeItem("balcomp_gh_settings");
const mapping = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));
ctx._mapeo = mapping;
contenidoRemoto = mapping;
dentro("currentMapping = _mapeo; lastResult = null;");
ctx.cbAbrirPanel();
elemento("cbStatus").innerHTML = "";
pedidos = [];
(async () => {
  await ctx.cbGuardar();
  check(/Configur/i.test(estado()), "sin configurar GitHub, avisa que hay que configurarlo");
  check(pedidos.length === 0, "y no sale a la red");

  // ------------------------------------------------------------- configurado: guarda
  ctx.saveGhSettings(AJUSTES);
  elemento("cbStatus").innerHTML = "";
  pedidos = [];
  dentro("cbCambios = ['un cambio de prueba'];");
  await ctx.cbGuardar();

  check(!/Configurá GitHub/i.test(estado()),
    `con GitHub configurado YA NO pide configurarlo (decía: "${estado().replace(/<[^>]*>/g, "").trim()}")`);
  check(pedidos.some(p => p.metodo === "PUT"),
    "sale a guardar de verdad: " + (pedidos.map(p => p.metodo).join(", ") || "no hizo ningún pedido"));
  check(/ok/.test(estado()) && /[Gg]uardado/.test(estado()), "y avisa que quedó guardado");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
