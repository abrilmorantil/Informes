// ¿La página arranca de verdad?
//
// Los .js del sitio se cargan como <script> sueltos, así que comparten UN solo ámbito
// global: dos `const` con el mismo nombre en archivos distintos son un SyntaxError que
// tumba el segundo archivo entero, y el único síntoma es un "X is not defined" al usarlo.
// Las pruebas normales de Node NO ven esto, porque ahí cada archivo es su propio módulo.
//
// Por eso este test carga los scripts en el orden exacto de index.html dentro de UN
// contexto (vm), con lo mínimo de navegador simulado, y comprueba que las funciones que
// la página llama desde sus onclick quedaron definidas.
//
//   node informe-a/test_arranque.js
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const AQUI = __dirname;
const SITIO = path.join(AQUI, "..");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

// --- lo mínimo de navegador para que los scripts terminen de cargar -----------
function elemento() {
  const el = {
    value: "", innerHTML: "", textContent: "", disabled: false, files: [],
    classList: { toggle() {}, add() {}, remove() {}, contains() { return false; } },
    style: {}, dataset: {},
    addEventListener() {}, removeEventListener() {}, appendChild() {}, remove() {},
    setAttribute() {}, removeAttribute() {}, getAttribute() { return null; },
    querySelector() { return elemento(); }, querySelectorAll() { return []; },
    closest() { return null }, focus() {}, scrollIntoView() {}, insertBefore() {},
    parentNode: null, children: [], options: [],
  };
  el.parentNode = { insertBefore() {}, removeChild() {} };
  return el;
}
const documento = {
  getElementById: () => elemento(),
  querySelector: () => elemento(),
  querySelectorAll: () => [],
  createElement: () => elemento(),
  addEventListener() {},
  body: elemento(),
};
const contexto = {
  console, document: documento, localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  fetch: () => Promise.reject(new Error("sin red en el test")),
  alert() {}, confirm: () => true, prompt: () => null,
  setTimeout, clearTimeout, location: { reload() {}, href: "" },
  navigator: { userAgent: "node" }, atob: (s) => Buffer.from(s, "base64").toString("binary"),
  btoa: (s) => Buffer.from(s, "binary").toString("base64"),
  URL: { createObjectURL: () => "blob:x", revokeObjectURL() {} },
  Blob: function () {}, TextEncoder, TextDecoder, Uint8Array, ArrayBuffer, Promise, Date, Math, JSON,
};
contexto.window = contexto;
contexto.globalThis = contexto;
contexto.self = contexto;
vm.createContext(contexto);

// --- el orden es el de index.html --------------------------------------------
const html = fs.readFileSync(path.join(AQUI, "index.html"), "utf8");
// El `?v=...` que llevan los src es para que el navegador no reuse la version vieja
// al publicar; como ruta de archivo hay que sacarlo.
const scripts = [...html.matchAll(/<script src="([^"]+)"/g)].map(m => m[1].split("?")[0]);
console.log(`index.html carga ${scripts.length} scripts`);

for (const src of scripts) {
  const archivo = path.resolve(AQUI, src);
  if (!fs.existsSync(archivo)) { check(false, `falta el archivo ${src}`); continue; }
  // Las librerías de vendor son pesadas y no aportan a esta prueba, pero sí tienen que
  // dejar sus globales (ExcelJS, XLSX) porque los otros scripts las usan al cargar.
  try {
    vm.runInContext(fs.readFileSync(archivo, "utf8"), contexto, { filename: src });
  } catch (e) {
    check(false, `${src} no se ejecutó: ${e.message}`);
  }
}

// --- lo que la página llama desde sus onclick --------------------------------
const DEL_HTML = [...html.matchAll(/on(?:click|change|input)="([a-zA-Z_$][\w$]*)\(/g)]
  .map(m => m[1]);
const esperadas = [...new Set(DEL_HTML)];
console.log(`\nfunciones que index.html llama desde sus onclick: ${esperadas.length}`);
for (const f of esperadas) {
  check(typeof contexto[f] === "function", `${f}() está definida`);
}

// --- las que se llaman desde el HTML que arma el panel -----------------------
// El regex de arriba sólo ve los onclick escritos en index.html. Estas viven en los
// template strings de categorias.js, así que no aparecen ahí y son igual de rompibles.
console.log("\nfunciones que el panel arma en sus botones:");
for (const f of ["catEditarCategoria", "catAplicarCategoria", "catCancelarEdicion",
                 "catRenombrar", "catQuitar"]) {
  check(typeof contexto[f] === "function", `${f}() está definida`);
}

console.log("\npiezas del motor que usa el panel:");
for (const f of ["moverCuentaDeCategoria", "quitarCuentaDeDistribucion", "renombrarCuenta",
                 "insertarCuentaEnBalance", "mapaDeDistribucion", "quitarRefDistDeGastos",
                 "refsDeCuentaEnDist", "conBuscadorTodos", "abrirWorkbook", "guardarTodo"]) {
  check(typeof contexto[f] === "function", `${f}() quedó en el ámbito global`);
}

// `const` y `let` de nivel superior NO quedan como propiedad del objeto global: son
// bindings léxicos. Se los alcanza evaluando dentro del mismo contexto.
const enContexto = (expr) => vm.runInContext(expr, contexto);
check(enContexto("typeof NO_INCLUIR") === "string" && enContexto("NO_INCLUIR") === "__no_incluir__",
      "NO_INCLUIR quedó definida");
check(enContexto("typeof catWb") !== "undefined", "el estado del panel (catWb) quedó declarado");

console.log(`\n${fallos ? `${fallos} FALLA(S)` : "todo OK"}`);
process.exit(fallos ? 1 : 0);
