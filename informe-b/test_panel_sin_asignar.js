// El bloque "Sin asignar" del panel Configurar cuentas.
//
// Muestra las dos puntas sueltas entre los dos planes de cuentas y deja resolverlas ahí mismo.
// Lo que importa probar es que no mienta: que una cuenta que ya tiene fila no aparezca, que una
// fila "sin cuentas asignadas" no se cuele como si fuera una cuenta de Onvio a ubicar, y que
// asignar una cuenta la saque de la lista.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const AQUI = __dirname;

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

// ---------------------------------------------------------------- un DOM mínimo
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

let ultimoConfirm = null;
const ctx = {
  console,
  document: {
    getElementById: elemento,
    createElement: () => elemento("__nuevo__" + Math.random()),
    addEventListener() {}, querySelectorAll: () => [], querySelector: () => null,
    body: elemento("__body"),
  },
  localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  fetch: () => Promise.resolve({ json: () => Promise.resolve([]) }),
  alert(m) { ultimoConfirm = m; },
  confirm(m) { ultimoConfirm = m; return true; },
  setTimeout, clearTimeout, Blob: function () {},
  URL: { createObjectURL: () => "", revokeObjectURL() {} },
};
ctx.window = ctx;
ctx.globalThis = ctx;
vm.createContext(ctx);

const html = fs.readFileSync(path.join(AQUI, "index.html"), "utf8");
const srcs = [...html.matchAll(/<script src="([^"]+)"/g)].map(m => m[1].split("?")[0]);
for (const src of srcs) {
  if (/vendor\//.test(src) || /xlsx|exceljs/.test(src)) continue;
  const f = path.join(AQUI, src);
  if (!fs.existsSync(f)) continue;
  try {
    vm.runInContext(fs.readFileSync(f, "utf8"), ctx, { filename: src });
  } catch (e) {
    check(false, `${src} no se pudo cargar: ${e.message}`);
  }
}
check(typeof ctx.cbSinAsignar === "function", "el panel trae el bloque de lo que quedó sin asignar");
check(typeof ctx.cbConfirmarAsignacion === "function", "y la acción de ubicar una cuenta");

// ---------------------------------------------------------------- el mapeo real
const mapping = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));

// `currentMapping` y `lastResult` son `let` de app.js: eso crea un binding léxico, no una
// propiedad del global, así que asignarlos desde afuera (ctx.currentMapping = ...) crea otra
// variable y la página sigue viendo null. Hay que asignarlos DENTRO del contexto.
// Está documentado en CLAUDE.md; ya hizo perder tiempo una vez.
// Lo mismo para LEERlas: `cbMapping` y `cbCambios` son `let` de cuentas.js. Las funciones sí
// quedan colgadas del global (por eso ctx.cbSinAsignar() anda), las variables no.
const dentro = (expr) => vm.runInContext(expr, ctx);

ctx._mapeo = mapping;
const ponerEstado = (cuentas) => {
  ctx._cuentas = cuentas;
  vm.runInContext("currentMapping = _mapeo; lastResult = { cuentas: _cuentas };", ctx);
};

// Un export con: una cuenta que no tiene fila, una que sí, y una dormida que tampoco.
ponerEstado({
  "429990000": { descripcion: "Cuenta nueva de Onvio", debe: 1234.5, haber: 0 },
  "111010001": { descripcion: "Caja", debe: 100, haber: 0 },
  "424140000": { descripcion: "Canon", debe: 0, haber: 0 },
});
ctx.cbAbrirPanel();
check(dentro("cbMapping && cbMapping.length") === mapping.length,
  `el panel tomó su copia de trabajo: ${dentro("cbMapping && cbMapping.length")} filas`);

let s = ctx.cbSinAsignar();
const codigos = s.cuentas.map(c => c.code);
check(codigos.includes("429990000"), "la cuenta que ninguna fila reclama aparece");
check(!codigos.includes("111010001"), "y la que sí tiene fila (Caja) no aparece");
check(!codigos.includes("424140000"),
  "tampoco la que ya asignamos a una madre, aunque este mes esté en cero");
check(s.cuentas.find(c => c.code === "429990000").movio === true,
  "se distingue la que movió de la que no");

// Una fila "sin cuentas asignadas" NO es una cuenta de Onvio a ubicar
check(!codigos.includes("421170000"),
  '"421170000" no se cuela como cuenta de Onvio: ese número es de una fila del cliente');
const filas = s.filas.map(f => f.cliente.code);
check(filas.includes("421170000") && filas.includes("42433000"),
  `las filas que no se llenan con nada sí se listan: ${filas.join(", ")}`);
check(s.filas.every(f => /sin cuentas asignadas|no tiene ninguna subcuenta/.test(f.motivo)),
  "y cada una dice por qué");

// ---------------------------------------------------------------- ubicarla
const cuantasAntes = dentro('(cbMapping.find(x => x.code === "42102000").children || []).length');
elemento("cbSelDestino").value = "42102000";           // Honorarios profesionales, ya es madre
ctx.cbConfirmarAsignacion("429990000");
const madre = dentro('cbMapping.find(x => x.code === "42102000")');
check((madre.children || []).length === cuantasAntes + 1 &&
      madre.children.some(h => h.code === "429990000"),
  `la cuenta entra como subcuenta de "42102000" (${cuantasAntes} → ${madre.children.length})`);
check(!ctx.cbSinAsignar().cuentas.some(c => c.code === "429990000"),
  "y deja de figurar como suelta");
const cambios = dentro("cbCambios");
check(cambios.some(c => /429990000/.test(c)),
  `queda anotado para el commit: "${cambios.find(c => /429990000/.test(c))}"`);
check(elemento("btnCbGuardar").disabled === false, "y se habilita Guardar cambios");

// ---------------------------------------------------------------- una fila simple avisa
ponerEstado({ "429991111": { descripcion: "Otra suelta", debe: 5, haber: 0 } });
ultimoConfirm = null;
elemento("cbSelDestino").value = "111010001";           // Caja, fila simple
ctx.cbConfirmarAsignacion("429991111");
const caja = dentro('cbMapping.find(x => x.code === "111010001")');
check(/pasar? a ser cuenta madre/.test(ultimoConfirm || ""),
  "asignar a una fila simple avisa que va a pasar a ser cuenta madre");
check(caja.type === "parent" && caja.children.map(h => h.code).join(",") === "111010001,429991111",
  "y su propia cuenta queda como primera subcuenta, para no perderla");

// ---------------------------------------------------------------- nada se guardó solo
const enDisco = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));
check(JSON.stringify(enDisco) === JSON.stringify(mapping),
  "el mapping del disco no se tocó: los cambios esperan a Guardar");

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
