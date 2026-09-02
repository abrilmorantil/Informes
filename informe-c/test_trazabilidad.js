// La hoja "De dónde sale cada saldo" del balance.
//
// Sigue el camino completo de cada importe —Onvio → Hoja1 → SALDOS → Anexo II o Nota 4— y
// marca en qué escalón se corta. Lo que se prueba acá es justamente que los cuatro escalones
// se reconozcan, incluidos los dos que no son una referencia directa:
//
//   - una SUBCUENTA de resultados llega al Anexo II a través del subtotal de su madre;
//   - un PROVEEDOR llega a la Nota 4 a través de `SALDOS!I117 = SUM(G63:G117)`.
//
// Si esos dos saltos no se siguen, la hoja diría que 280 cuentas "no entran a ningún estado",
// que es lo contrario de lo que pasa, y el aviso no serviría para nada.
//
//   node informe-c/test_trazabilidad.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const M = require(path.join(AQUI, "motor_balances.js"));
for (const k of Object.keys(M)) global[k] = M[k];
const clas = require(path.join(AQUI, "clasificacion.js"));
const tz = require(path.join(AQUI, "trazabilidad.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const T = (c) => {
  const v = c.value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return typeof v.result === "string" ? v.result : "";
    return String(v.result ?? "");
  }
  return typeof v === "string" ? v : String(v);
};

(async () => {
  global.clasificacion = clas.indexarClasificacion(
    JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8")));

  console.log("=== 1) las filas que nombra una fórmula, con rangos ===");
  {
    const f = tz.tzFilasQueNombra("SUM(G63:G65)");
    check(f.has(63) && f.has(64) && f.has(65) && f.size === 3, "un rango abarca todas sus filas");
    const g = tz.tzFilasQueNombra("+SALDOS!G12+SALDOS!G20");
    check(g.has(12) && g.has(20) && g.size === 2, "dos referencias sueltas");
    const h = tz.tzFilasQueNombra("=SUM(F10:I10)");
    check(h.has(10) && h.size === 1, "un rango horizontal es una sola fila");
  }

  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const mapeo = derivarMapeoMaestro(wb, "pesos", global.clasificacion);
  const filas = tz.tzFilas(wb, "pesos", mapeo, madresResultados(wb, "pesos"));

  console.log("\n=== 2) una fila por cuenta, con los cuatro escalones ===");
  check(filas.length === Object.keys(mapeo.cuentas).length,
    `hay una fila por cada cuenta del maestro (${filas.length})`);
  const caja = filas.find(f => f.codigo === "111010001");
  check(!!caja && caja.hoja1Fila && caja.saldosFila && caja.destinos.length,
    `"Caja" tiene los cuatro: Hoja1 f${caja && caja.hoja1Fila} → SALDOS f${caja && caja.saldosFila}` +
    ` → ${caja && caja.destinos[0] && caja.destinos[0].hoja} ${caja && caja.destinos[0] && caja.destinos[0].celda}`);

  console.log("\n=== 3) los dos saltos que no son referencia directa ===");
  const conMadre = filas.filter(f => f.madreCodigo && f.destinos.length);
  check(conMadre.length > 50,
    `${conMadre.length} subcuentas llegan al Anexo II por el subtotal de su madre`);
  const unaHija = conMadre[0];
  check(unaHija.destinos[0].hoja === "Anexo II",
    `ej: ${unaHija.codigo} (madre ${unaHija.madreCodigo}) va a "${unaHija.destinos[0].etiqueta}"`);

  const prov = filas.filter(f => /^21101/.test(f.codigo));
  const provConDestino = prov.filter(f => f.destinos.length);
  check(prov.length > 30 && provConDestino.length === prov.length,
    `los ${prov.length} proveedores llegan a la Nota 4 por el subtotal del bloque`);
  check(provConDestino[0].destinos[0].viaSubtotal,
    `y queda anotado por qué celda pasan (${provConDestino[0].destinos[0].viaSubtotal})`);

  console.log("\n=== 4) qué marca como cortado ===");
  const sinH1 = filas.filter(f => f.encuentraEnHoja1 === false);
  const sinDest = filas.filter(f => f.sinDestino);
  console.log(`       ${sinH1.length} no están en Hoja1 · ${sinDest.length} no entran a ningún estado`);
  check(sinH1.every(f => f.claveBuscada && f.hoja1Fila === null),
    "las que no están en Hoja1 dicen qué texto buscaban y no traen fila");
  // una cuenta que SÍ está no puede figurar como cortada
  check(!sinH1.some(f => f.codigo === "111010001") && !sinDest.some(f => f.codigo === "111010001"),
    '"Caja", que está entera, no figura entre las cortadas');
  check(sinDest.length < filas.length * 0.1,
    `las cortadas son pocas (${sinDest.length} de ${filas.length}): si fueran cientos, el aviso no serviría`);

  console.log("\n=== 5) la hoja ===");
  const r = tz.tzEscribirHoja(wb, "pesos");
  const ws = wb.getWorksheet("De dónde sale cada saldo");
  check(!!ws, "la hoja se agrega al libro");
  check(r.filas === filas.length, `con ${r.filas} filas`);
  check(T(ws.getCell(4, 1)) === "Código" && T(ws.getCell(4, 11)) === "¿Dónde se corta?",
    "con sus encabezados");
  check(!!ws.autoFilter, "y con el filtro puesto, que es como se usa");

  // escribirla dos veces no puede duplicarla: la descarga puede correrse más de una vez
  tz.tzEscribirHoja(wb, "pesos");
  check(wb.worksheets.filter(w => w.name === "De dónde sale cada saldo").length === 1,
    "escribirla de nuevo la reemplaza, no la duplica");

  console.log("\n=== 6) y en dólares, que tiene otra geometría ===");
  {
    const wbu = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_dolares.xlsx")));
    const mu = derivarMapeoMaestro(wbu, "dolares", global.clasificacion);
    const fu = tz.tzFilas(wbu, "dolares", mu, madresResultados(wbu, "dolares"));
    check(fu.length === Object.keys(mu.cuentas).length, `${fu.length} filas`);
    check(fu.some(f => f.destinos.length && f.destinos[0].hoja === "Anexo II"),
      "las de gastos llegan al Anexo II");
    check(fu.filter(f => f.hoja1Fila).length > 50,
      `${fu.filter(f => f.hoja1Fila).length} encuentran su cuenta en Hoja1`);
  }

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
