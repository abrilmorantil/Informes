// El Anexo II: en qué columna cae el gasto de cada cuenta madre.
//
// El invariante es el que importa: cada madre entra EXACTAMENTE UNA VEZ. Dos veces cuenta el
// gasto doble; cero lo hace desaparecer del estado de resultados. Y como el Anexo II lee el
// SUBTOTAL de la madre y nunca las subcuentas, agregar una referencia de más no se ve en
// ningún control de suma: el balance sigue cerrando y el resultado queda mal.
//
//   node informe-c/test_anexo2.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const motor = require(path.join(AQUI, "motor_balances.js"));
const a2 = require(path.join(AQUI, "anexo2.js"));

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };
const F = (ws, r, c) => String(ws.getCell(r, c).formula || "");

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const ws = wb.getWorksheet("Anexo II");
  const madres = motor.madresResultados(wb, "pesos");

  // ---------------------------------------------------------------- lo que hay hoy
  const cols = a2.a2Columnas(wb);
  check(cols.length === 4, `las 4 columnas de gasto salen del archivo: ${cols.map(c => c.letra + "=" + c.nombre).join(", ")}`);

  const v = a2.a2Verificar(wb, madres);
  console.log(`       ${v.total} cuentas madre · ${v.unaVez} entran una sola vez · ` +
              `${v.dobles.length} dos o más · ${v.sinAnexo.length} ninguna`);
  check(v.ok, "el maestro está sano: cada madre entra exactamente una vez");

  const { lineas } = a2.a2Mapa(wb, madres);
  const reparto = {};
  for (const l of lineas) for (const d of l.donde) reparto[d.columna] = (reparto[d.columna] || 0) + 1;
  check(Object.values(reparto).reduce((a, b) => a + b, 0) === madres.length,
    `reparto: ${Object.entries(reparto).map(([k, n]) => k + " " + n).join(", ")}`);

  // un concepto partido en dos columnas es normal y tiene que verse
  const porConcepto = {};
  for (const l of lineas) for (const d of l.donde) {
    (porConcepto[d.concepto] = porConcepto[d.concepto] || new Set()).add(d.columna);
  }
  const partidos = Object.entries(porConcepto).filter(([, s]) => s.size > 1);
  check(partidos.length > 0,
    `hay ${partidos.length} concepto(s) repartido(s) en más de una columna, ej. "${partidos[0] && partidos[0][0]}"`);

  // ---------------------------------------------------------------- mover una madre de columna
  const admin = cols.find(c => /ADMINISTRACION/i.test(c.nombre));
  const explo = cols.find(c => /EXPLORACION/i.test(c.nombre));
  const victima = lineas.find(l => l.donde.length === 1 && l.donde[0].columna === admin.nombre &&
    F(ws, l.donde[0].anexoFila, l.donde[0].col) === `+SALDOS!G${l.filaSaldos}`);
  check(!!victima, `hay una madre sola en su celda para mover: "${victima && victima.nombre}"`);

  const desde = victima.donde[0];
  const r = a2.a2Mover({ wb, madres, filaSaldos: victima.filaSaldos,
    anexoFilaDestino: desde.anexoFila, colDestino: explo.col });
  check(r.hasta.columna === explo.nombre, `pasó de ${admin.nombre} a ${explo.nombre}`);
  check(F(ws, desde.anexoFila, admin.col) === "", "la celda de origen quedó vacía, no con un + suelto");
  check(F(ws, desde.anexoFila, explo.col).includes(`SALDOS!G${victima.filaSaldos}`),
    `y la de destino la tiene: ${F(ws, desde.anexoFila, explo.col)}`);
  check(a2.a2Verificar(wb, madres).ok, "el invariante sigue en pie después de mover");

  // vuelve a su lugar
  a2.a2Mover({ wb, madres, filaSaldos: victima.filaSaldos, anexoFilaDestino: desde.anexoFila, colDestino: admin.col });
  check(F(ws, desde.anexoFila, admin.col) === `+SALDOS!G${victima.filaSaldos}`,
    "y vuelve exactamente como estaba");

  // ---------------------------------------------------------------- moverla a una celda con otra cuenta
  const conVarias = lineas.find(l => l.donde.length === 1 &&
    (F(ws, l.donde[0].anexoFila, l.donde[0].col).match(/SALDOS!/g) || []).length > 1);
  if (conVarias) {
    const d = conVarias.donde[0];
    const antesF = F(ws, d.anexoFila, d.col);
    a2.a2Mover({ wb, madres, filaSaldos: conVarias.filaSaldos, anexoFilaDestino: desde.anexoFila, colDestino: explo.col });
    const ahora = F(ws, d.anexoFila, d.col);
    check(!ahora.includes(`SALDOS!G${conVarias.filaSaldos}`) && ahora.includes("SALDOS!"),
      `sacarla de una celda con varias deja las otras: "${antesF}" -> "${ahora}"`);
    check(a2.a2Verificar(wb, madres).ok, "y el invariante sigue bien");
    a2.a2Mover({ wb, madres, filaSaldos: conVarias.filaSaldos, anexoFilaDestino: d.anexoFila, colDestino: d.col });
    check(a2.a2Verificar(wb, madres).ok, "vuelve sin romper nada");
  }

  // ---------------------------------------------------------------- lo que NO se puede hacer
  let tiro = null;
  try {
    a2.a2Mover({ wb, madres, filaSaldos: victima.filaSaldos, anexoFilaDestino: desde.anexoFila, colDestino: 99 });
  } catch (e) { tiro = e.message; }
  check(!!tiro && /columna de gastos/.test(tiro), "una columna que no es de gastos se rechaza");
  check(F(ws, desde.anexoFila, admin.col) === `+SALDOS!G${victima.filaSaldos}`, "y no se tocó nada");

  tiro = null;
  try {
    a2.a2Mover({ wb, madres, filaSaldos: 99999, anexoFilaDestino: desde.anexoFila, colDestino: explo.col });
  } catch (e) { tiro = e.message; }
  check(!!tiro && /no está en el maestro/.test(tiro), "mover algo que no es una madre se rechaza");

  // ---------------------------------------------------------------- el maestro del disco, intacto
  const disco = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const wsD = disco.getWorksheet("Anexo II");
  check(F(wsD, desde.anexoFila, admin.col) === `+SALDOS!G${victima.filaSaldos}`,
    "el maestro del disco quedó intacto");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
