// La columna de TOTALES SALDOS tiene que sumar TODOS los centros de costo.
//
// Es una lista explícita de celdas (+E7+H7+K7+...), no un rango, así que un centro puede
// quedar afuera y la columna sigue dando un número — de menos, sin avisar. En el archivo real
// pasaba en 3 filas que se saltean SAMENTA. Es la cuarta vez que aparece el mismo problema en
// este archivo (la fila TOTALES, esta columna, y el total de Gastos Acumulados), así que la
// regla no se parchea: se reconstruye desde el mapeo en cada corrida.
//
//   node informe-a/test_totales_saldos.js
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");

global.XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));
const idx = (c) => c.toUpperCase().split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
const COL_TOT = Math.max(...mapeo.cc_blocks.map(b => idx(b.col_saldo))) + 1;
const SALDOS = mapeo.cc_blocks.map(b => b.col_saldo);

// Qué filas de la columna de totales NO suman los 22 centros
function incompletas(ws) {
  const malas = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const f = ws.getCell(r, COL_TOT).formula;
    if (!f || /^\s*SUM\(/i.test(String(f))) continue;
    const refs = new Set([...String(f).matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(m => m[1]));
    const faltan = SALDOS.filter(c => !refs.has(c));
    if (faltan.length) malas.push({ fila: r, faltan });
  }
  return malas;
}

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const ws = wb.getWorksheet("Sumas y Saldos");

  // ---------------------------------------------- el archivo tal como está hoy
  const antes = incompletas(ws);
  console.log(`       en el archivo guardado hay ${antes.length} fila(s) que no suman los ${SALDOS.length} centros`);
  for (const m of antes) {
    const cuenta = ws.getCell(m.fila, 1).value;
    const nombres = m.faltan.map(c => (mapeo.cc_blocks.find(b => b.col_saldo === c) || {}).nombre_balance || c);
    console.log(`          fila ${m.fila} (cuenta ${cuenta}) — le falta ${nombres.join(", ")}`);
  }
  check(antes.length > 0, "el control detecta el problema en el archivo real (si no, no probaría nada)");

  // ---------------------------------------------- se repara
  const r = motor.repararTotalesSaldos({ wb, mapeo, log: () => {} });
  check(r.arregladas.length === antes.length,
    `repara exactamente las que estaban mal: ${r.arregladas.length} de ${antes.length}`);
  check(incompletas(ws).length === 0, "después no queda ninguna fila incompleta");

  // ---------------------------------------------- y es idempotente
  const otra = motor.repararTotalesSaldos({ wb, mapeo, log: () => {} });
  check(otra.arregladas.length === 0, "correrla de nuevo no toca nada");

  // ---------------------------------------------- no rompe lo que estaba bien
  // Una fila que ya sumaba los 22 tiene que seguir sumando los 22, ni más ni menos.
  let sanas = 0, duplicadas = 0;
  for (let f = 1; f <= ws.rowCount; f++) {
    const fo = ws.getCell(f, COL_TOT).formula;
    if (!fo || /^\s*SUM\(/i.test(String(fo))) continue;
    sanas++;
    const cols = [...String(fo).matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(m => m[1]);
    if (cols.length !== new Set(cols).size) duplicadas++;
    // toda referencia tiene que ser a la MISMA fila: sumar la fila de al lado sería peor
    const filas = new Set([...String(fo).matchAll(/\$?[A-Z]{1,3}\$?(\d+)/g)].map(m => +m[1]));
    if (filas.size !== 1 || !filas.has(f)) { check(false, `la fila ${f} referencia otras filas: ${fo}`); break; }
  }
  check(duplicadas === 0, `ninguna fórmula suma dos veces el mismo centro (${sanas} filas revisadas)`);

  // ---------------------------------------------- la fila de totales generales no se toca
  const totGen = ws.getCell(mapeo.fila_totales, COL_TOT).formula;
  check(!!totGen && /^\s*SUM\(/i.test(String(totGen)),
    `la fila de totales generales sigue siendo su SUM: ${totGen}`);

  // ---------------------------------------------- y el maestro del disco no se tocó
  const enDisco = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  check(incompletas(enDisco.getWorksheet("Sumas y Saldos")).length === antes.length,
    "el archivo base del disco quedó intacto: la reparación es sobre la copia de la corrida");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
