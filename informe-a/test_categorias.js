// ¿La cirugía de fórmulas de la categorización deja el archivo correcto?
//
// Se corre sobre base_actual.xlsx, el archivo real, SIN guardarlo: el workbook se
// modifica en memoria y se compara contra sí mismo antes y después.
//
//   node informe-a/test_categorias.js
//
// Lo que se controla, en este orden:
//   1. Que mover una cuenta no cambie NINGÚN total. Cada cuenta suma una vez por
//      columna, así que mover referencias de fila es una reorganización pura: el
//      conjunto de referencias por columna tiene que quedar idéntico.
//   2. Que la cuenta quede en UNA sola categoría, la elegida.
//   3. Que quitarla la saque de todas.
//   4. Que renombrarla no toque ninguna fórmula.
const BASE = require("path").join(__dirname, "..");
const fs = require("fs");
const ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
global.ExcelJS = ExcelJS;
require(BASE + "/informe-a/formula_utils.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

const ARCHIVO = BASE + "/informe-a/base_actual.xlsx";
const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));

// Todas las referencias del tablero, como un conjunto "columna!filaSyS" -> veces.
// Si un total cambiara, este conteo cambiaría.
function censo(wsDist, mapeo) {
  const cuenta = {};
  const filas = mapeo.categorias.map(c => c.dist_row);
  const cols = Object.keys(mapeo.dist_col_to_cc).concat([motor.columnaCrDeDist(wsDist)]).filter(Boolean);
  for (const r of filas) {
    for (const c of cols) {
      const v = wsDist.getCell(`${c}${r}`).value;
      const f = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";
      const re = /'Sumas y Saldos'!\$?([A-Z]{1,3})\$?(\d+)/g;
      let m;
      while ((m = re.exec(f)) !== null) {
        const k = `${c}|${m[1]}${m[2]}`;      // columna de Dist + celda de SyS
        cuenta[k] = (cuenta[k] || 0) + 1;
      }
    }
  }
  return cuenta;
}

function categoriasDe(wsDist, mapeo, ssRow) {
  const refs = motor.refsDeCuentaEnDist(wsDist, mapeo, ssRow, mapeo.categorias.map(c => c.dist_row));
  return [...new Set(refs.map(r => r.distRow))].sort((a, b) => a - b);
}

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(ARCHIVO));
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const wsSs = wb.getWorksheet("Sumas y Saldos");

  const antes = censo(wsDist, mapeo);
  console.log(`referencias en el tablero: ${Object.values(antes).reduce((a, b) => a + b, 0)}`);

  // Ninguna celda de 'Sumas y Saldos' puede estar sumada dos veces en la misma columna
  // del tablero: seria el mismo importe contado dos veces. Antes habia 29 en la columna
  // CR —haberes que habian quedado colgados de la categoria de al lado— y se limpiaron
  // al completar la columna.
  const repetidas = Object.entries(antes).filter(([, v]) => v > 1).map(([k]) => k);
  check(repetidas.length === 0,
        `ninguna referencia esta repetida en el tablero (hay ${repetidas.length}: ${repetidas.slice(0, 8).join(", ")})`);

  // --- 1. mover una cuenta a otra categoría ----------------------------------
  // El test no puede dar por sentado en cuántas categorías arranca cada cuenta: eso
  // depende de cómo esté categorizado el archivo en ese momento. Lo que se prueba es la
  // operación — mover deja la cuenta en UNA sola categoría, la elegida, sin tocar totales.
  const COD = "433070000";                       // Gst de Vehículos Gst de Campo USD
  const cta = mapeo.cuentas[COD];
  const catsAntes = categoriasDe(wsDist, mapeo, cta.ss_row);
  console.log(`\n${COD} ${cta.label} — fila ${cta.ss_row}, hoy en la(s) fila(s) ${catsAntes.join(", ")}`);
  check(catsAntes.length >= 1, `hoy se distribuye en alguna categoría (${catsAntes.length})`);

  // se elige a propósito una categoría DISTINTA de donde está, para que haya algo que mover
  const otra = mapeo.categorias.find(c => !catsAntes.includes(c.dist_row) && c.desc !== "TOTAL GASTOS");
  const DESTINO = otra.desc;
  const r = motor.moverCuentaDeCategoria({ wb, mapeo, codigo: COD, categoriaDestino: otra.dist_row, log: () => {} });
  console.log(`  movidas ${r.movidas}, ya estaban ${r.yaEstaban}, destino "${r.categoria}" (fila ${otra.dist_row})`);
  check(r.movidas > 0, `movió referencias (${r.movidas})`);

  const catsDespues = categoriasDe(wsDist, mapeo, cta.ss_row);
  check(catsDespues.length === 1, `queda en UNA sola categoría (quedó en ${catsDespues.length})`);
  check(catsDespues[0] === otra.dist_row, `y es la elegida (fila ${catsDespues[0]}, esperaba ${otra.dist_row})`);
  check(mapeo.cuentas[COD].categoria === DESTINO, "el mapeo quedó con la categoría nueva");

  // El control que de verdad importa: el censo cuenta CUÁNTAS veces suma cada celda de
  // 'Sumas y Saldos' en cada columna del tablero. Mover una cuenta de fila no puede
  // cambiar ese conteo — si cambiara, algún total cambió.
  const despues = censo(wsDist, mapeo);
  const dif = [];
  for (const k of new Set([...Object.keys(antes), ...Object.keys(despues)])) {
    if ((antes[k] || 0) !== (despues[k] || 0)) dif.push(`${k}: ${antes[k] || 0} -> ${despues[k] || 0}`);
  }
  check(dif.length === 0, `NINGÚN total cambió: cada referencia suma las mismas veces que antes${dif.length ? " — " + dif.slice(0, 6).join(" · ") : ""}`);

  const ssRowsDestino = mapeo.categorias.find(c => c.desc === DESTINO).ss_rows;
  check(ssRowsDestino.includes(cta.ss_row), "la categoría destino tiene la fila en su lista ss_rows");
  const otras = mapeo.categorias.filter(c => c.desc !== DESTINO && c.ss_rows.includes(cta.ss_row));
  check(otras.length === 0, `ninguna otra categoría la sigue listando (${otras.length})`);

  // --- 2. quitarla de la distribución ---------------------------------------
  const COD2 = "426270000";                      // GST VARIOS CAMPO GST. CAMPO, 4 categorías
  const cta2 = mapeo.cuentas[COD2];
  const q = motor.quitarCuentaDeDistribucion({ wb, mapeo, codigo: COD2, log: () => {} });
  console.log(`\n${COD2} ${cta2.label} — se quitaron ${q.quitadas} referencias`);
  check(categoriasDe(wsDist, mapeo, cta2.ss_row).length === 0, "ya no está en ninguna categoría");
  check(mapeo.cuentas[COD2] !== undefined, "PERO sigue en el mapeo (si no, la próxima corrida la duplicaría)");
  check(mapeo.cuentas[COD2].excluida === true, "queda marcada como excluida, así no se vuelve a preguntar");

  // --- 3. renombrar ----------------------------------------------------------
  const COD3 = "424250000";
  console.log(`\n${COD3}: la fila ${COD3 in mapeo.cuentas ? mapeo.cuentas[COD3].ss_row : "?"} ` +
              `${COD3 in mapeo.cuentas ? "está" : "NO está"} en el mapeo`);
  const codRenombrar = COD3 in mapeo.cuentas ? COD3 : "421050000";
  const filaR = mapeo.cuentas[codRenombrar].ss_row;
  const formulaAntes = JSON.stringify(wsSs.getCell(`C${filaR}`).value);
  motor.renombrarCuenta({ wb, mapeo, codigo: codRenombrar, nombre: "PRUEBA DE NOMBRE", log: () => {} });
  check(wsSs.getCell(`B${filaR}`).value === "PRUEBA DE NOMBRE", "el nombre se escribió en la columna B");
  check(JSON.stringify(wsSs.getCell(`C${filaR}`).value) === formulaAntes,
        "y la fórmula VLOOKUP de la fila quedó intacta");

  let tiro = false;
  try { motor.renombrarCuenta({ wb, mapeo, codigo: codRenombrar, nombre: "   " }); } catch (e) { tiro = true; }
  check(tiro, "no deja dejar el nombre vacío");

  console.log(`\n${fallos ? `${fallos} FALLA(S)` : "todo OK"}`);
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.message); process.exit(1); });
