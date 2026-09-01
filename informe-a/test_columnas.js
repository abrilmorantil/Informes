// Insertar una columna y que las fórmulas queden apuntando bien.
//
// Es la pieza riesgosa del alta de un centro de costo nuevo: si las referencias no se corren,
// las 118 fórmulas de la zona de meses de `Dist.de gastos` quedan una columna a la izquierda y
// cada mes muestra el importe del mes anterior. Peor: el archivo sigue dando números.
//
//   node informe-a/test_columnas.js
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");

global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const { insertColumnEn, icShiftConHoja, icShiftLocal, icIndice, icLetra } =
  require(BASE + "/informe-a/columnas.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };
const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));

// ------------------------------------------------------------------ la aritmética, sin Excel
const Z = icIndice("Z");
check(icShiftLocal("SUM(Z8:AK8)", Z) === "SUM(AA8:AL8)", "SUM(Z8:AK8) -> SUM(AA8:AL8)");
check(icShiftLocal("+Y100", Z) === "+Y100", "una columna ANTES del corte no se mueve: +Y100");
check(icShiftLocal("SUM(F8:Y8)", Z) === "SUM(F8:Y8)", "un rango entero antes del corte queda igual");
check(icShiftLocal("SUM(Y8:AK8)", Z) === "SUM(Y8:AL8)",
  "un rango que CONTIENE el corte se expande: SUM(Y8:AK8) -> SUM(Y8:AL8)");
check(icShiftLocal("$Z$8", Z) === "$AA$8", "se respetan los $: $Z$8 -> $AA$8");
check(icShiftLocal('VLOOKUP("30_"&A7,SyS!$B:$P,10,FALSE())', Z) === 'VLOOKUP("30_"&A7,SyS!$B:$P,10,FALSE())',
  "no se toca el texto entre comillas ni las referencias a otra hoja");
check(icShiftLocal("SUM(AB8:AB99)", Z) === "SUM(AC8:AC99)", "las filas no se tocan, sólo las columnas");
check(icShiftConHoja("+'Dist.de gastos'!Z100", "Dist.de gastos", Z) === "+'Dist.de gastos'!AA100",
  "una referencia con nombre de hoja también corre");
check(icShiftConHoja("+'Dist.de gastos'!O100", "Dist.de gastos", Z) === "+'Dist.de gastos'!O100",
  "y si está antes del corte, no");
check(icShiftConHoja("+'Otra hoja'!Z100", "Dist.de gastos", Z) === "+'Otra hoja'!Z100",
  "las referencias a OTRA hoja no se tocan");

// ------------------------------------------------------------------ sobre el archivo real
(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const dist = wb.getWorksheet("Dist.de gastos");

  const corte = Math.max(...Object.keys(mapeo.dist_col_to_cc).map(icIndice)) + 1;
  console.log(`\n       insertando en ${icLetra(corte)}, la primera columna después de los centros de costo`);

  const T = (c) => {
    const v = c.value;
    if (v == null) return "";
    if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
    return String(v);
  };
  const titulo = (c) => (T(dist.getCell(6, c)) + " " + T(dist.getCell(7, c))).trim();

  const antes = {};
  for (let c = 1; c <= dist.columnCount; c++) {
    antes[c] = { titulo: titulo(c), total: dist.getCell(100, c).formula || null, ancho: dist.getColumn(c).width };
  }
  const columnas = dist.columnCount;

  const r = insertColumnEn(wb, "Dist.de gastos", corte);
  console.log(`       ${r.formulasReacomodadas} fórmula(s) reacomodadas`);

  // 1) lo que estaba a la izquierda del corte no se movió
  let izqOk = true;
  for (let c = 1; c < corte; c++) {
    if (antes[c].titulo !== titulo(c)) izqOk = false;
    if (antes[c].total !== (dist.getCell(100, c).formula || null)) izqOk = false;
  }
  check(izqOk, `las ${corte - 1} columnas anteriores al corte quedaron intactas`);

  // 2) lo del corte en adelante se corrió UNA posición, con su ancho
  let derOk = true, ejemplo = "";
  for (let c = corte; c <= columnas; c++) {
    if (antes[c].titulo !== titulo(c + 1)) {
      derOk = false;
      if (!ejemplo) ejemplo = `${icLetra(c)} decía "${antes[c].titulo}" y en ${icLetra(c + 1)} quedó "${titulo(c + 1)}"`;
    }
    if (antes[c].ancho !== dist.getColumn(c + 1).width) { derOk = false; if (!ejemplo) ejemplo = `el ancho de ${icLetra(c)}`; }
  }
  check(derOk, "todo lo del corte en adelante se corrió una columna, con su ancho" + (ejemplo ? ` — ${ejemplo}` : ""));

  // 3) la columna nueva quedó vacía
  let vacia = true;
  for (let f = 1; f <= dist.rowCount; f++) {
    const v = dist.getCell(f, corte).value;
    if (v !== null && v !== undefined && v !== "") vacia = false;
  }
  check(vacia, `la columna ${icLetra(corte)} quedó vacía y lista para el centro de costo nuevo`);

  // 4) LO QUE IMPORTA: cada TOTAL GASTOS sigue sumando SU columna, no la de al lado
  let bien = 0; const mal = [];
  for (let c = 1; c <= dist.columnCount; c++) {
    const f = dist.getCell(100, c).formula;
    if (!f) continue;
    const cols = new Set([...String(f).matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(m => m[1]));
    if (cols.size === 1 && cols.has(icLetra(c))) bien++;
    else mal.push(`${icLetra(c)}100 = ${f}`);
  }
  check(mal.length === 0, `las ${bien} fórmulas de TOTAL GASTOS siguen sumando su propia columna`);
  mal.slice(0, 5).forEach(x => console.log("        " + x));

  // 5) el TOTAL AÑO sigue abarcando exactamente los 12 meses
  const totalAnio = [];
  for (let c = 1; c <= dist.columnCount; c++) if (/TOTAL\s+A[ÑN]O/i.test(titulo(c))) totalAnio.push(c);
  check(totalAnio.length === 1, `hay una sola columna TOTAL AÑO (la ${icLetra(totalAnio[0])})`);
  const fAnio = dist.getCell(8, totalAnio[0]).formula;
  const m = /SUM\(\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+\)/i.exec(String(fAnio || ""));
  check(!!m, `y su fórmula sigue siendo un SUM: ${fAnio}`);
  if (m) {
    check(icIndice(m[2]) - icIndice(m[1]) + 1 === 12, `que abarca 12 columnas de mes (${m[1]}..${m[2]})`);
    check(icIndice(m[2]) + 1 === totalAnio[0], "y termina justo antes del TOTAL AÑO, sin incluirse a sí mismo");
  }

  // 6) el archivo del disco no se tocó
  const disco = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  check(disco.getWorksheet("Dist.de gastos").columnCount === columnas,
    "el archivo base del disco quedó intacto");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
