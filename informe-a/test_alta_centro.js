// Las dos salidas cuando aparece un centro de costo que el balance no conoce:
//
//   1) NO es nuevo, está escrito distinto -> se declara la equivalencia y de ahí en más
//      resuelve solo, sin adivinar nada.
//   2) SÍ es nuevo -> se da de alta igual que los demás: su bloque DEBE/HABER/SALDO en
//      "Sumas y Saldos" y su columna en "Dist.de gastos", con el nombre TAL CUAL lo escribe
//      Onvio, y entrando en la columna de TOTALES.
//
// El caso real que lo motivó: el export trae "Proyecto Lonco Vaca- Palenque" y el bloque se
// llama "LONCO VACA - PELENQUE" — se movió el guion y dice PALENQUE con A. Ese nombre ya está
// declarado en `mapeo.json`, así que el test NO lo usa: ver el comentario del bloque 1.
//
//   node informe-a/test_alta_centro.js
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");

global.XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const { insertColumnEn, icIndice, icLetra } = require(BASE + "/informe-a/columnas.js");
global.insertColumnEn = insertColumnEn;
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };
const leerMapeo = () => JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));
const T = (c) => {
  const v = c.value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
};

// ------------------------------------------------------------ 1) el mismo, escrito distinto
{
  const mapeo = leerMapeo();

  // El nombre crudo se ARMA a partir de un bloque real del archivo, en vez de usar uno de
  // verdad. Los nombres que Onvio escribe distinto se van declarando y quedan guardados en
  // `mapeo.json`: un test que use uno real se pone en rojo el día que se lo declara, y ahí
  // pasa a medir el archivo en vez del código. Es lo que pasó con "Proyecto Lonco Vaca-
  // Palenque", que hoy ya está declarado y por eso sí resuelve solo.
  const BLOQUE = mapeo.cc_blocks[0].nombre_balance;
  const CRUDO = `Proyecto ${BLOQUE} (así lo escribiría Onvio)`;
  const clave = CRUDO.replace(/\s+/g, " ").trim().toUpperCase();
  check(!(mapeo.cc_nombres_onvio || {})[clave],
    "punto de partida: ese nombre no está declarado como equivalencia");
  check(motor.resolverCcBlock(mapeo, CRUDO) === null,
    `"${CRUDO}" no resuelve solo — y está bien que no adivine`);

  const r = motor.declararEquivalenciaCc({ mapeo, nombreOnvio: CRUDO, nombreBalance: BLOQUE });
  check(r.nombre_balance === BLOQUE, "se declara contra el bloque que existe");
  const b = motor.resolverCcBlock(mapeo, CRUDO);
  check(!!b && b.nombre_balance === BLOQUE,
    `y ahora sí resuelve: "${CRUDO}" -> "${b && b.nombre_balance}"`);
  check(motor.resolverCcBlock(mapeo, `${CRUDO} II`) === null,
    "y una variante distinta sigue sin adivinarse sola");

  let tiro = null;
  try { motor.declararEquivalenciaCc({ mapeo, nombreOnvio: "X", nombreBalance: "NO EXISTE" }); }
  catch (e) { tiro = e.message; }
  check(!!tiro && /no es ninguno/.test(tiro), "declarar contra un bloque inexistente falla claro");
}

// ------------------------------------------------------------ 2) uno nuevo de verdad
(async () => {
  const mapeo = leerMapeo();
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const wsDist = wb.getWorksheet("Dist.de gastos");

  const NOMBRE = "PROYECTO TOTORAL";     // como lo escribiría Onvio
  const CODIGO = "77";

  const bloquesAntes = mapeo.cc_blocks.length;
  const distAntes = Object.keys(mapeo.dist_col_to_cc).length;
  const colsDistAntes = wsDist.columnCount;
  const totalAnioAntes = (() => {
    for (let c = 1; c <= wsDist.columnCount; c++) {
      if (/TOTAL\s+A[ÑN]O/i.test((T(wsDist.getCell(6, c)) + " " + T(wsDist.getCell(7, c))).trim())) {
        return { col: c, formula: wsDist.getCell(8, c).formula };
      }
    }
    return null;
  })();

  const salida = [];
  const r = motor.agregarCentroDeCosto({
    wb, mapeo, nombreOnvio: NOMBRE, ccCodigo: CODIGO, log: (m) => salida.push(String(m).trim()),
  });
  salida.filter(Boolean).forEach(l => console.log("       " + l));

  // --- queda registrado y resuelve por su nombre de Onvio, sin equivalencia
  check(mapeo.cc_blocks.length === bloquesAntes + 1, "se registró el bloque en el mapeo");
  check(Object.keys(mapeo.dist_col_to_cc).length === distAntes + 1, "y su columna de Dist.de gastos");
  const b = motor.resolverCcBlock(mapeo, NOMBRE);
  check(!!b && b.nombre_balance === NOMBRE,
    `resuelve por el nombre tal cual lo escribe Onvio, sin declarar nada: "${NOMBRE}"`);

  // --- Sumas y Saldos: el bloque, con sus fórmulas apuntando al código nuevo
  const iD = icIndice(r.bloque.col_debe), iH = icIndice(r.bloque.col_haber), iS = icIndice(r.bloque.col_saldo);
  check(T(wsSs.getCell(4, iD)) === NOMBRE, `el encabezado dice "${NOMBRE}" (${r.bloque.col_debe}4)`);
  check(T(wsSs.getCell(5, iS)) === "SALDO", "y la fila 5 rotula DEBE / HABER / SALDO");

  // Las filas de CUENTA y la de SUBTOTAL del bloque son distintas y hay que mirarlas aparte:
  // el modelo tiene en la 299 un SUM de su propia columna, y copiarlo tal cual haría que el
  // bloque nuevo mostrara los importes del bloque modelo con otro nombre.
  let cuentas = 0, conCodigo = 0, saldoOk = 0;
  const subtotales = [];
  for (let f = 1; f <= wsSs.rowCount; f++) {
    const fd = wsSs.getCell(f, iD).formula;
    if (!fd) continue;
    if (/^\s*\+?SUM\(/i.test(String(fd))) { subtotales.push({ fila: f, formula: String(fd) }); continue; }
    cuentas++;
    if (new RegExp(`"${CODIGO}_"`).test(String(fd))) conCodigo++;
    const fs2 = String(wsSs.getCell(f, iS).formula || "");
    if (fs2 === `+${r.bloque.col_debe}${f}-${r.bloque.col_haber}${f}`) saldoOk++;
  }
  check(cuentas === r.filas && cuentas > 200, `${cuentas} filas de cuenta, como el bloque modelo`);
  check(conCodigo === cuentas, `las ${conCodigo} buscan el centro de costo ${CODIGO} en SyS`);
  check(saldoOk === cuentas, "y el SALDO de cada fila es su propio DEBE menos su propio HABER");

  check(subtotales.length === 1, `el bloque tiene su fila de subtotal (la ${subtotales[0] && subtotales[0].fila})`);
  const sub = subtotales[0];
  const colsSub = new Set([...String(sub.formula).matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(x => x[1]));
  check(colsSub.size === 1 && colsSub.has(r.bloque.col_debe),
    `y suma SU PROPIA columna, no la del bloque modelo: ${sub.formula}`);

  // --- entra en la columna de TOTALES
  const iTot = icIndice(r.colTotales);
  const fTot = String(wsSs.getCell(7, iTot).formula || "");
  check(/^\s*\+/.test(fTot) && fTot.includes(`+${r.bloque.col_saldo}7`),
    `la columna de TOTALES (${r.colTotales}) ahora suma también ${r.bloque.col_saldo}`);
  const refs = new Set([...fTot.matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(m => m[1]));
  check(mapeo.cc_blocks.every(x => refs.has(x.col_saldo)),
    `y suma los ${mapeo.cc_blocks.length} centros de costo, sin faltar ninguno`);

  // --- Dist.de gastos: la columna quedó ENTRE los centros y los meses
  const iDist = icIndice(r.distCol);
  check(T(wsDist.getCell(6, iDist)) === NOMBRE, `en Dist.de gastos la columna ${r.distCol} se llama igual`);
  check(wsDist.columnCount === colsDistAntes + 1, "la hoja tiene una columna más");
  const fGastos = String(wsDist.getCell(100, iDist).formula || "");
  check(new RegExp(`SUM\\(${r.distCol}\\d+:${r.distCol}\\d+\\)`).test(fGastos),
    `su TOTAL GASTOS suma su propia columna: ${fGastos}`);

  // --- y los meses siguen enteros: esto es lo que se rompe si el corrimiento falla
  let totalAnioAhora = null;
  for (let c = 1; c <= wsDist.columnCount; c++) {
    if (/TOTAL\s+A[ÑN]O/i.test((T(wsDist.getCell(6, c)) + " " + T(wsDist.getCell(7, c))).trim())) {
      totalAnioAhora = { col: c, formula: wsDist.getCell(8, c).formula };
    }
  }
  check(totalAnioAhora && totalAnioAhora.col === totalAnioAntes.col + 1,
    `TOTAL AÑO se corrió de ${icLetra(totalAnioAntes.col)} a ${icLetra(totalAnioAhora.col)}`);
  const m = /SUM\(\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+\)/i.exec(String(totalAnioAhora.formula));
  check(!!m && icIndice(m[2]) - icIndice(m[1]) + 1 === 12,
    `y sigue abarcando los 12 meses (${m && m[1]}..${m && m[2]})`);
  check(!!m && icIndice(m[2]) + 1 === totalAnioAhora.col, "sin incluirse a sí mismo");

  // --- cada TOTAL GASTOS sigue sumando SU columna
  const mal = [];
  for (let c = 1; c <= wsDist.columnCount; c++) {
    const f = wsDist.getCell(100, c).formula;
    if (!f) continue;
    const cols = new Set([...String(f).matchAll(/\$?([A-Z]{1,3})\$?\d+/g)].map(x => x[1]));
    if (!(cols.size === 1 && cols.has(icLetra(c)))) mal.push(`${icLetra(c)}100 = ${f}`);
  }
  check(mal.length === 0, "ninguna columna quedó sumando la de al lado");
  mal.slice(0, 4).forEach(x => console.log("        " + x));

  // --- Gastos Acumulados: su fila, y DENTRO de los totales
  const gac = wb.getWorksheet("Gastos Acumulados");
  check(typeof r.gastosAcumulados === "number", `se le creó la fila ${r.gastosAcumulados} en Gastos Acumulados`);
  const fg = r.gastosAcumulados;
  check(T(gac.getCell(fg, 1)) === NOMBRE, "con el nombre del proyecto en la columna A");
  check(gac.getCell(fg, 2).value === 0 && gac.getCell(fg, 3).value === 0,
    "el acumulado de años anteriores y el del año arrancan en 0");
  const fD = String(gac.getCell(fg, 4).formula || "");
  check(fD === `+'Dist.de gastos'!${r.distCol}100`,
    `y la columna del mes lee el TOTAL GASTOS de su propia columna: ${fD}`);
  check(String(gac.getCell(fg, 5).formula || "") === `+C${fg}+D${fg}`, "E = C + D");
  check(String(gac.getCell(fg, 6).formula || "") === `+B${fg}+E${fg}`, "F = B + E");

  // Lo que de verdad importa: que el total la incluya. Un rango que termina antes del punto de
  // inserción NO se expande solo, así que la fila quedaría afuera y la hoja mostraría importes
  // correctos que no suman — el mismo error que ya apareció cuatro veces en este archivo.
  let filaTot = null;
  for (let f = 1; f <= gac.rowCount; f++) {
    const ff = gac.getCell(f, 2).formula;
    if (ff && /^\s*\+?SUM\(/i.test(String(ff)) && !T(gac.getCell(f, 1))) { filaTot = f; break; }
  }
  check(filaTot === fg + 1, `la fila de totales quedó justo debajo (${filaTot})`);
  let dentro = 0, fuera = [];
  for (let c = 2; c <= 6; c++) {
    const ff = String(gac.getCell(filaTot, c).formula || "");
    const m = /SUM\(\$?[A-Z]{1,3}\$?(\d+):\$?[A-Z]{1,3}\$?(\d+)\)/i.exec(ff);
    if (!m) continue;
    if (fg >= +m[1] && fg <= +m[2]) dentro++;
    else fuera.push(`${icLetra(c)}${filaTot} = ${ff}`);
  }
  check(fuera.length === 0, `las ${dentro} columnas del total incluyen la fila nueva`);
  fuera.forEach(x => console.log("        " + x));

  // --- no se puede dar de alta dos veces
  let tiro = null;
  try { motor.agregarCentroDeCosto({ wb, mapeo, nombreOnvio: NOMBRE, ccCodigo: "78" }); }
  catch (e) { tiro = e.message; }
  check(!!tiro && /ya existe/.test(tiro), "darlo de alta de nuevo falla en vez de duplicarlo");

  // --- y el código tiene que ser un número
  tiro = null;
  try { motor.agregarCentroDeCosto({ wb, mapeo, nombreOnvio: "OTRO", ccCodigo: "" }); }
  catch (e) { tiro = e.message; }
  check(!!tiro && /numero/i.test(tiro), "sin el código de centro de costo de Onvio, no se da de alta");

  // --- el maestro del disco no se tocó
  const disco = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  check(disco.getWorksheet("Dist.de gastos").columnCount === colsDistAntes,
    "el archivo base del disco quedó intacto");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
