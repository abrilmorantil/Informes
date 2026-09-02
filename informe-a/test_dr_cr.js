// Julio 2026 es el mes que distingue lo correcto de lo incorrecto: es el unico con
// HABERES. En junio todas las lineas son al debe, asi que debe, saldo y "movimiento del
// mes" dan lo mismo y cualquier error de signo pasa desapercibido.
//
// La geometria de "Dist.de gastos" es:
//   columnas de centro de costo = el SALDO de cada cuenta en ese centro (debe - haber)
//   columna CR                  = el HABER
//   MOVIMIENTO MES DR = SUM(centros) + CR = (debe - haber) + haber = el DEBE
//   la fila de control del archivo hace DR - CR, que vuelve a dar el SALDO
//
// Este test comprueba las tres identidades contra el export de Onvio. Si alguien vuelve a
// tocar como se engancha la columna CR, aca se ve.
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");
const EXPORT = path.join(BASE, "..", "ejemplos", "SyS_por_Centro_de_costos_mensual_07-26.xls");

global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
const XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
global.XLSX = XLSX;
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
global.insertRowEn = require(BASE + "/informe-c/formula_hojas.js").insertRowEn;
const motor = require(BASE + "/informe-a/motor.js");
const { parseOnvioExport } = require(BASE + "/informe-a/parser_onvio.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };
const form = (c) => { const v = c.value; return (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : null; };
const n2 = (x) => Math.round(x * 100) / 100;

(async () => {
  if (!fs.existsSync(EXPORT)) {
    console.log("  (falta " + EXPORT + ", salteo la prueba)");
    process.exit(0);
  }
  const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));

  const libro = XLSX.read(fs.readFileSync(EXPORT), { type: "buffer", raw: true });
  const filas = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]], { header: 1, raw: true, defval: null });
  const lineas = parseOnvioExport(filas);

  const deResultado = lineas.filter(l => String(l.cuenta_codigo).startsWith("4"));
  const onvio = deResultado.reduce((a, l) => ({
    debe: a.debe + (Number(l.debe) || 0),
    haber: a.haber + (Number(l.haber) || 0),
    saldo: a.saldo + (Number(l.saldo) || 0),
  }), { debe: 0, haber: 0, saldo: 0 });
  check(onvio.haber > 0, `el export de julio tiene haberes (${n2(onvio.haber).toFixed(2)}), que es lo que hace util esta prueba`);

  motor.procesar({ wb, lineas, mapeo, periodo: "2026-07", log: () => {} });

  // reproducir las columnas evaluando las formulas contra los datos del export
  const wsD = wb.getWorksheet("Dist.de gastos");
  const colCr = motor.columnaCrDeDist(wsD);
  check(!!colCr, `encontre la columna CR (${colCr})`);

  const queEs = {};
  for (const b of mapeo.cc_blocks) {
    queEs[b.col_debe] = { cc: b.nombre_balance, tipo: "debe" };
    queEs[b.col_haber] = { cc: b.nombre_balance, tipo: "haber" };
    queEs[b.col_saldo] = { cc: b.nombre_balance, tipo: "saldo" };
  }
  const mov = {};
  for (const l of lineas) {
    const b = motor.resolverCcBlock(mapeo, l.cc_nombre_onvio);
    const cuenta = mapeo.cuentas[l.cuenta_codigo];
    if (!b || !cuenta) continue;
    const k = `${cuenta.ss_row}|${b.nombre_balance}`;
    mov[k] = mov[k] || { debe: 0, haber: 0 };
    mov[k].debe += Number(l.debe) || 0;
    mov[k].haber += Number(l.haber) || 0;
  }
  const valorDe = (colSs, ssRow) => {
    const q = queEs[colSs]; if (!q) return 0;
    const m = mov[`${ssRow}|${q.cc}`]; if (!m) return 0;
    return q.tipo === "debe" ? m.debe : q.tipo === "haber" ? m.haber : (m.debe - m.haber);
  };
  const evaluar = (fila, colDist) => {
    const f = form(wsD.getCell(fila, motor.colAIndice(colDist)));
    if (!f) return 0;
    let t = 0;
    for (const m of f.matchAll(/([+-])?\s*'Sumas y Saldos'!\$?([A-Z]{1,3})\$?(\d+)(?!:)/g)) {
      t += (m[1] === "-" ? -1 : 1) * valorDe(m[2], Number(m[3]));
    }
    return t;
  };

  let saldos = 0, cr = 0;
  for (const cat of mapeo.categorias) {
    if (motor.esFilaDeTotales(cat.desc)) continue;
    for (const col of Object.keys(mapeo.dist_col_to_cc)) saldos += evaluar(cat.dist_row, col);
    cr += evaluar(cat.dist_row, colCr);
  }
  const dr = saldos + cr;

  const igual = (a, b) => Math.abs(a - b) < 0.005;
  check(igual(dr, onvio.debe), `MOVIMIENTO MES DR = el DEBE de Onvio: ${n2(dr).toFixed(2)} vs ${n2(onvio.debe).toFixed(2)}`);
  check(igual(cr, onvio.haber), `la columna CR = el HABER de Onvio: ${n2(cr).toFixed(2)} vs ${n2(onvio.haber).toFixed(2)}`);
  check(igual(dr - cr, onvio.saldo), `DR - CR = el SALDO de Onvio: ${n2(dr - cr).toFixed(2)} vs ${n2(onvio.saldo).toFixed(2)}`);
  check(igual(saldos, onvio.saldo), `las columnas de centro de costo suman el SALDO: ${n2(saldos).toFixed(2)}`);

  // ------------------------------------------------ y la columna del mes tiene que dar eso
  //
  // Es lo que quedaba suelto: las tres identidades de arriba estaban bien, pero la columna
  // "TOTAL <mes>" —la que queda en el informe— seguia a D, o sea al DEBITO, no al saldo. Con
  // julio da 90.260,44 cuando Onvio dice 89.977,67: de mas por los 282,77 de haberes.
  const mesesM = require(BASE + "/informe-a/meses.js");
  const colMes = mesesM.columnaDeMes(wsD, 7);
  check(colMes !== null, "esta la columna TOTAL JULIO en Dist.de gastos");

  mesesM.activarColumnaMes(wsD, 7, () => {});
  const filasCat = mesesM.filasDeCategoria(wsD);
  const formulas = filasCat.map(r => String(wsD.getCell(r, colMes).formula || ""));
  check(formulas.every((f, i) => f === `D${filasCat[i]}-E${filasCat[i]}`),
    `la columna del mes queda en D<fila>-E<fila> (ej: ${formulas[0]})`);

  // Por que eso da el saldo: la propia hoja arma D como SUM(centros)+E, asi que D-E vuelve a
  // ser la suma de los centros, que es lo que arriba ya dio igual al SALDO de Onvio. Con D
  // solo, la columna del mes daba saldo + CR.
  const formD = filasCat.map(r => String(wsD.getCell(r, 4).formula || ""));
  const comoSuma = formD.filter((f, i) => {
    const t = f.replace(/\s+/g, "").replace(/^\+/, "").toUpperCase();
    return t.startsWith(`SUM(F${filasCat[i]}:`) && t.endsWith(`)+E${filasCat[i]}`);
  });
  check(comoSuma.length === formD.length,
    `y D es SUM(centros)+E en las ${formD.length} filas, asi que D-E = la suma de los centros (ej: ${formD[0]})`);
  check(igual(saldos, onvio.saldo) && !igual(saldos + cr, onvio.saldo),
    `la suma de los centros da ${n2(saldos).toFixed(2)} (el saldo) y no ${n2(saldos + cr).toFixed(2)} (el debito), que es lo que daba antes`);

  // el mes sigue reconociendose como vivo con la forma nueva
  const vivos = mesesM.mesesVivos(wsD);
  check(vivos.length === 1 && vivos[0].mes === 7,
    `y el mes sigue reconociendose como abierto: ${JSON.stringify(vivos)}`);

  // la forma vieja tambien, o un archivo ya cargado quedaria con el mes invisible
  for (const r of filasCat) wsD.getCell(r, colMes).value = { formula: `D${r}` };
  const viejos = mesesM.mesesVivos(wsD);
  check(viejos.length === 1 && viejos[0].mes === 7,
    "un archivo con la forma vieja (=D<fila>) se sigue reconociendo abierto");

  // ninguna cuenta puede aportar algo distinto de su saldo a las columnas de centro de costo
  const desviadas = [];
  for (const l of deResultado) {
    const cuenta = mapeo.cuentas[l.cuenta_codigo];
    const b = motor.resolverCcBlock(mapeo, l.cc_nombre_onvio);
    if (!cuenta || !b) { desviadas.push(`${l.cuenta_codigo} no se pudo ubicar`); continue; }
    let aporta = 0;
    for (const cat of mapeo.categorias) {
      if (motor.esFilaDeTotales(cat.desc)) continue;
      for (const col of Object.keys(mapeo.dist_col_to_cc)) {
        const f = form(wsD.getCell(cat.dist_row, motor.colAIndice(col)));
        if (!f) continue;
        for (const m of f.matchAll(/([+-])?\s*'Sumas y Saldos'!\$?([A-Z]{1,3})\$?(\d+)(?!:)/g)) {
          if (Number(m[3]) !== cuenta.ss_row) continue;
          const q = queEs[m[2]];
          if (!q || q.cc !== b.nombre_balance) continue;
          aporta += (m[1] === "-" ? -1 : 1) * valorDe(m[2], cuenta.ss_row);
        }
      }
    }
    const suyo = (Number(l.debe) || 0) - (Number(l.haber) || 0);
    if (Math.abs(aporta - suyo) > 0.005) {
      desviadas.push(`${l.cuenta_codigo} en ${b.nombre_balance}: aporta ${n2(aporta).toFixed(2)} y su saldo es ${n2(suyo).toFixed(2)}`);
    }
  }
  check(desviadas.length === 0,
    `cada una de las ${deResultado.length} lineas aporta exactamente su saldo` +
    (desviadas.length ? "\n         " + desviadas.slice(0, 6).join("\n         ") : ""));

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack || e.message); process.exit(1); });
