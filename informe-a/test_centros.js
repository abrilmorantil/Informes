// Prueba de los centros de costo: que las columnas nuevas reciban plata de verdad,
// que las columnas se muestren y se oculten segun el movimiento del mes, y que los
// dos nombres de un mismo proyecto (SOL / PROYECTO SOL) vayan a la misma columna.
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");

global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
global.XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
global.insertRowEn = require(BASE + "/informe-c/formula_hojas.js").insertRowEn;
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };
const form = (c) => { const v = c.value; return (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : null; };

(async () => {
  const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const wsDist = wb.getWorksheet("Dist.de gastos");

  // --- todos los centros de costo tienen a donde ir
  const sinSalida = mapeo.cc_blocks
    .map(b => b.nombre_balance)
    .filter(n => {
      const destino = (mapeo.cc_alias || {})[n] || n;
      return !Object.values(mapeo.dist_col_to_cc).some(i => i.nombre_balance === destino);
    });
  check(sinSalida.length === 0,
    `los ${mapeo.cc_blocks.length} centros de costo tienen columna en Dist.de gastos${sinSalida.length ? " — faltan " + sinSalida.join(", ") : ""}`);

  // --- los dos nombres del mismo proyecto caen en la misma columna
  for (const [huerfano, gemelo] of Object.entries(mapeo.cc_alias || {})) {
    const b = motor.resolverCcBlock(mapeo, huerfano);
    check(b && b.nombre_balance === gemelo, `"${huerfano}" va a ${b ? b.nombre_balance : "ningun lado"} (esperado ${gemelo})`);
  }

  // --- el centro de costo se resuelve por texto exacto, NUNCA por parecido.
  // Antes el motor se quedaba con el nombre mas parecido que encontrara, asi que un centro
  // de costo nuevo entraba callado en la columna de otro proyecto. Estos nombres se parecen
  // a uno real y no existen: ninguno tiene que resolver.
  const TRAMPAS = ["Tanque Blanco", "Cerro Amarillo", "Las Termas", "Sheyla II",
                   "Oficina Buenos Aires", "La Voluntad", "Esperanzas", "Gastos Generale"];
  const adivinados = TRAMPAS
    .map(t => ({ t, b: motor.resolverCcBlock(mapeo, t) }))
    .filter(x => x.b)
    .map(x => `"${x.t}" -> ${x.b.nombre_balance}`);
  check(adivinados.length === 0,
    `no adivina ningun centro de costo parecido${adivinados.length ? " — adivina " + adivinados.join(", ") : ` (probados ${TRAMPAS.length})`}`);

  // y las equivalencias declaradas si tienen que resolver
  for (const [texto, destino] of Object.entries(mapeo.cc_nombres_onvio || {})) {
    const b = motor.resolverCcBlock(mapeo, texto);
    const esperado = (mapeo.cc_alias || {})[destino] || destino;
    check(!!b && b.nombre_balance === esperado,
      `la equivalencia declarada "${texto}" va a ${b ? b.nombre_balance : "ningun lado"} (esperado ${esperado})`);
  }

  // --- una cuenta con movimiento en un centro de costo nuevo llega a su columna
  const CC = "SAMENTA";
  const colSamenta = Object.entries(mapeo.dist_col_to_cc).find(([, i]) => i.nombre_balance === CC);
  check(!!colSamenta, `${CC} tiene columna (${colSamenta ? colSamenta[0] : "ninguna"})`);
  const codigo = Object.keys(mapeo.cuentas)[0];
  const cuenta = mapeo.cuentas[codigo];

  const lineas = [{
    cuenta_codigo: codigo, cuenta_label: cuenta.label, cc_nombre_onvio: "Samenta",
    debe: 1234.56, haber: 0, saldo: 1234.56,
  }];
  motor.procesar({ wb, lineas, mapeo, periodo: "2026-08", log: () => {} });

  const distRow = (mapeo.categorias.find(c => c.desc === mapeo.cuentas[codigo].categoria) || {}).dist_row;
  const f = distRow ? form(wsDist.getCell(distRow, motor.colAIndice(colSamenta[0]))) : null;
  const bloque = mapeo.cc_blocks.find(b => b.nombre_balance === CC);
  check(!!f && f.includes(`'Sumas y Saldos'!${bloque.col_saldo}${cuenta.ss_row}`),
    `el importe de ${CC} llega a la columna ${colSamenta[0]}: ${f ? f.slice(0, 70) : "la celda quedo vacia"}`);

  // --- las columnas se muestran y se ocultan solas
  check(wsDist.getColumn(motor.colAIndice(colSamenta[0])).hidden === false,
    `la columna de ${CC} se mostro sola porque este mes tuvo movimiento`);
  const otras = Object.entries(mapeo.dist_col_to_cc)
    .filter(([, i]) => i.nombre_balance !== CC)
    .filter(([col]) => !wsDist.getColumn(motor.colAIndice(col)).hidden);
  check(otras.length === 0,
    `los otros ${Object.keys(mapeo.dist_col_to_cc).length - 1} centros de costo quedaron ocultos${otras.length ? " — quedaron a la vista " + otras.map(o => o[0]).join(", ") : ""}`);

  // --- cada fila de "Gastos Acumulados" lee el TOTAL GASTOS DE SU PROPIO centro de costo.
  // La columna del mes estaba arrastrada con el mouse: siete filas leian celdas vacias del
  // bloque de control que hay debajo de TOTAL GASTOS (O108, O118, O120...). Daban 0 siempre,
  // asi que el gasto de esos proyectos nunca iba a aparecer en su acumulado.
  // Hay que mirar la FILA ademas de la columna: solo la de TOTAL GASTOS trae el gasto del mes.
  const wsGA = wb.getWorksheet("Gastos Acumulados");
  const filaTG = motor.filaTotalGastosDeDist(wsDist);
  const colDeCc = {};
  for (const [c, i] of Object.entries(mapeo.dist_col_to_cc)) colDeCc[i.nombre_balance] = c;

  const desalineadas = [], usoDeColumna = {};
  for (let r = 1; r <= wsGA.rowCount; r++) {
    const nombre = String(wsGA.getCell(r, 1).value || "").trim();
    const f = form(wsGA.getCell(r, 4));
    if (!nombre || !f) continue;
    const m = f.match(/'Dist\.de gastos'!\$?([A-Z]{1,3})\$?(\d+)/);
    const b = motor.resolverCcBlock(mapeo, nombre);
    if (!b) continue;                                  // proyecto dado de baja: lee 0 a propósito
    const colOk = colDeCc[b.nombre_balance];
    if (!m || m[1] !== colOk || Number(m[2]) !== filaTG) {
      desalineadas.push(`fila ${r} "${nombre}" lee ${m ? m[1] + m[2] : "otra cosa"} y deberia leer ${colOk}${filaTG}`);
    } else {
      (usoDeColumna[m[1]] = usoDeColumna[m[1]] || []).push(nombre);
    }
  }
  const SANGRIA = "\n         ";
  check(desalineadas.length === 0,
    `cada fila de Gastos Acumulados lee el TOTAL GASTOS de su centro de costo` +
    (desalineadas.length ? SANGRIA + desalineadas.join(SANGRIA) : ""));
  const dosVeces = Object.entries(usoDeColumna).filter(([, v]) => v.length > 1);
  check(dosVeces.length === 0,
    `ninguna columna alimenta a dos proyectos${dosVeces.length ? " — " + dosVeces.map(([c, v]) => c + ": " + v.join(" y ")).join(" | ") : ""}`);

  // --- la fila de titulos de SyS sobrevive la corrida
  const wsSys = wb.getWorksheet("SyS");
  const rotulos = [11, 13, 16].map(c => {
    const v = wsSys.getCell(1, c).value;
    return typeof v === "object" && v !== null ? String(v.result ?? "") : String(v ?? "");
  });
  check(rotulos.every(t => /u\$s/i.test(t)),
    `la fila de titulos de SyS sigue entera: ${rotulos.join(" | ")}`);

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack || e.message); process.exit(1); });
