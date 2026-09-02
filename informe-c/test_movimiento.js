// La hoja "MOVI <mes>-<año>" del EE RR: el balance de comprobación en dólares.
//
// Es la hoja que Abril venía armando a mano. Una fila por cuenta del balance, con el saldo del
// mes pasado, lo que se movió y cómo queda.
//
// El "Saldo anterior" NO se guarda en ningún lado ni se importa como en el BALCOMPROBDOLARES:
// sale de Hoja1 leída ANTES de que la corrida la pise. El maestro tiene los importes del mes
// pasado hasta el instante en que se los reemplaza.
//
//   node informe-c/test_movimiento.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

const X = require(path.join(AQUI, "..", "informe-a", "vendor", "xlsx.full.min.js"));
global.XLSX = X;
global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const usd = require(path.join(AQUI, "dolares.js"));
const libro = require(path.join(AQUI, "libro_eerr.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const cerca = (a, b) => Math.abs(a - b) < 0.02;
const T = (c) => {
  const v = c.value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return "";
    return String(v.result ?? "");
  }
  return typeof v === "string" ? v : String(v);
};

const CUENTAS = [
  { fila: 3, col: 2, codigo: "111010001", nombre: "Caja", clave: "111010001 - Caja" },
  { fila: 4, col: 2, codigo: "211010001", nombre: "Un Proveedor", clave: "211010001 - Un Proveedor" },
  { fila: 5, col: 2, codigo: "421010000", nombre: "Un Gasto", clave: "421010000 - Un Gasto" },
  { fila: 6, col: 2, codigo: "311000000", nombre: "Capital", clave: "311000000 - Capital" },
];

console.log("=== 1) saldo anterior, movimiento y saldo final ===");
{
  const previos = new Map([
    ["111010001 - CAJA", 1000], ["211010001 - UN PROVEEDOR", -300], ["421010000 - UN GASTO", 50],
  ]);
  const finales = new Map([
    ["111010001 - CAJA", 1200], ["211010001 - UN PROVEEDOR", -500], ["421010000 - UN GASTO", 50],
    ["311000000 - CAPITAL", -750],
  ]);
  const m = usd.movimientoDelMes(CUENTAS, previos, finales);
  check(m.length === CUENTAS.length, `una fila por cuenta del maestro (${m.length})`);
  const caja = m.find(x => x.codigo === "111010001");
  check(cerca(caja.anterior, 1000) && cerca(caja.final, 1200) && cerca(caja.movimiento, 200),
    `Caja: 1000 + 200 = 1200 (dio ${caja.anterior} + ${caja.movimiento} = ${caja.final})`);
  const gasto = m.find(x => x.codigo === "421010000");
  check(cerca(gasto.movimiento, 0), "una cuenta que no se movió queda con movimiento 0, y aparece igual");
  const cap = m.find(x => x.codigo === "311000000");
  check(cerca(cap.anterior, 0) && cerca(cap.movimiento, -750),
    "una cuenta que ANTES no estaba arranca en 0 y todo su saldo es movimiento del mes");
  check(m.map(x => x.capitulo).join(",") === "ACTIVO,PASIVO,RESULTADOS,PATRIMONIO NETO",
    "cada una queda en su capítulo, por el primer dígito del código");
}

console.log("\n=== 2) van TODAS las cuentas, muevan o no ===");
{
  const m = usd.movimientoDelMes(CUENTAS, new Map(), new Map());
  check(m.length === CUENTAS.length,
    "sin ningún saldo, siguen saliendo las 4: es un balance de comprobación, no un listado " +
    "de movimientos");
  check(m.every(x => x.anterior === 0 && x.final === 0 && x.movimiento === 0),
    "todas en cero");
  check(m.every(x => x.enHoja1 === false),
    "y marcadas como que Hoja1 no las tiene, que es lo que hay que mirar");
}

console.log("\n=== 3) la hoja que se arma ===");
{
  const movimiento = usd.movimientoDelMes(CUENTAS,
    new Map([["111010001 - CAJA", 1000]]),
    new Map([["111010001 - CAJA", 1200], ["211010001 - UN PROVEEDOR", -1200]]));
  const wb = libro.construirLibroEERR({
    actual: { ingresosOperacion: 0, gastosOperacion: -1, gastosAdministracion: -1,
              extraordinarios: 0, ajusteTraduccion: 0, otrosIngresos: 0, impuesto: 0,
              antesDeImpuestos: -2, resultadoEjercicio: -2 },
    anterior: null, periodoFin: "2026-08-31", titulo: "X", movimiento,
  });
  check(libro.nombreHojaMovimiento("2026-08-31") === "MOVI 08-2026",
    'la hoja se llama "MOVI 08-2026", como la que se venía armando a mano');
  check(wb.SheetNames.length === 2 && wb.SheetNames[1] === "MOVI 08-2026",
    `el EE RR queda con dos hojas: ${wb.SheetNames.join(" | ")}`);

  const ws = wb.Sheets["MOVI 08-2026"];
  const cel = (a) => (ws[a] ? ws[a].v : undefined);
  check(cel("A1") === "Cuenta Contable" && cel("B1") === "Saldo anterior" &&
        cel("C1") === "Movimiento del Mes Dolares" && cel("D1") === "Saldo Final",
    "con los mismos encabezados que la hoja hecha a mano");
  check(cel("A2") === "ACTIVO", "y agrupada por capítulo");
  check(cel("A3") === "111010001 - Caja" && cel("B3") === 1000 && cel("D3") === 1200,
    "la primera cuenta con sus tres importes");

  // el TOTAL suma sólo las filas de cuenta: un rango se comería los títulos de capítulo
  const filaTotal = Object.keys(ws).filter(k => ws[k] && ws[k].v === "TOTAL")[0];
  check(!!filaTotal, "hay una fila TOTAL");
  const nTot = filaTotal.replace(/[A-Z]/g, "");
  const fB = ws["B" + nTot];
  check(fB && fB.f && !/:/.test(fB.f),
    `el total es una lista de sumandos y no un rango: ${fB && fB.f}`);
  check(fB && cerca(fB.v, 1000), `y su valor es la suma de los saldos anteriores (${fB && fB.v})`);
}

console.log("\n=== 4) sin movimiento, el EE RR sale igual que antes ===");
{
  const wb = libro.construirLibroEERR({
    actual: { ingresosOperacion: 0, gastosOperacion: -1, gastosAdministracion: -1,
              extraordinarios: 0, ajusteTraduccion: 0, otrosIngresos: 0, impuesto: 0,
              antesDeImpuestos: -2, resultadoEjercicio: -2 },
    anterior: null, periodoFin: "2026-08-31", titulo: "X",
  });
  check(wb.SheetNames.length === 1,
    "si no hay movimiento que mostrar, el EE RR queda con su única hoja de siempre");
}

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
