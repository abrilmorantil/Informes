// El arrastre del EEPN: lo que quedó en "Saldos al <fecha>" tiene que pasar a "Saldos al
// inicio" del mes siguiente, y la línea del aumento del mes anterior tiene que salir.
//
//     junio:  inicio 14.700.098.884 + aumento 16/06 148.546.600 = cierre 14.848.645.484
//     julio:  inicio 14.848.645.484 + aumento 24/07 261.309.720 = cierre 15.109.955.204
//
// Se corre sobre el maestro real, en memoria: no lo toca.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const {
  ubicarPatNeto, capitalDeclarado, pasarCierreAlInicio, agregarAumentoDeCapital,
} = require(path.join(AQUI, "capital.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const f = (x) => (typeof x === "number"
  ? x.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  : String(x));
const cerca = (a, b) => Math.abs(a - b) < 0.005;
const texto = (ws, r, c) => {
  const v = ws.getCell(r, c).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
};

const APERTURA_JUNIO = 14700098884;
const AUMENTO_JUNIO = 148546600;
const CIERRE_JUNIO = 14848645484;
const AUMENTO_JULIO = 261309720;
const CIERRE_JULIO = 15109955204;

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const ws = wb.getWorksheet("Pat.Neto");
  const u0 = ubicarPatNeto(ws);
  check(!!u0, "se ubica la estructura de Pat.Neto");
  console.log(`       apertura ${u0.filaInicio}, aumento ${u0.filaAumento}, cierre ${u0.filaCierre}, ` +
              `capital en la columna ${ws.getColumn(u0.colCapital).letter}`);

  const dec0 = capitalDeclarado(ws, u0);
  check(cerca(dec0.total, CIERRE_JUNIO), `el maestro declara el cierre de junio: ${f(dec0.total)}`);
  check(dec0.filas.length === 2 && cerca(dec0.filas[0].importe, APERTURA_JUNIO) &&
        cerca(dec0.filas[1].importe, AUMENTO_JUNIO),
        `y lo hace en dos líneas: ${dec0.filas.map(x => x.etiqueta + " " + f(x.importe)).join(" | ")}`);

  // --------------------------------------------------- se cierra julio: el arrastre
  const r1 = pasarCierreAlInicio(wb, { anio: 2026, mes: 7, dia: 31 });
  check(r1.limpiadas.length === 1 && /16\/06\/2026/.test(r1.limpiadas[0].etiqueta),
        `se limpió la línea de junio: "${r1.limpiadas[0] && r1.limpiadas[0].etiqueta}"`);
  check(cerca(r1.apertura, CIERRE_JUNIO),
        `el capital declarado NO cambió con el arrastre: ${f(r1.apertura)}`);

  const u1 = ubicarPatNeto(ws);
  check(!!u1, "Pat.Neto se sigue pudiendo ubicar aunque no quede ningún aumento cargado");
  check(u1 && u1.colCapital === u0.colCapital && u1.filaInicio === u0.filaInicio &&
        u1.filaCierre === u0.filaCierre, "y da la misma geometría que antes");

  const cap = ws.getCell(u0.filaInicio, u0.colCapital).value;
  check(cerca(cap, CIERRE_JUNIO), `la apertura de julio quedó en el cierre de junio: ${f(cap)}`);
  check(texto(ws, r1.limpiadas[0].fila, 2) === "", "la fila de junio quedó sin etiqueta");
  check(ws.getCell(r1.limpiadas[0].fila, u0.colCapital).value == null,
        "y sin importe");
  const modelo = ws.getCell(r1.limpiadas[0].fila, 6);
  check(!!modelo.formula, `pero con el cableado puesto: F${r1.limpiadas[0].fila} = ${modelo.formula}`);

  const dec1 = capitalDeclarado(ws, u1);
  check(dec1.filas.length === 1, "queda una sola línea de capital, la apertura");
  check(cerca(dec1.total, CIERRE_JUNIO), `que suma lo mismo: ${f(dec1.total)}`);

  // --------------------------------------------------- correrlo dos veces no rompe nada
  const r2 = pasarCierreAlInicio(wb, { anio: 2026, mes: 7, dia: 31 });
  check(r2.limpiadas.length === 0 && cerca(
        ws.getCell(u0.filaInicio, u0.colCapital).value, CIERRE_JUNIO),
        "correr el arrastre de nuevo no vuelve a mover nada");

  // --------------------------------------------------- entra el aumento de julio
  const ag = agregarAumentoDeCapital(wb, {
    etiqueta: "Aumento de capital 24/07/2026", importe: AUMENTO_JULIO,
  });
  check(cerca(ag.importe, AUMENTO_JULIO), `se carga el aumento de julio en la fila ${ag.fila}`);
  const dec2 = capitalDeclarado(ws, ubicarPatNeto(ws));
  check(cerca(dec2.total, CIERRE_JULIO), `y el capital declarado pasa a ${f(dec2.total)}`);
  console.log("       " + dec2.filas.map(x => `${x.etiqueta || "(apertura)"} ${f(x.importe)}`).join("  +  "));

  // --------------------------------------------------- lo del mes en curso NO se arrastra
  const r3 = pasarCierreAlInicio(wb, { anio: 2026, mes: 7, dia: 31 });
  check(r3.limpiadas.length === 0 && r3.sonDelMes.length === 1,
        "volver a cerrar julio deja el aumento de julio donde está");
  check(cerca(capitalDeclarado(ws, ubicarPatNeto(ws)).total, CIERRE_JULIO),
        "y el capital sigue en el cierre de julio");

  // --------------------------------------------------- ahora se cierra agosto
  const r4 = pasarCierreAlInicio(wb, { anio: 2026, mes: 8, dia: 31 });
  check(r4.limpiadas.length === 1 && /24\/07\/2026/.test(r4.limpiadas[0].etiqueta),
        "al cerrar agosto sí se arrastra el aumento de julio");
  check(cerca(ws.getCell(u0.filaInicio, u0.colCapital).value, CIERRE_JULIO),
        `y la apertura de agosto queda en ${f(ws.getCell(u0.filaInicio, u0.colCapital).value)}`);

  // --------------------------------------------------- lo que NO se toca
  const j12 = wb.getWorksheet("Pat.Neto").getCell(u0.filaInicio, 10);
  check(!!j12.formula && /SALDOS/i.test(j12.formula),
        `"Resultados no asignados" al inicio sigue siendo la fórmula de siempre: =${j12.formula}`);

  // ----------------------------------------------------------------- y el maestro de dólares
  // Acá la trampa es otra: "Capital suscripto" (E12) está escrito a mano, no es una fórmula
  // como en pesos, así que también tiene que arrastrarse o el mes siguiente abre por menos de
  // lo que cerró.
  console.log("\n--- dólares ---");
  const APERTURA_USD = 51613898.9, CIERRE_USD_JUNIO = 51718582.9;
  const AUMENTO_USD_JULIO = 176680, CIERRE_USD_JULIO = 51895262.9;
  const COL_SUSCRIPTO = 5;                                  // E, "Capital suscripto"
  const COL_RNA = 9;                                        // I, "Resultados no asignados"

  const wbu = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_dolares.xlsx")));
  const wsu = wbu.getWorksheet("Pat.Neto");
  const uu = ubicarPatNeto(wsu);
  const valu = (c) => wsu.getCell(uu.filaInicio, c).value;
  check(cerca(valu(uu.colCapital), APERTURA_USD) && cerca(valu(COL_SUSCRIPTO), APERTURA_USD),
        "en dólares la apertura trae el capital escrito a mano en DOS columnas");

  const ru = pasarCierreAlInicio(wbu, { anio: 2026, mes: 7, dia: 31 });
  check(cerca(valu(uu.colCapital), CIERRE_USD_JUNIO),
        `capital integrado arrastrado: ${f(valu(uu.colCapital))}`);
  check(cerca(valu(COL_SUSCRIPTO), CIERRE_USD_JUNIO),
        `capital suscripto también, que es el que se quedaba atrás: ${f(valu(COL_SUSCRIPTO))}`);
  check(cerca(valu(COL_RNA), -48116541.37),
        `y "Resultados no asignados" quedó donde estaba: ${f(valu(COL_RNA))}`);

  const agu = agregarAumentoDeCapital(wbu, {
    etiqueta: "Aumento de capital 24/07/2026", importe: AUMENTO_USD_JULIO,
  });
  check(agu.fila === ru.limpiadas[0].fila,
        `el aumento nuevo entra en la fila que dejó libre el anterior (${agu.fila})`);
  check(cerca(capitalDeclarado(wsu, ubicarPatNeto(wsu)).total, CIERRE_USD_JULIO),
        `y el capital en dólares queda en ${f(capitalDeclarado(wsu, ubicarPatNeto(wsu)).total)}`);

  const ru2 = pasarCierreAlInicio(wbu, { anio: 2026, mes: 8, dia: 31 });
  check(ru2.limpiadas.length === 1 && cerca(valu(uu.colCapital), CIERRE_USD_JULIO) &&
        cerca(valu(COL_SUSCRIPTO), CIERRE_USD_JULIO),
        `y agosto abre con las dos columnas en ${f(valu(COL_SUSCRIPTO))}`);

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
