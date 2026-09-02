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

const { fpCierreDeCelda, fpMesSiguiente } = require(path.join(AQUI, "..", "fechas.js"));

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

// NADA de esto se clava a los importes ni al mes del maestro: cada cierre los mueve, y un
// test clavado a "el cierre de junio" se pone en rojo todos los meses sin que se haya roto
// nada — pasó al cerrar julio. Lo que se prueba son RELACIONES, que valen cualquier mes: el
// arrastre no cambia el capital declarado, pliega la línea del mes anterior en la apertura, y
// deja libre esa fila para el aumento nuevo.
const AUMENTO_PRUEBA = 261309720;          // un importe cualquiera, lo pone el test

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const ws = wb.getWorksheet("Pat.Neto");
  const u0 = ubicarPatNeto(ws);
  check(!!u0, "se ubica la estructura de Pat.Neto");
  console.log(`       apertura ${u0.filaInicio}, aumento ${u0.filaAumento}, cierre ${u0.filaCierre}, ` +
              `capital en la columna ${ws.getColumn(u0.colCapital).letter}`);

  // el mes que sigue al que el archivo dice tener cerrado
  const cierreArchivo = fpCierreDeCelda(wb.getWorksheet("Balance").getCell("F5"));
  check(!!cierreArchivo, `el maestro dice a qué mes está cerrado: ${JSON.stringify(cierreArchivo)}`);
  const SIG = fpMesSiguiente(cierreArchivo);

  const dec0 = capitalDeclarado(ws, u0);
  const CIERRE_BASE = dec0.total;
  const conAumento = dec0.filas.filter(x => /aumento/i.test(x.etiqueta || ""));
  check(dec0.filas.length >= 1 && CIERRE_BASE > 0,
        `el maestro declara su capital en ${dec0.filas.length} línea(s): ${f(CIERRE_BASE)}`);
  check(cerca(dec0.filas.reduce((a, x) => a + x.importe, 0), CIERRE_BASE),
        `y las líneas suman ese total: ${dec0.filas.map(x => (x.etiqueta || "(apertura)") + " " + f(x.importe)).join(" | ")}`);

  // --------------------------------------------------- se cierra julio: el arrastre
  const etiquetasAntes = conAumento.map(x => x.etiqueta);
  const r1 = pasarCierreAlInicio(wb, SIG);
  check(r1.limpiadas.length === conAumento.length &&
        r1.limpiadas.every(x => etiquetasAntes.includes(x.etiqueta)),
        `se limpió la línea del mes anterior: ${r1.limpiadas.map(x => '"' + x.etiqueta + '"').join(", ") || "(ninguna)"}`);
  check(cerca(r1.apertura, CIERRE_BASE),
        `el capital declarado NO cambió con el arrastre: ${f(r1.apertura)}`);

  const u1 = ubicarPatNeto(ws);
  check(!!u1, "Pat.Neto se sigue pudiendo ubicar aunque no quede ningún aumento cargado");
  check(u1 && u1.colCapital === u0.colCapital && u1.filaInicio === u0.filaInicio &&
        u1.filaCierre === u0.filaCierre, "y da la misma geometría que antes");

  const cap = ws.getCell(u0.filaInicio, u0.colCapital).value;
  check(cerca(cap, CIERRE_BASE), `la apertura del mes nuevo quedó en el cierre anterior: ${f(cap)}`);
  check(texto(ws, r1.limpiadas[0].fila, 2) === "", "la fila de junio quedó sin etiqueta");
  check(ws.getCell(r1.limpiadas[0].fila, u0.colCapital).value == null,
        "y sin importe");
  const modelo = ws.getCell(r1.limpiadas[0].fila, 6);
  check(!!modelo.formula, `pero con el cableado puesto: F${r1.limpiadas[0].fila} = ${modelo.formula}`);

  const dec1 = capitalDeclarado(ws, u1);
  check(dec1.filas.length === 1, "queda una sola línea de capital, la apertura");
  check(cerca(dec1.total, CIERRE_BASE), `que suma lo mismo: ${f(dec1.total)}`);

  // --------------------------------------------------- correrlo dos veces no rompe nada
  const r2 = pasarCierreAlInicio(wb, SIG);
  check(r2.limpiadas.length === 0 && cerca(
        ws.getCell(u0.filaInicio, u0.colCapital).value, CIERRE_BASE),
        "correr el arrastre de nuevo no vuelve a mover nada");

  // --------------------------------------------------- entra el aumento de julio
  const ETIQUETA = `Aumento de capital ${SIG.dia}/${String(SIG.mes).padStart(2, "0")}/${SIG.anio}`;
  const CIERRE_NUEVO = CIERRE_BASE + AUMENTO_PRUEBA;
  const ag = agregarAumentoDeCapital(wb, { etiqueta: ETIQUETA, importe: AUMENTO_PRUEBA });
  check(cerca(ag.importe, AUMENTO_PRUEBA), `se carga el aumento del mes en la fila ${ag.fila}`);
  const dec2 = capitalDeclarado(ws, ubicarPatNeto(ws));
  check(cerca(dec2.total, CIERRE_NUEVO), `y el capital declarado pasa a ${f(dec2.total)}`);
  console.log("       " + dec2.filas.map(x => `${x.etiqueta || "(apertura)"} ${f(x.importe)}`).join("  +  "));

  // --------------------------------------------------- lo del mes en curso NO se arrastra
  const r3 = pasarCierreAlInicio(wb, SIG);
  check(r3.limpiadas.length === 0 && r3.sonDelMes.length === 1,
        "volver a cerrar el mismo mes deja su aumento donde está");
  check(cerca(capitalDeclarado(ws, ubicarPatNeto(ws)).total, CIERRE_NUEVO),
        "y el capital sigue igual");

  // --------------------------------------------------- ahora se cierra el mes de después
  const r4 = pasarCierreAlInicio(wb, fpMesSiguiente(SIG));
  check(r4.limpiadas.length === 1 && r4.limpiadas[0].etiqueta === ETIQUETA,
        `al cerrar el mes siguiente sí se arrastra ese aumento: "${r4.limpiadas[0] && r4.limpiadas[0].etiqueta}"`);
  check(cerca(ws.getCell(u0.filaInicio, u0.colCapital).value, CIERRE_NUEVO),
        `y la apertura queda en ${f(ws.getCell(u0.filaInicio, u0.colCapital).value)}`);

  // --------------------------------------------------- lo que NO se toca
  const j12 = wb.getWorksheet("Pat.Neto").getCell(u0.filaInicio, 10);
  check(!!j12.formula && /SALDOS/i.test(j12.formula),
        `"Resultados no asignados" al inicio sigue siendo la fórmula de siempre: =${j12.formula}`);

  // ----------------------------------------------------------------- y el maestro de dólares
  // Acá la trampa es otra: "Capital suscripto" (E12) está escrito a mano, no es una fórmula
  // como en pesos, así que también tiene que arrastrarse o el mes siguiente abre por menos de
  // lo que cerró.
  console.log("\n--- dólares ---");
  const AUMENTO_USD = 176680;               // un importe cualquiera, lo pone el test
  const COL_SUSCRIPTO = 5;                                  // E, "Capital suscripto"
  const COL_RNA = 9;                                        // I, "Resultados no asignados"

  const wbu = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_dolares.xlsx")));
  const wsu = wbu.getWorksheet("Pat.Neto");
  const uu = ubicarPatNeto(wsu);
  const valu = (c) => wsu.getCell(uu.filaInicio, c).value;
  const SIGU = fpMesSiguiente(fpCierreDeCelda(wbu.getWorksheet("Balance").getCell("F5")));
  check(cerca(valu(uu.colCapital), valu(COL_SUSCRIPTO)),
        `en dólares la apertura trae el capital escrito a mano en DOS columnas: ${f(valu(COL_SUSCRIPTO))}`);

  const DECU = capitalDeclarado(wsu, uu).total;
  const RNA_ANTES = valu(COL_RNA);
  const ru = pasarCierreAlInicio(wbu, SIGU);
  check(cerca(valu(uu.colCapital), DECU),
        `capital integrado arrastrado: ${f(valu(uu.colCapital))}`);
  check(cerca(valu(COL_SUSCRIPTO), DECU),
        `capital suscripto también, que es el que se quedaba atrás: ${f(valu(COL_SUSCRIPTO))}`);
  check(cerca(valu(COL_RNA), RNA_ANTES),
        `y "Resultados no asignados" quedó donde estaba: ${f(valu(COL_RNA))}`);

  const agu = agregarAumentoDeCapital(wbu, {
    etiqueta: `Aumento de capital ${SIGU.dia}/${String(SIGU.mes).padStart(2, "0")}/${SIGU.anio}`,
    importe: AUMENTO_USD,
  });
  if (ru.limpiadas.length) {
    check(agu.fila === ru.limpiadas[0].fila,
          `el aumento nuevo entra en la fila que dejó libre el anterior (${agu.fila})`);
  }
  check(cerca(capitalDeclarado(wsu, ubicarPatNeto(wsu)).total, DECU + AUMENTO_USD),
        `y el capital en dólares queda en ${f(capitalDeclarado(wsu, ubicarPatNeto(wsu)).total)}`);

  const ru2 = pasarCierreAlInicio(wbu, fpMesSiguiente(SIGU));
  check(ru2.limpiadas.length === 1 && cerca(valu(uu.colCapital), DECU + AUMENTO_USD) &&
        cerca(valu(COL_SUSCRIPTO), DECU + AUMENTO_USD),
        `y el mes siguiente abre con las dos columnas en ${f(valu(COL_SUSCRIPTO))}`);

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
