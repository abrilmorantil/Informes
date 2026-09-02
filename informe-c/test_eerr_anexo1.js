// El EE RR no puede creerle al valor GUARDADO del Anexo I.
//
// `Anexo II!F20 (DEPRECIACIONES) = +'Anexo I'!I24`, y el Anexo I no está cargado a mano: se
// calcula desde SALDOS —`J = -SALDOS!C49` (amortización acumulada al cierre) e `I = J - F`,
// con `F` fija al inicio del ejercicio—. Su valor guardado es el que Excel calculó la ÚLTIMA
// vez que se abrió el archivo, o sea el del mes anterior.
//
// Leyéndolo tal cual, el estado de resultados salía con la amortización del mes pasado.
// Medido en agosto 2026: el Anexo I traía 12.378,87 (julio) cuando con los saldos de agosto
// da 18.318,84, y el EE RR informó "Gastos de Administración" 616.622,64 en vez de 622.562,61
// — 5.939,97 de menos. Sólo salía bien si antes se abría el balance en Excel.
//
//   node informe-c/test_eerr_anexo1.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
// En Node cada archivo es su propio módulo, así que los nombres de hoja compartidos —que en
// el navegador define motor_balances.js para todos— hay que dejarlos en el global a mano.
for (const [k, v] of Object.entries(require(path.join(AQUI, "motor_balances.js")))) {
  if (global[k] === undefined) global[k] = v;
}
const eerr = require(path.join(AQUI, "eerr.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const cerca = (a, b) => Math.abs(a - b) < 0.02;
const N = (c) => {
  const v = c.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
};
const f2 = (x) => (x == null ? "—" : x.toFixed(2));

// el saldo de cada fila de SALDOS, como se lo pasa el motor tras cargar el mes
function saldosDelArchivo(wb) {
  const sa = hojaDistrib(wb);
  return (fila) => {
    const c = sa.getCell(fila, 3);
    if (typeof c.result === "number") return c.result;
    if (typeof c.value === "number") return c.value;
    return 0;
  };
}

(async () => {
  const ruta = path.join(AQUI, "base_dolares.xlsx");

  console.log("=== 1) el Anexo I se calcula, no se lee ===");
  const wb = await abrirWorkbook(fs.readFileSync(ruta));
  const ax = wb.getWorksheet("Anexo I");
  const saldos = saldosDelArchivo(wb);
  const base = eerr.totalesEstadoResultados(wb, saldos);
  const i24 = N(ax.getCell(24, 9));
  console.log(`       Anexo I I24 guardado: ${f2(i24)} · administración: ${f2(base.gastosAdministracion)}`);
  check(base.gastosAdministracion < 0, "el estado calcula los gastos de administración");

  // Se le mete al Anexo I un valor guardado ABSURDO, dejando las fórmulas intactas. Si el
  // motor lo leyera, la administración se movería; como lo recalcula, no se tiene que mover.
  const wb2 = await abrirWorkbook(fs.readFileSync(ruta));
  const ax2 = wb2.getWorksheet("Anexo I");
  let tocadas = 0;
  for (let r = 15; r <= 24; r++) {
    for (const c of [9, 10]) {          // I (del ejercicio) y J (acumulada al cierre)
      const cel = ax2.getCell(r, c);
      if (!cel.formula) continue;
      cel.value = { formula: cel.formula, result: 999999 };
      tocadas++;
    }
  }
  check(tocadas > 5, `se ensucian los ${tocadas} valores guardados del Anexo I, sin tocar sus fórmulas`);
  const conBasura = eerr.totalesEstadoResultados(wb2, saldosDelArchivo(wb2));
  console.log(`       con el Anexo I ensuciado, administración: ${f2(conBasura.gastosAdministracion)}`);
  check(cerca(conBasura.gastosAdministracion, base.gastosAdministracion),
    "la administración NO cambia: el motor recalcula el Anexo I en vez de creerle a lo guardado");

  console.log("\n=== 2) y sigue lo que dicen los saldos del mes ===");
  // si cambian los saldos de las amortizaciones acumuladas, la administración tiene que moverse
  const wb3 = await abrirWorkbook(fs.readFileSync(ruta));
  const ax3 = wb3.getWorksheet("Anexo I");
  // las filas de SALDOS que el Anexo I lee en su columna J
  const filasAmort = [];
  for (let r = 15; r <= 23; r++) {
    const m = new RegExp(`${REF_DISTRIB}!\\$?[A-Z]{1,3}\\$?(\\d+)`, "i")
      .exec(String(ax3.getCell(r, 10).formula || ""));
    if (m) filasAmort.push(+m[1]);
  }
  check(filasAmort.length >= 5,
    `el Anexo I lee ${filasAmort.length} cuentas de amortización acumulada de ${HOJA_DISTRIB}`);

  const baseSaldos = saldosDelArchivo(wb3);
  const EXTRA = -1000;                 // mil dólares más de amortización acumulada
  const conMas = (fila) => baseSaldos(fila) + (filasAmort.includes(fila) ? EXTRA : 0);
  const antes = eerr.totalesEstadoResultados(wb3, baseSaldos).gastosAdministracion;
  const despues = eerr.totalesEstadoResultados(wb3, conMas).gastosAdministracion;
  const esperado = antes - EXTRA * -1 * filasAmort.length;
  console.log(`       antes ${f2(antes)} · con ${filasAmort.length} × ${EXTRA} más: ${f2(despues)}`);
  check(!cerca(antes, despues),
    "mover las amortizaciones acumuladas mueve los gastos de administración");
  check(cerca(despues - antes, EXTRA * filasAmort.length),
    `y se mueve exactamente lo que cambió: ${f2(despues - antes)}`);

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
