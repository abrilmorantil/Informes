// Las fechas que los informes muestran adentro tienen que seguir al período que se emite.
// Este test cubre el reescritor de `fechas.js` y, sobre todo, que los maestros arranquen con
// sus fechas al día: la regla es que sólo se reescribe lo que coincide con el cierre actual,
// así que una fecha que se atrasa no vuelve sola.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "informe-a/vendor/exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "informe-a/formula_utils.js"));
const F = require(path.join(AQUI, "fechas.js"));

let fallas = 0;
function check(cond, msg) {
  console.log(`  ${cond ? "OK  " : "FALLA"} ${msg}`);
  if (!cond) fallas++;
}

const JUNIO = { anio: 2026, mes: 6, dia: 30 };
const JULIO = { anio: 2026, mes: 7, dia: 31 };

// texto -> cómo queda al pasar de junio a julio (null = no lo tocó)
const reescribir = (t, viejo = JUNIO, nuevo = JULIO) => F.fpReescribirTexto(t, viejo, nuevo, []);

(async () => {
  console.log("=== 1) el año de dos dígitos ===");
  // Es el caso del encabezado "AL 31.5.26" del Anexo I de pesos: mientras no se reconocía,
  // esa celda quedaba afuera del reescritor y nadie se enteraba.
  check(reescribir("AL 30.6.26") === "AL 31.7.26",
    'con dos dígitos: "AL 30.6.26" pasa a "AL 31.7.26"');
  check(reescribir("AL 30.6.26").length === "AL 31.7.26".length,
    "y el año sigue con dos dígitos, no se convierte en 2026");
  check(reescribir("Total al 30.06.2026") === "Total al 31.07.2026",
    "con cuatro dígitos sigue andando igual que antes");
  check(JSON.stringify(F.fpCierreDeCelda({ value: "AL 30.6.26" })) === JSON.stringify(JUNIO),
    "y leer el cierre de un texto con dos dígitos da 2026, no el año 26");

  console.log("\n=== 2) lo que NO se toca ===");
  check(reescribir("Aumento de capital 16/06/2026") === null,
    "las fechas con barras nunca: son hechos puntuales, no el período");
  check(reescribir("AL 30.04.2026") === null,
    "una fecha que no es el cierre queda como está");
  check(reescribir("1.234.56") === null,
    "un número con puntos no se confunde con una fecha");
  check(reescribir("31.03.04") === null,
    "una fecha vieja de otro año tampoco se toca");

  console.log("\n=== 3) la redacción de cada archivo se conserva ===");
  check(reescribir("Al 30 de junio 2026") === "Al 31 de julio 2026",
    'pesos escribe "Al 30 de junio 2026" y sigue sin el "de" antes del año');
  check(reescribir("Al 30 de junio de 2026") === "Al 31 de julio de 2026",
    'dólares escribe "de 2026" y lo conserva');
  check(reescribir("finalizado el 30 de Junio de 2026") === "finalizado el 31 de Julio de 2026",
    "y el mes con mayúscula sigue con mayúscula");

  console.log("\n=== 4) las fechas que quedaron sin tocar se avisan ===");
  const otras = [];
  F.fpReescribirTexto("AL 30.04.2026", JUNIO, JULIO, otras);
  check(otras.includes("2026-04-30"),
    "el reescritor devuelve las fechas que vio y no cambió, para poder mostrarlas");

  console.log("\n=== 5) los maestros arrancan con sus fechas al día ===");
  for (const arch of ["base_pesos.xlsx", "base_dolares.xlsx"]) {
    const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "informe-c", arch)));
    const cierre = F.fpCierreDeCelda(wb.getWorksheet("Balance").getCell("F5"));
    check(!!cierre, `${arch}: se puede leer a qué fecha está cerrado (${F.fpDescribir(cierre)})`);
    if (!cierre) continue;
    const sig = F.fpMesSiguiente(cierre);

    // Toda celda de texto con pinta de fecha de período tiene que moverse con el cierre. Las
    // de barras y las de los encabezados históricos ocultos quedan afuera a propósito.
    const OCULTAS = new Set(["Balance!G5", "Anexo II!J8"]);
    const atrasadas = [];
    for (const ws of wb.worksheets) {
      if (ws.name === "SALDOS" || ws.name === "Hoja1") continue;
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        row.eachCell({ includeEmpty: false }, (cell, ci) => {
          if (cell.formula) return;
          const v = cell.value;
          const t = typeof v === "string" ? v
            : (v && typeof v === "object" && v.richText ? v.richText.map(x => x.text).join("") : null);
          if (!t) return;
          const dir = `${ws.name}!${ws.getColumn(ci).letter}${r}`;
          if (OCULTAS.has(dir)) return;
          if (!/\d{1,2}\.\d{1,2}\.\d{2,4}|\d{1,2}\s+de\s+[A-Za-zÁÉÍÓÚáéíóú]+/.test(t)) return;
          if (F.fpReescribirTexto(t, cierre, sig, []) === null) atrasadas.push(`${dir} "${t.trim()}"`);
        });
      });
    }
    check(atrasadas.length === 0,
      atrasadas.length
        ? `${arch}: hay ${atrasadas.length} fecha(s) que el mes que viene no se van a actualizar: ${atrasadas.join(" , ")}`
        : `${arch}: todas sus fechas de período se van a actualizar solas`);
  }

  console.log(fallas ? `\n${fallas} FALLAS` : "\ntodo OK");
  process.exit(fallas ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
