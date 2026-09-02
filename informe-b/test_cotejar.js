// Cotejar los saldos de un informe importado contra el plan de cuentas.
//
// Dos cosas distintas que antes se trataban igual —ninguna quedaba guardada— y que tienen
// consecuencias opuestas:
//
//   - Una cuenta `ocultar_si_cero` no sale impresa mientras esté en cero. Que falte es normal
//     y su saldo es cero. Si no se guarda, el mes siguiente sale AMARILLA pidiendo que la
//     carguen a mano, sin tener nada que revisar.
//   - Cualquier otra cuenta que falte significa que el informe está incompleto. Ahí completar
//     con cero sería peor que no hacer nada: convertiría plata en ceros que parecen datos.
//     Se avisa y se deja que la persona decida.
//
//   node informe-b/test_cotejar.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;
const { cotejarConElPlan } = require(path.join(AQUI, "estado_b.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

const PLAN = [
  { code: "111010001", description: "Caja" },
  { code: "111020001", description: "ICBC S.A. c/c" },
  { code: "42433000", description: "Gs. Importación", ocultar_si_cero: true },
  { code: "423060000", description: "Diferencia de cambio $", ocultar_si_cero: true },
  { code: "42101000", description: "Honorarios legales" },
];

console.log("=== 1) las que se ocultan por estar en cero se completan con cero ===");
{
  const r = cotejarConElPlan({ "111010001": 1014.2, "111020001": 8064.94, "42101000": 0 }, PLAN);
  check(r.completados["42433000"] === 0 && r.completados["423060000"] === 0,
    "las dos `ocultar_si_cero` que faltaban quedan guardadas en 0");
  check(r.enCero.length === 2, `y se informan como completadas (${r.enCero.map(x => x.code).join(", ")})`);
  check(r.faltan.length === 0, "no falta ninguna otra");
  check(r.completados["111010001"] === 1014.2,
    "y las que sí venían conservan su saldo, no se pisan");
  check(Object.keys(r.completados).length === PLAN.length,
    "el plan queda cubierto entero: ninguna cuenta se va a pintar de amarillo sin motivo");
}

console.log("\n=== 2) una cuenta que falta y NO es de las que se ocultan se avisa, no se rellena ===");
{
  const r = cotejarConElPlan({ "111010001": 1014.2 }, PLAN);
  check(r.faltan.map(f => f.code).sort().join(",") === "111020001,42101000",
    `avisa las dos que faltan de verdad (${r.faltan.map(f => f.code).join(", ")})`);
  check(!Object.prototype.hasOwnProperty.call(r.completados, "111020001"),
    "y NO las inventa en cero: rellenar a ciegas volvería plata en ceros que parecen datos");
  check(r.completados["42433000"] === 0,
    "las `ocultar_si_cero` se completan igual, que ésas sí sabemos por qué faltan");
}

console.log("\n=== 3) casos de borde ===");
{
  const r = cotejarConElPlan({}, []);
  check(Object.keys(r.completados).length === 0 && r.faltan.length === 0,
    "sin plan y sin saldos no inventa nada");
  const r2 = cotejarConElPlan({ "111010001": 0 }, PLAN);
  check(r2.completados["111010001"] === 0 && !r2.faltan.some(f => f.code === "111010001"),
    "un saldo que YA venía en cero cuenta como presente, no como faltante");
  const r3 = cotejarConElPlan(null, PLAN);
  check(r3.faltan.length + r3.enCero.length === PLAN.length,
    "sin saldos, todas las del plan salen listadas");
}

// ---------------------------------------------------------------- contra el caso real
console.log("\n=== 4) el caso que lo motivó: el informe de julio de 102 cuentas ===");
{
  const mapping = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));
  const estado = JSON.parse(fs.readFileSync(path.join(AQUI, "estado_b.json"), "utf8"));
  const r = cotejarConElPlan(estado.saldos, mapping);
  console.log(`       el plan tiene ${mapping.length} cuentas; lo guardado hoy, ` +
              `${Object.keys(estado.saldos).length}`);
  console.log(`       se completan en cero: ${r.enCero.length}   quedan avisadas: ${r.faltan.length}`);
  check(r.faltan.length > 0,
    "con lo que hay guardado hoy el aviso salta: al informe importado le faltaban cuentas del plan");
  check(r.enCero.length <= mapping.filter(x => x.ocultar_si_cero).length,
    "y sólo se completan las que el plan marca como ocultables en cero");
}

console.log("\n=== 5) el lector usa las DOS columnas de cuenta ===");
{
  // El informe tiene "Cuenta real (Onvio)" en A y "Cuenta Contable" en B. La A queda VACÍA en
  // las cuentas madre y en el renglón de proveedores, que se alimentan de varias cuentas
  // reales. Leyendo sólo la A se salteaban 57 filas de 197 y el balance importado no cerraba:
  // daba −659.684,53 en vez de −0,02.
  global.ExcelJS = require(path.join(AQUI, "exceljs.min.js"));
  const imp = require(path.join(AQUI, "importar_saldos.js"));
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getCell("A5").value = "Cuenta real (Onvio)";
  ws.getCell("B5").value = "Cuenta Contable";
  ws.getCell("C5").value = "Saldo anterior";
  ws.getCell("D5").value = "Debitos del Mes Dolares";
  ws.getCell("E5").value = "Creditos del mes Dolares";
  // una fila normal: la cuenta real en A y la del cliente en B
  ws.getCell("A6").value = "124010001 - Muebles y Utiles";
  ws.getCell("B6").value = "12402000 - Muebles y Utiles";
  ws.getCell("C6").value = 100; ws.getCell("D6").value = 20; ws.getCell("E6").value = 5;
  // una cuenta madre: A vacía, la cuenta va en B
  ws.getCell("B7").value = "42207000 - Seguridad";
  ws.getCell("C7").value = 50; ws.getCell("D7").value = 0; ws.getCell("E7").value = 10;

  const r = imp.leerSaldosDeBalcomp(wb);
  check(r.cuentas === 2, `lee las dos filas, no sólo la que tiene la columna A (leyó ${r.cuentas})`);
  check(r.saldos["124010001"] === 115,
    "de la fila normal toma el código REAL de la columna A, que es la clave con la que se guarda");
  check(r.saldos["42207000"] === 40,
    "y de la cuenta madre, que tiene la A vacía, lo toma de la B");
  check(!Object.prototype.hasOwnProperty.call(r.saldos, "12402000"),
    "el código del cliente NO se guarda: estado_b se indexa por el real");
}

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
