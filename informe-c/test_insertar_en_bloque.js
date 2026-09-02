// Una cuenta nueva tiene que quedar DENTRO del subtotal que la suma.
//
// El motor la insertaba siempre debajo de su vecina. Cuando la vecina es la última fila de un
// bloque, "debajo" queda FUERA del rango: Excel no estira un `SUM(G63:G117)` porque se inserte
// en la 118. El importe se carga en SALDOS y no entra a ningún estado.
//
// Pasó en agosto 2026 con dos proveedores nuevos —"Cristales Patagonicos S.R.L" y "Latin
// Mining S.A.S"— que entraron en las filas 118 y 119. El subtotal de proveedores seguía
// sumando hasta la 117 y el pasivo del balance quedó 1.915.260,00 por encima de Onvio. El
// activo y el patrimonio neto cerraban perfecto, así que la diferencia no decía de dónde venía.
//
//   node informe-c/test_insertar_en_bloque.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const { insertRowEn } = require(path.join(AQUI, "formula_hojas.js"));
global.insertRowEn = insertRowEn;
const M = require(path.join(AQUI, "motor_balances.js"));
// En Node cada archivo es su propio módulo, así que los nombres de hoja compartidos —que en
// el navegador define motor_balances.js para todos— hay que dejarlos en el global a mano.
for (const [k, v] of Object.entries(M)) { if (global[k] === undefined) global[k] = v; }

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const T = (c) => {
  const v = c.value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return typeof v.result === "string" ? v.result : "";
    return String(v.result ?? "");
  }
  return typeof v === "string" ? v : "";
};
// el rango del subtotal de proveedores, sea cual sea la fila donde esté
function rangoProveedores(ws) {
  let r = null;
  ws.eachRow({ includeEmpty: false }, (row, fila) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (r) return;
      const f = cell.formula && String(cell.formula).replace(/\s+/g, "");
      const m = f && /^\+?SUM\(([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)\)$/i.exec(f);
      if (!m) return;
      const desde = +m[2], hasta = +m[4];
      // el de proveedores es el que abarca filas cuya cuenta empieza con 21101
      let prov = 0;
      for (let x = desde; x <= hasta; x++) {
        const t = (T(ws.getCell(x, 3)) || T(ws.getCell(x, 4))).trim();
        if (/^21101/.test(t)) prov++;
      }
      if (prov >= 10) r = { fila, desde, hasta, formula: f };
    });
  });
  return r;
}

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const ws = hojaDistrib(wb);
  const mapeo = M.derivarMapeoMaestro(wb, "pesos");

  const antes = rangoProveedores(ws);
  check(!!antes, `el bloque de proveedores va de la fila ${antes && antes.desde} a la ${antes && antes.hasta}`);

  console.log("\n=== la última fila del bloque cierra un subtotal ===");
  check(M.filaCierraUnSubtotal(wb, antes.hasta, M.PARAMS.pesos),
    `la fila ${antes.hasta} es la última de un rango que la suma`);
  check(!M.filaCierraUnSubtotal(wb, antes.hasta - 1, M.PARAMS.pesos),
    `la ${antes.hasta - 1}, que está en el medio, no`);

  console.log("\n=== una cuenta nueva de proveedores entra ADENTRO del subtotal ===");
  // un código más alto que todos los del bloque: su vecina va a ser la última fila
  const nueva = { codigo: "211019999", nombre: "Proveedor Nuevo De Prueba SA",
                  capitulo: "PASIVO", saldo_ars: 1234.56 };
  const r = M.insertarCuentaEnSaldos(wb, mapeo, nueva, "pesos", () => {});
  const despues = rangoProveedores(ws);
  console.log(`       antes ${antes.formula}  ->  después ${despues.formula}`);
  check(r.filaNueva >= despues.desde && r.filaNueva <= despues.hasta,
    `la cuenta nueva quedó en la fila ${r.filaNueva}, dentro del rango ${despues.desde}-${despues.hasta}`);
  check(despues.hasta === antes.hasta + 1,
    `y el subtotal se estiró una fila (${antes.hasta} -> ${despues.hasta})`);

  // y su cuenta está escrita donde corresponde
  const txt = (T(ws.getCell(r.filaNueva, 3)) || T(ws.getCell(r.filaNueva, 4))).trim();
  check(txt.startsWith("211019999"), `la fila nueva tiene su cuenta: "${txt}"`);

  console.log("\n=== una cuenta cuya vecina NO cierra bloque sigue entrando debajo ===");
  {
    const wb2 = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
    const mapeo2 = M.derivarMapeoMaestro(wb2, "pesos");
    // una del ACTIVO: sus vecinas son filas sueltas con VLOOKUP, sin subtotal que las cierre
    const cuenta = { codigo: "114060099", nombre: "Otro Deudor De Prueba",
                     capitulo: "ACTIVO", saldo_ars: 10 };
    const vecina = mapeo2.cuentas["114060002"];
    const cierra = M.filaCierraUnSubtotal(wb2, vecina.fila, M.PARAMS.pesos);
    const r2 = M.insertarCuentaEnSaldos(wb2, mapeo2, cuenta, "pesos", () => {});
    check(!cierra ? r2.filaNueva > vecina.fila : true,
      cierra ? "(su vecina cierra un bloque, así que entra adentro)"
             : `entra debajo de su vecina, como siempre (fila ${r2.filaNueva})`);
  }

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
