// Las dos hojas de trabajo del maestro cambiaron de nombre en septiembre de 2026:
//   Hoja1  -> "Balance de sumas y saldos"     (lo que manda Onvio, tal cual)
//   SALDOS -> "Distribución por línea"        (dónde sale impresa cada cuenta)
//
// Este test existe porque la forma de romperlo es silenciosa: si un archivo del motor vuelve
// a escribir "SALDOS!" a mano en una expresión, esa fórmula deja de encontrarse, el IFERROR
// del maestro devuelve vacío y la línea del balance sale en CERO sin ningún aviso.
//
// Se controlan tres cosas:
//   1. los maestros que se publican tienen las hojas con el nombre nuevo;
//   2. ningún archivo del motor nombra las hojas a mano: todos pasan por las constantes;
//   3. un maestro VIEJO, con los nombres de antes, se sigue resolviendo igual.
//
//   node informe-c/test_nombres_hojas.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const M = require(path.join(AQUI, "motor_balances.js"));
for (const [k, v] of Object.entries(M)) { if (global[k] === undefined) global[k] = v; }

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

(async () => {
  console.log("=== 1) los maestros publicados ===");
  for (const moneda of ["pesos", "dolares"]) {
    const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, `base_${moneda}.xlsx`)));
    const nombres = wb.worksheets.map(w => w.name);
    check(nombres.indexOf(HOJA_SUMAS) >= 0 && nombres.indexOf(HOJA_DISTRIB) >= 0,
      `base_${moneda}: ${nombres.join(" | ")}`);
    check(!!hojaSumas(wb) && !!hojaDistrib(wb), `  y los dos resolutores las encuentran`);

    // ninguna fórmula puede nombrar una hoja que no existe: así se ve un renombre a medias
    const hojas = new Set(nombres);
    const nombra = /(?:'([^']+)'|\b([A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*))!\$?[A-Z]{1,3}\$?\d/g;
    const huerfanas = [];
    for (const ws of wb.worksheets) {
      ws.eachRow({ includeEmpty: false }, (row, r) => row.eachCell({ includeEmpty: false }, (c, ci) => {
        const v = c.value;
        if (!(v && typeof v === "object" && typeof v.formula === "string")) return;
        let m; nombra.lastIndex = 0;
        while ((m = nombra.exec(v.formula))) {
          const h = (m[1] || m[2]).trim();
          if (!hojas.has(h)) huerfanas.push(`${ws.name}!${ci},${r} -> ${h}`);
        }
      }));
    }
    check(huerfanas.length === 0,
      `  ninguna fórmula nombra una hoja inexistente${huerfanas.length ? ": " + huerfanas.slice(0, 4).join(" · ") : ""}`);
  }

  console.log("\n=== 2) el motor no escribe los nombres a mano ===");
  const sueltos = [];
  for (const nombre of fs.readdirSync(AQUI).filter(x => /\.js$/.test(x) && !/^test_/.test(x))) {
    const t = fs.readFileSync(path.join(AQUI, nombre), "utf8");
    t.split(/\r?\n/).forEach((l, i) => {
      const limpio = l.trim();
      if (limpio.startsWith("//") || limpio.startsWith("*")) return;      // los comentarios cuentan la historia
      if (/HOJA_SUMAS_ANTES|HOJA_DISTRIB_ANTES|REF_SUMAS =|REF_DISTRIB =/.test(l)) return;
      if (/\bSALDOS!|\bHoja1!|"SALDOS"|"Hoja1"/.test(l)) sueltos.push(`${nombre}:${i + 1}`);
    });
  }
  check(sueltos.length === 0,
    sueltos.length ? `hay ${sueltos.length} lugar(es) con el nombre escrito a mano: ${sueltos.join(", ")}`
                   : "ningún archivo del motor nombra las hojas a mano");

  console.log("\n=== 3) un maestro con los nombres VIEJOS se sigue entendiendo ===");
  {
    const wb = new ExcelJS.Workbook();
    const h1 = wb.addWorksheet(HOJA_SUMAS_ANTES);
    const sa = wb.addWorksheet(HOJA_DISTRIB_ANTES);
    h1.getCell("A2").value = "111010001 - Caja";
    sa.getCell("C4").value = "111010001 - Caja";
    check(hojaSumas(wb) === h1 && hojaDistrib(wb) === sa,
      `"${HOJA_SUMAS_ANTES}" y "${HOJA_DISTRIB_ANTES}" siguen resolviendo`);
    check(new RegExp(REF_DISTRIB + "!", "i").test("+SALDOS!G12") &&
          new RegExp(REF_DISTRIB + "!", "i").test("+'Distribución por línea'!G12") &&
          new RegExp(REF_DISTRIB + "!", "i").test("+'SALDOS'!G12"),
      "y las expresiones reconocen las tres formas: SALDOS!, 'SALDOS'! y el nombre nuevo");
    check(refDeHoja(HOJA_DISTRIB) === `'${HOJA_DISTRIB}'` && refDeHoja("Hoja1") === "Hoja1",
      "al escribir una fórmula, el nombre con espacios va entre comillas y el otro no");
  }

  console.log("\n=== 4) el nombre entra en una pestaña de Excel ===");
  for (const n of [HOJA_SUMAS, HOJA_DISTRIB]) {
    check(n.length <= 31, `"${n}" mide ${n.length} caracteres (Excel admite hasta 31)`);
  }

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error(e.stack); process.exit(1); });
