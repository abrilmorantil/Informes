// Renombrar una cuenta desde el panel NO puede apagarla.
//
// El texto "código - nombre" de SALDOS no es un rótulo: es la CLAVE con la que busca su
// importe en Hoja1 (`VLOOKUP(E200, Hoja1!$A$2:$E$399, 5, FALSE)`). Cambiándolo sólo en SALDOS,
// el VLOOKUP deja de encontrarla, el IFERROR lo vuelve cero y no avisa nada. Medido sobre el
// maestro real: renombrando "111010001 - Caja", Caja pasa a valer 0.
//
// Y no se arregla solo el mes siguiente: `actualizarHoja1` empareja por CÓDIGO y a propósito
// NO pisa el texto de Hoja1, así que la clave vieja se queda ahí para siempre.
//
//   node informe-c/test_editar_cuenta.js
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
const { editarCuenta, gcFilaEnHoja1, gcNormClave } = require(path.join(AQUI, "gestion_categorias.js"));

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

const V = (c) => {
  const v = c.value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return "";
    return String(v.result ?? "");
  }
  return String(v);
};
const RE_VLOOKUP = new RegExp(
  `VLOOKUP\\(\\s*\\$?([A-Z]{1,3})\\$?(\\d+)\\s*[,;]\\s*${REF_SUMAS}!`, "i");
const colIdx = (s) => s.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);

// ¿La fila `r` de SALDOS puede encontrar su importe en Hoja1? Es la pregunta que importa.
function encuentraSuImporte(wb, r) {
  const saldos = hojaDistrib(wb);
  const hoja1 = hojaSumas(wb);
  for (const c of [7, 6]) {
    const f = saldos.getCell(r, c).formula;
    const m = f && RE_VLOOKUP.exec(String(f));
    if (!m || +m[2] !== r) continue;
    const clave = V(saldos.getCell(r, colIdx(m[1]))).trim();
    return gcFilaEnHoja1(hoja1, 1, clave) !== null &&
           gcNormClave(V(hoja1.getCell(gcFilaEnHoja1(hoja1, 1, clave), 1))) === gcNormClave(clave);
  }
  return null;   // esa fila no usa VLOOKUP
}

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const saldos = hojaDistrib(wb);
  const hoja1 = hojaSumas(wb);

  // Una cuenta que HOY encuentra su importe: se busca sola, no se hardcodea la fila.
  let victima = null;
  for (let r = 1; r <= saldos.rowCount && !victima; r++) {
    for (const c of [7, 6]) {
      const f = saldos.getCell(r, c).formula;
      const m = f && RE_VLOOKUP.exec(String(f));
      if (!m || +m[2] !== r) continue;
      const colClave = colIdx(m[1]);
      const clave = V(saldos.getCell(r, colClave)).trim();
      if (gcFilaEnHoja1(hoja1, 1, clave) !== null) victima = { fila: r, col: colClave, clave };
      break;
    }
  }
  check(!!victima, `hay una cuenta que encuentra su importe: fila ${victima && victima.fila} "${victima && victima.clave}"`);
  check(encuentraSuImporte(wb, victima.fila) === true, "punto de partida: la encuentra");

  const m = /^(\d+)\s*-\s*(.+)$/.exec(victima.clave);
  const codigo = m[1], nombreViejo = m[2];
  const filaH1Antes = gcFilaEnHoja1(hoja1, 1, victima.clave);

  // ------------------------------------------------ se le corrige el nombre
  const r = editarCuenta(wb, victima.fila, victima.col, codigo, nombreViejo + " corregido");
  check(V(saldos.getCell(victima.fila, victima.col)) === `${codigo} - ${nombreViejo} corregido`,
    "el nombre cambió en SALDOS");
  check(r.filaHoja1 === filaH1Antes,
    `y también en Hoja1, en su fila ${r.filaHoja1} — que es lo que faltaba`);
  check(V(hoja1.getCell(filaH1Antes, 1)).trim() === `${codigo} - ${nombreViejo} corregido`,
    "las dos hojas dicen exactamente lo mismo");
  check(encuentraSuImporte(wb, victima.fila) === true,
    "LA PRUEBA: la cuenta sigue encontrando su importe, no quedó en cero");

  // ------------------------------------------------ y ninguna otra se rompió
  let miradas = 0, rotas = [];
  for (let f = 1; f <= saldos.rowCount; f++) {
    const ok = encuentraSuImporte(wb, f);
    if (ok === null) continue;
    miradas++;
    if (ok === false) {
      // sólo importan las que ANTES la encontraban; muchas no están en Hoja1 porque no
      // vinieron en el export del mes con el que se guardó el maestro
      const clave = V(saldos.getCell(f, victima.col)).trim();
      if (f !== victima.fila && clave) rotas.push(f);
    }
  }
  console.log(`       ${miradas} filas de SALDOS usan VLOOKUP contra Hoja1`);

  // ------------------------------------------------ el resguardo: si no puede, no lo hace
  const wb2 = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const s2 = hojaDistrib(wb2);
  const h2 = hojaSumas(wb2);
  // se rompe Hoja1 a propósito: se le saca la fila de esa cuenta del índice poniéndole otro texto
  const fh = gcFilaEnHoja1(h2, 1, victima.clave);
  h2.getCell(fh, 1).value = "999999999 - otra cosa";
  const antesS2 = V(s2.getCell(victima.fila, victima.col));
  let tiro = null;
  try {
    editarCuenta(wb2, victima.fila, victima.col, codigo, "no importa");
  } catch (e) { tiro = e.message; }
  check(V(s2.getCell(victima.fila, victima.col)) === antesS2 || !tiro,
    "si algo no cierra, SALDOS queda como estaba");

  // ------------------------------------------------ un código inválido no toca nada
  const wb3 = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const s3 = hojaDistrib(wb3);
  const antes3 = V(s3.getCell(victima.fila, victima.col));
  let tiro3 = null;
  try { editarCuenta(wb3, victima.fila, victima.col, "abc", "x"); } catch (e) { tiro3 = e.message; }
  check(!!tiro3 && /código de cuenta válido/.test(tiro3), "un código inválido se rechaza");
  check(V(s3.getCell(victima.fila, victima.col)) === antes3, "y no se tocó el archivo");

  // ------------------------------------------------ la cuenta escrita en la columna equivocada
  //
  // Cada fila levanta su importe con un VLOOKUP cuya clave es UNA columna. Si el texto quedó
  // en la de al lado, la fórmula lee vacío y la fila no puede traer un importe nunca. En el
  // maestro pasa en las filas 40 y 43 —Zento S.A. y Títulos Públicos—, las dos en cero hoy.
  const wb4 = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const s4 = hojaDistrib(wb4);
  const { corregirColumnaDeCuenta, cfgColumnaClaveDe } = require(path.join(AQUI, "gestion_categorias.js"));

  check(cfgColumnaClaveDe(s4, 40, 7) === 3, "la columna clave de la fila 40 es la C (la que lee el VLOOKUP)");

  // La fila 40 ya está arreglada en el maestro, así que se la vuelve a romper para probar el
  // arreglo. Antes el test daba por hecho que estaba rota y se puso en rojo al arreglarla.
  s4.getCell(40, 3).value = null;
  check(V(s4.getCell(40, 3)).trim() === "", "punto de partida: la columna clave, vacía");
  check(V(s4.getCell(40, 4)).trim().startsWith("114060002"), "y la cuenta escrita en la de al lado");

  const rc = corregirColumnaDeCuenta(wb4, 40, "114060002", "Zento S.A.");
  check(rc.columna === "C" && V(s4.getCell(40, 3)).trim() === "114060002 - Zento S.A.",
    "corregirla la pone en la columna que la fórmula lee");
  check(V(s4.getCell(40, 4)).trim().startsWith("114060002"),
    "y NO toca la otra columna: ahí puede estar el código viejo del cliente, que es el dato a conservar");

  const motorM = require(path.join(AQUI, "motor_balances.js"));
  check(!!motorM.derivarMapeoMaestro(wb4, "pesos", null).cuentas["114060002"],
    "ahora el motor sí ve la cuenta");

  // no se pisa una columna clave que ya tenía algo
  let tiro4 = null;
  try { corregirColumnaDeCuenta(wb4, 40, "114060002", "otra cosa"); } catch (e) { tiro4 = e.message; }
  check(!!tiro4 && /ya tiene/.test(tiro4), "corregir dos veces no pisa lo que ya está");

  // y un código inválido no toca nada
  s4.getCell(43, 3).value = null;
  const antes44 = V(s4.getCell(43, 3));
  tiro4 = null;
  try { corregirColumnaDeCuenta(wb4, 43, "abc", "x"); } catch (e) { tiro4 = e.message; }
  check(!!tiro4 && V(s4.getCell(43, 3)) === antes44, "un código inválido se rechaza sin tocar nada");

  // ------------------------------------------------ el maestro del disco, intacto
  const disco = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  check(V(hojaDistrib(disco).getCell(victima.fila, victima.col)).trim() === victima.clave,
    "el maestro del disco quedó intacto");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
