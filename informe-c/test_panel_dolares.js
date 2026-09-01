// El panel de configuración de cuentas del balance en DÓLARES, sobre el maestro real.
//
// Lo que se prueba es justamente lo que no se puede copiar de pesos: en dólares no hay cuentas
// madre con subcuentas adentro, así que la unidad del Anexo II son las líneas de RESULTADOS; y
// una línea puede estar legítimamente fuera del Anexo II porque la lee el estado de resultados.
// Distinguir esas de las que no lee nadie es todo el valor del aviso.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const { insertRowEn, borrarFilaEn } = require(path.join(AQUI, "formula_hojas.js"));
global.insertRowEn = insertRowEn;
global.borrarFilaEn = borrarFilaEn;
const M = require(path.join(AQUI, "motor_balances.js"));
global.PARAMS = M.PARAMS;
global.derivarMapeoMaestro = M.derivarMapeoMaestro;
global.lineasDeNota4 = M.lineasDeNota4;
global.filasQueAgrega = M.filasQueAgrega;
const { derivarConfigBalance } = require(path.join(AQUI, "config_balances.js"));
global.derivarConfigBalance = derivarConfigBalance;
const a2 = require(path.join(AQUI, "anexo2.js"));
for (const k of Object.keys(a2)) global[k] = a2[k];
const clas = require(path.join(AQUI, "clasificacion.js"));
const pcc = require(path.join(AQUI, "panel_config_cuentas.js"));
for (const k of Object.keys(pcc)) global[k] = pcc[k];
const pcd = require(path.join(AQUI, "panel_dolares.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

(async () => {
  const clasif = clas.indexarClasificacion(JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8")));
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_dolares.xlsx")));
  const mapeo = M.derivarMapeoMaestro(wb, "dolares", clasif);

  console.log("=== 1) la unidad del Anexo II en dólares son las líneas, no las madres ===");
  check(M.madresResultados(wb, "dolares").length === 0,
    "el maestro de dólares no tiene cuentas madre con subcuentas (por eso no se puede copiar pesos)");
  const lineas = pcd.pcdLineasResultados(wb, mapeo);
  check(lineas.length > 100 && lineas.every(l => String(l.codigo).startsWith("4") && l.fila > 0),
    `las líneas de RESULTADOS salen del mapeo (${lineas.length})`);
  check(lineas.every(l => l.nombre && !/^\d/.test(l.nombre)),
    "cada una trae su nombre separado del código");

  const v = a2.a2Verificar(wb, lineas);
  console.log(`       ${v.total} líneas · ${v.unaVez} en el Anexo II una vez · ` +
              `${v.dobles.length} dos o más · ${v.sinAnexo.length} fuera`);
  check(v.dobles.length === 0,
    v.dobles.length
      ? `hay ${v.dobles.length} línea(s) contadas dos veces: ` + v.dobles.map(d => d.nombre).join(", ")
      : "ninguna línea entra dos veces al Anexo II: no se cuenta doble ningún gasto");

  console.log("\n=== 2) estar fuera del Anexo II no es lo mismo que estar suelta ===");
  const lectores = pcd.pcdLectoresFuera(wb);
  const fuera = v.sinAnexo;
  const conLector = fuera.filter(l => (lectores.get(l.filaSaldos) || []).length);
  const sueltas = fuera.filter(l => !(lectores.get(l.filaSaldos) || []).length);
  check(conLector.length > 0,
    `${conLector.length} de las ${fuera.length} que están fuera las lee otra hoja, y están bien: ` +
    conLector.map(l => l.nombre).join(", "));
  check(conLector.every(l => (lectores.get(l.filaSaldos) || []).some(x => x.hoja === "Resultados")),
    "y quien las lee es el estado de resultados");
  check(sueltas.length < fuera.length,
    `las que no lee nadie son ${sueltas.length}, no las ${fuera.length} de afuera`);

  // El total de la hoja abarca TODAS las filas: si contara como lector, ninguna línea
  // parecería suelta nunca y el aviso no serviría para nada.
  check(sueltas.length > 0,
    "el total de la hoja no cuenta como lector (si contara, no quedaría ninguna suelta)");

  console.log("\n=== 3) el aviso ===");
  const html = pcd.pcdSueltasHtml(sueltas);
  check(html.includes("status-msg bad") && sueltas.every(s => html.includes(s.codigo)),
    "el bloque las lista una por una, en rojo");
  check(pcd.pcdSueltasHtml([]).includes("status-msg ok"),
    "y si no hubiera ninguna, lo dice en verde");

  console.log("\n=== 4) mover una línea de columna ===");
  // a2Mover tenía dos cosas cableadas para pesos: la columna G y exigir CERO líneas fuera del
  // anexo. Las dos rompían en dólares, donde la columna es C y hay 6 legítimamente fuera.
  const conLinea = a2.a2Mapa(wb, lineas).lineas.find(l => l.donde.length === 1);
  const origen = conLinea.donde[0];
  const cols = a2.a2Columnas(wb);
  const otra = cols.find(c => c.col !== origen.col);
  const r = a2.a2Mover({
    wb, madres: lineas, filaSaldos: conLinea.filaSaldos,
    anexoFilaDestino: origen.anexoFila, colDestino: otra.col, colSaldosDefecto: "C",
  });
  check(r.hasta.columna === otra.nombre,
    `"${conLinea.nombre}" se movió de ${origen.columna} a ${otra.nombre} (con 6 líneas fuera del anexo, que antes lo bloqueaba)`);
  const v2 = a2.a2Verificar(wb, lineas);
  check(v2.unaVez === v.unaVez && v2.dobles.length === 0 && v2.sinAnexo.length === v.sinAnexo.length,
    "y el Anexo II quedó igual de sano que antes");
  const f = String(wb.getWorksheet("Anexo II").getCell(origen.anexoFila, otra.col).formula || "");
  check(/SALDOS!C\d+/.test(f) && !/SALDOS!G\d+/.test(f),
    `la referencia quedó a la columna C, que es la de dólares: ${f.slice(0, 40)}`);

  // y vuelve a su lugar
  a2.a2Mover({
    wb, madres: lineas, filaSaldos: conLinea.filaSaldos,
    anexoFilaDestino: origen.anexoFila, colDestino: origen.col, colSaldosDefecto: "C",
  });
  check(a2.a2Verificar(wb, lineas).unaVez === v.unaVez, "vuelve sin romper nada");

  console.log("\n=== 5) los pendientes ===");
  const cfg = derivarConfigBalance(wb, "dolares", mapeo, M.lineasDeNota4(wb), M.PARAMS.dolares,
    { filasQueAgrega: M.filasQueAgrega });
  const pend = cfg.avisos.filter(a => a.tipo !== "linea_sin_cuenta");
  console.log("       " + pend.map(a => `${a.tipo}${a.code ? " " + a.code : ""}`).join(", "));
  check(pcc.pccResumenHtml(cfg, pend.length, 0).includes("Pendientes de revisar"),
    `el mismo chip que en pesos, con ${pend.length} pendiente(s)`);
  check(pend.every(a => a.tipo === "codigo_repetido"),
    "y hoy los pendientes del maestro de dólares son códigos repetidos");

  console.log(fallos ? `\n${fallos} FALLAS` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
