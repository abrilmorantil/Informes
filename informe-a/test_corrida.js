// Regresión de la corrida mensual completa.
//
// Existe porque la inserción de cuentas nuevas se extrajo de procesar() a la función
// insertarCuentaEnBalance(), que ahora comparte con el panel de categorización. Este
// test corre el camino que corre la app —parsear el export, detectar pendientes,
// procesar— sobre el archivo base real y una copia del export de Onvio, y comprueba
// que el resultado sigue siendo el mismo.
//
//   node informe-a/test_corrida.js
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");
const RAIZ = path.join(BASE, "..");

const XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
global.XLSX = XLSX;
const ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
global.ExcelJS = ExcelJS;
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const { parseOnvioExport } = require(BASE + "/informe-a/parser_onvio.js");
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

const EXPORT = RAIZ + "/ejemplos/SyS_por_Centro_de_costos_mensual_06-26.xls";
const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));

(async () => {
  if (!fs.existsSync(EXPORT)) {
    console.log(`  (salteado: no está ${EXPORT})`);
    process.exit(0);
  }
  const libro = XLSX.read(fs.readFileSync(EXPORT), { type: "buffer", raw: true });
  const filas = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]], { header: 1, raw: true, defval: null });
  const lineas = parseOnvioExport(filas);
  console.log(`export leído: ${lineas.length} líneas`);
  check(lineas.length > 0, "el parser de Onvio devuelve líneas");

  const det = motor.detectarPendientes(lineas, mapeo);
  console.log(`pendientes: ${det.pendientes.length} · centros sin identificar: ${det.sinCc.length}` +
              ` · fuera del balance: ${det.fueraDelBalance.length}`);

  // Toda cuenta pendiente necesita una categoría para poder correr; se les da la
  // primera de la lista, que alcanza para probar que el camino no se rompió.
  const elegidas = {};
  for (const p of det.pendientes) elegidas[p.codigo] = mapeo.categorias[0].desc;

  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const total0 = motor.detectarPendientes(lineas, mapeo).pendientes.length;

  let r = null;
  try {
    r = motor.procesar({
      wb, lineas, mapeo, categoriasElegidas: elegidas, periodo: "2026-07", log: () => {},
    });
  } catch (e) {
    check(false, `procesar() tiró: ${e.message}`);
    console.log(`\n${fallos} FALLA(S)`);
    process.exit(1);
  }

  const res = r.resumen;
  console.log(`resultado: ${res.lineas} líneas · ${res.conocidas} conocidas · ${res.nuevas} nuevas` +
              ` · total ${res.totalSaldo.toFixed(2)}`);
  check(res.lineas > 0, "cargó líneas en la hoja SyS");
  check(res.nuevas === total0, `insertó una fila por cada cuenta pendiente (${res.nuevas} de ${total0})`);
  check(Number.isFinite(res.totalSaldo), "el total es un número");

  // Que el total sea un numero no prueba nada: lo que importa es que la hoja SyS —lo
  // unico que el motor escribe— sume exactamente el saldo del export. Si una linea se
  // pierde por el camino, aca se ve. El export de junio 2026 da 95.792,60.
  const wsSys = wb.getWorksheet("SyS");
  let sumaSys = 0;
  for (const fila of motor.filasDeDatosSys(wsSys).values()) {
    const v = wsSys.getCell(fila, 16).value;          // P = saldo
    if (typeof v === "number") sumaSys += v;
  }
  const saldoOnvio = lineas
    .filter(l => String(l.cuenta_codigo).startsWith("4"))
    .reduce((s, l) => s + (Number(l.saldo) || 0), 0);
  check(Math.abs(sumaSys - saldoOnvio) < 0.005,
        `la hoja SyS suma el saldo del export: ${sumaSys.toFixed(2)} vs ${saldoOnvio.toFixed(2)}`);

  // El mapeo devuelto tiene que quedar consistente: cada cuenta en la fila que dice.
  const wsSs = wb.getWorksheet("Sumas y Saldos");
  let desalineadas = 0;
  for (const [cod, info] of Object.entries(r.mapeo.cuentas)) {
    const v = wsSs.getCell(info.ss_row, 1).value;
    const enArchivo = typeof v === "number" ? String(Math.round(v)) : String(v == null ? "" : v).trim();
    if (enArchivo !== cod) desalineadas++;
  }
  check(desalineadas === 0, `el mapeo quedó alineado con el archivo (${desalineadas} desalineadas)`);

  // Y las cuentas nuevas tienen que haber quedado con nombre: sin nombre no se registran.
  const sinNombre = Object.entries(r.mapeo.cuentas)
    .filter(([, i]) => !String(wsSs.getCell(i.ss_row, 2).value || "").trim());
  check(sinNombre.length === 0, `ninguna cuenta quedó sin nombre en la columna B (${sinNombre.length})`);

  console.log(`\n${fallos ? `${fallos} FALLA(S)` : "todo OK"}`);
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack || e.message); process.exit(1); });
