// Ocultar los centros de costo sin movimiento.
//
// La regla es "si la columna va a mostrar cero, se oculta". El error que apareció en agosto
// 2026 fue que se decidía con OTRO número: |debe| + |haber| de las líneas crudas del export.
// No es lo mismo que lo que muestra la columna — una línea puede tener movimiento y no llegar
// a Dist.de gastos (la cuenta no es de resultado, no tiene categoría, o quedó excluida) — así
// que quedaban columnas a la vista mostrando cero.
//
// Y el caso peor: un centro de costo cuyo NOMBRE no resuelve. Sus líneas no se cargan, su
// columna queda en cero y se ocultaba, con lo que la plata que faltaba dejaba de verse.
// Pasó con "Proyecto Lonco Vaca- Palenque" contra el bloque "LONCO VACA - PELENQUE".
//
//   node informe-a/test_ocultar.js
const path = require("path");
const fs = require("fs");
const BASE = path.join(__dirname, "..");

global.XLSX = require(BASE + "/informe-a/vendor/xlsx.full.min.js");
global.ExcelJS = require(BASE + "/informe-a/vendor/exceljs.min.js");
const { abrirWorkbook } = require(BASE + "/informe-a/formula_utils.js");
const motor = require(BASE + "/informe-a/motor.js");

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? "  OK  " : " FALLA"} ${m}`); if (!ok) fallos++; };

const mapeo = JSON.parse(fs.readFileSync(BASE + "/informe-a/mapeo.json", "utf8"));
const colIdx = (col) => col.toUpperCase().split("")
  .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
const nombreDe = (info) => (info && (info.nombre_balance || info.nombre)) || String(info);

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(BASE + "/informe-a/base_actual.xlsx"));
  const wsDist = wb.getWorksheet("Dist.de gastos");

  const columnas = Object.entries(mapeo.dist_col_to_cc)
    .map(([col, info]) => ({ col, nombre: nombreDe(info) }));
  check(columnas.length > 0, `el mapeo tiene ${columnas.length} columnas de centro de costo`);

  const estado = () => {
    const o = {};
    for (const c of columnas) o[c.nombre] = !!wsDist.getColumn(colIdx(c.col)).hidden;
    return o;
  };
  const oculta = (nombre) => !!wsDist.getColumn(colIdx(columnas.find(c => c.nombre === nombre).col)).hidden;

  // ------------------------------------------------ se decide con el aporte a la columna
  const uno = columnas[0].nombre, dos = columnas[1].nombre, tres = columnas[2].nombre;

  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [uno]: 1234.5, [dos]: 0 }, () => {});
  check(!oculta(uno), `"${uno}" aporta 1.234,50 y queda visible`);
  check(oculta(dos), `"${dos}" aporta 0 y se oculta`);
  check(oculta(tres), `"${tres}" no aparece en el aporte y también se oculta`);

  // vuelve a mostrarse el mes que sí mueve
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [dos]: 10 }, () => {});
  check(!oculta(dos), `"${dos}" vuelve a mostrarse cuando aporta algo`);
  check(oculta(uno), `y "${uno}" se oculta cuando deja de aportar`);

  // un aporte de menos de un centavo no cuenta como movimiento
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [uno]: 0.001 }, () => {});
  check(oculta(uno), `un aporte de 0,001 no alcanza: "${uno}" sigue oculto`);

  // ------------------------------------------------ el número que se usa NO es debe+haber
  // Una línea con movimiento que no llega a Dist.de gastos no tiene que mostrar la columna.
  // Antes se le pasaban las líneas crudas y esto daba "visible".
  const antes = estado();
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, {}, () => {});
  const todas = columnas.every(c => oculta(c.nombre));
  check(todas, "sin ningún aporte se ocultan todas las columnas");
  check(Object.keys(antes).length === columnas.length, "y el estado se pudo leer para las " + columnas.length);

  // ------------------------------------------------ la firma ya no acepta líneas crudas
  // Si alguien vuelve a pasarle un array (como antes), no puede quedar "todo visible" en
  // silencio: un array no tiene las claves por nombre, así que nada aporta y se oculta todo.
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, [
    { cc_nombre_onvio: uno, debe: 999, haber: 0 },
  ], () => {});
  check(columnas.every(c => oculta(c.nombre)),
    "pasarle líneas crudas ya no muestra columnas: oculta todo en vez de mentir");

  // ------------------------------------------------ con un centro sin resolver: mostrar sí, ocultar no
  //
  // Mostrar una columna que tiene movimiento nunca puede estar mal; ocultar sí. La primera
  // versión de este resguardo salteaba la función entera, y con eso las columnas que TENÍAN
  // movimiento pero venían ocultas del mes anterior se quedaban escondidas. Pasó en agosto
  // 2026 con CERRO LA MINA, CERRO ABANICO y LOS MENUCOS: tenían saldo y no se veían.
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, {}, () => {});          // todas ocultas
  check(columnas.every(c => oculta(c.nombre)), "punto de partida: todas ocultas");

  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [uno]: 500 }, () => {}, { soloMostrar: true });
  check(!oculta(uno), `"${uno}" aporta y se muestra, aunque no se pueda ocultar nada`);
  check(oculta(dos), `y "${dos}", que no aporta, queda como estaba`);

  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [dos]: 500 }, () => {}, { soloMostrar: true });
  check(!oculta(uno), `"${uno}" NO se oculta aunque deje de aportar: con un centro sin resolver no se oculta nada`);
  check(!oculta(dos), `y "${dos}" se muestra porque ahora sí aporta`);

  // ------------------------------------------------ las DOS hojas, no sólo Dist.de gastos
  //
  // Un centro de costo se ve en dos lados: una columna en `Dist.de gastos` y un bloque de tres
  // (Debe/Haber/Saldo) en `Sumas y Saldos`. Esto tocaba sólo la primera, así que un centro que
  // este mes empezaba a mover aparecía en la distribución y seguía escondido en Sumas y
  // Saldos, que es la hoja donde se mira cuenta por cuenta.
  console.log("\n=== y en Sumas y Saldos, que es la otra hoja donde se ven ===");
  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const bloque = (nombre) => (mapeo.cc_blocks || []).find(b => b.nombre_balance === nombre);
  const colsSs = (nombre) => {
    const b = bloque(nombre);
    return b ? [b.col_debe, b.col_haber, b.col_saldo] : [];
  };
  const ocultaSs = (nombre) => colsSs(nombre).every(c => !!wsSs.getColumn(colIdx(c)).hidden);
  const visibleSs = (nombre) => colsSs(nombre).every(c => !wsSs.getColumn(colIdx(c)).hidden);

  check(colsSs(uno).length === 3, `"${uno}" ocupa 3 columnas en Sumas y Saldos (${colsSs(uno).join("/")})`);

  // se lo esconde a mano en las dos hojas, como quedaría después de un mes sin movimiento
  for (const c of colsSs(uno)) wsSs.getColumn(colIdx(c)).hidden = true;
  wsDist.getColumn(colIdx(columnas.find(c => c.nombre === uno).col)).hidden = true;
  check(ocultaSs(uno), `punto de partida: "${uno}" escondido también en Sumas y Saldos`);

  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [uno]: 900 }, () => {}, {}, wsSs);
  check(!oculta(uno), `"${uno}" vuelve a mostrarse en Dist.de gastos`);
  check(visibleSs(uno), "y sus tres columnas de Sumas y Saldos también");
  check(ocultaSs(dos), `"${dos}", que no aporta, se esconde en las dos hojas`);

  // y sin pasarle la hoja sigue andando igual que antes, sin tocarla
  const antesSs = colsSs(uno).map(c => !!wsSs.getColumn(colIdx(c)).hidden);
  motor.ocultarCentrosSinMovimiento(wsDist, mapeo, { [dos]: 900 }, () => {});
  check(colsSs(uno).every((c, i) => !!wsSs.getColumn(colIdx(c)).hidden === antesSs[i]),
    "sin pasarle Sumas y Saldos no la toca: la firma vieja sigue sirviendo");

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
