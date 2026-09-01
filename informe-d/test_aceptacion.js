// Test de aceptación de la sección 9 de la especificación.
// Golden file: Grilla del asiento real Nº 2630 (06-2026). TC compra 1473, venta 1482.
const BASE = 'C:/Users/amoran/Downloads/motor_informes_sca (1)/motor_informes_sca/sitio';
const X = require(BASE + '/informe-a/vendor/xlsx.full.min.js');
global.XLSX = X;
const { parseSysBimonetario, revisarParseoSys } = require(BASE + '/informe-d/parser_sys.js');
const motor = require(BASE + '/informe-d/motor_difcambio.js');
const imp = require(BASE + '/informe-d/importador_onvio.js');
const fs = require('fs');

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? '  OK  ' : ' FALLA'} ${m}`); if (!ok) fallos++; };

const GOLDEN = 'C:/Users/amoran/Downloads/Grilla Listado de Asientos Detallado por Nro. de Asiento.xlsx';
const SYS = BASE + '/SyS_prueba_06-2026.xls';

// Ni el golden (la grilla del asiento real) ni el SyS de prueba estan en el repo. Sin
// ellos el test no puede decir nada, asi que lo dice y sale bien en vez de reventar.
const faltan = [GOLDEN, SYS].filter(f => !fs.existsSync(f));
if (faltan.length) {
  console.log('(salteado) faltan los archivos con los que se compara:');
  faltan.forEach(f => console.log('   ' + f));
  process.exit(0);
}

// El importador de prueba va a una carpeta temporal, NO al repo.
const SALIDA_IMP = require('path').join(require('os').tmpdir(), 'importador_prueba.xls');
const PARAMS = { periodoFin: '2026-06-30', tcCompra: 1473, tcVenta: 1482, numeroAsiento: 1 };

// --- golden: {codigo: round(DebeMEP - HaberMEP, 2)}  (columnas K=10 y L=11)
function leerGolden() {
  const wb = X.read(fs.readFileSync(GOLDEN), { type: 'buffer', raw: true });
  const filas = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  const out = new Map();
  for (const f of filas.slice(1)) {
    if (!f || !f[3]) continue;
    const cod = String(f[3]).trim();
    if (!/^\d{9}$/.test(cod)) continue;
    const debe = typeof f[10] === 'number' ? f[10] : 0;
    const haber = typeof f[11] === 'number' ? f[11] : 0;
    out.set(cod, motor.redondear2(debe - haber));
  }
  return out;
}

(async () => {
  console.log('=== 1) parseo del SyS ===');
  const wb = X.read(fs.readFileSync(SYS), { type: 'buffer', raw: true });
  const filas = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  const { cuentas, secciones, columnas } = parseSysBimonetario(filas);
  console.log(`   ${cuentas.length} cuentas | secciones: ${secciones.join(', ')} | columnas ${JSON.stringify(columnas)}`);
  const avisos = revisarParseoSys(cuentas);
  check(avisos.length === 0, `el parseo no levanta avisos${avisos.length ? ': ' + avisos.join(' | ') : ''}`);
  check(secciones.length === 4, 'reconoce las cuatro secciones');

  console.log('\n=== 2) cálculo con revisión (fase 1) ===');
  const cfg = motor.configDifCambio();
  const { lineasOk, lineasARevisar } = motor.calcularConRevision(cuentas, PARAMS, cfg);
  console.log(`   líneas OK: ${lineasOk.length} | a revisar: ${lineasARevisar.length}`);
  lineasARevisar.forEach(l => console.log(`      ${l.codigo} ${l.denominacion.slice(0, 30)} ajuste ${l.ajusteUsd.toFixed(2)} (${l.motivo})`));
  check(lineasARevisar.length === 0, 'con la config real no marca ninguna línea para revisar');

  console.log('\n=== 3) el asiento ===');
  const asiento = motor.armarAsiento(lineasOk, cfg);
  const { cierra, suma } = motor.verificarCierre(asiento);
  check(cierra, `el asiento suma 0 (suma = ${suma})`);

  const balanceo = asiento.find(l => l.codigo === cfg.cuentaBalanceo);
  check(!!balanceo, 'tiene la línea de balanceo 423050000');
  console.log(`   423050000 = ${balanceo.ajusteUsd.toFixed(2)} USD`);
  check(Math.abs(balanceo.ajusteUsd - 58374.79) <= 0.02 || Math.abs(balanceo.ajusteUsd - 58374.80) <= 0.02,
    `la cuenta de balanceo da ~58.374,79/80 (dio ${balanceo.ajusteUsd.toFixed(2)})`);

  console.log('\n=== 4) contra el asiento real Nº 2630 ===');
  const golden = leerGolden();
  console.log(`   golden: ${golden.size} cuentas`);
  const generado = new Map(asiento.map(l => [l.codigo, l.ajusteUsd]));
  // La cuenta de balanceo se controla aparte (aserción 3 de la spec): el manual difiere
  // en 0,01 USD porque el operador incluyo a mano una linea sub-centavo (Toro Franco,
  // 0,0026 USD). Esta documentado y es inmaterial.
  const materiales = [...golden.entries()]
    .filter(([c, v]) => Math.abs(v) >= 0.01 && c !== cfg.cuentaBalanceo);
  let coinciden = 0; const difieren = [], faltan = [];
  for (const [cod, esperado] of materiales) {
    if (!generado.has(cod)) { faltan.push(`${cod} (esperaba ${esperado.toFixed(2)})`); continue; }
    const obtenido = generado.get(cod);
    if (Math.abs(obtenido - esperado) <= 0.01) coinciden++;
    else difieren.push(`${cod}: esperaba ${esperado.toFixed(2)}, dio ${obtenido.toFixed(2)}`);
  }
  console.log(`   materiales en el golden (sin la de balanceo): ${materiales.length}`);
  console.log(`   coinciden al centavo   : ${coinciden}`);
  if (faltan.length) { console.log('   NO generadas:'); faltan.slice(0, 10).forEach(x => console.log('      ' + x)); }
  if (difieren.length) { console.log('   con diferencia:'); difieren.slice(0, 10).forEach(x => console.log('      ' + x)); }
  check(faltan.length === 0, `no falta ninguna línea del golden (${faltan.length})`);
  check(difieren.length === 0, `ninguna difiere en más de 0,01 (${difieren.length})`);
  // se compara en centavos enteros: en coma flotante 58374.80 - 58374.79 da 0.0100000000002
  const goldBal = golden.get(cfg.cuentaBalanceo);
  const centavos = Math.abs(Math.round(balanceo.ajusteUsd * 100) - Math.round(goldBal * 100));
  check(centavos <= 1,
    `la de balanceo queda dentro del centavo documentado: manual ${goldBal.toFixed(2)}, motor ${balanceo.ajusteUsd.toFixed(2)} (${centavos} centavo de diferencia, el de la linea sub-centavo que el operador sumo a mano)`);

  console.log('\n=== 5) las no monetarias quedan afuera ===');
  const PROHIBIDAS = ['124020001', '124020006', '125010002', '114010016', '114050005', '211050000'];
  const coladas = PROHIBIDAS.filter(c => generado.has(c));
  check(coladas.length === 0, `ninguna cuenta no monetaria entró al asiento (${coladas.join(', ') || 'ninguna'})`);

  console.log('\n=== 6) la salvaguarda detecta el olvido de clasificar ===');
  const sinExclusiones = motor.configDifCambio({ prefijosNoMonetarios: [], noMonetariasExactas: [] });
  const r2 = motor.calcularConRevision(cuentas, PARAMS, sinExclusiones);
  const marcadas = r2.lineasARevisar.map(l => l.codigo);
  console.log(`   sin exclusiones marca ${marcadas.length}: ${marcadas.join(', ')}`);
  const capturadas = PROHIBIDAS.filter(c => marcadas.indexOf(c) >= 0);
  // 114050005 (Seguros a Devengar) NO dispara la salvaguarda: su ratio es 0,06, igual al
  // de una monetaria. La spec lo dice expresamente; por eso sigue en la lista mantenida.
  check(capturadas.length === 5 && capturadas.indexOf('114050005') < 0,
    `captura las 5 no monetarias evidentes y deja fuera Seguros a Devengar, como dice la spec (${capturadas.join(', ')})`);
  const falsosPositivos = marcadas.filter(c => PROHIBIDAS.indexOf(c) < 0);
  check(falsosPositivos.length === 0, `y no marca ninguna monetaria legítima (${falsosPositivos.join(', ') || 'ninguna'})`);

  console.log('\n=== 7) el importador ===');
  const buf = imp.escribirImportador(asiento, PARAMS, 'xls');
  fs.writeFileSync(SALIDA_IMP, Buffer.from(buf));
  const wbi = X.read(fs.readFileSync(SALIDA_IMP), { type: 'buffer', cellNF: true });
  const wsi = wbi.Sheets['Asientos'];
  check(!!wsi, 'genera la hoja "Asientos"');
  const fi = X.utils.sheet_to_json(wsi, { header: 1, raw: true, defval: null });
  check(JSON.stringify(fi[0]) === JSON.stringify(imp.HEADERS_IMPORTADOR), 'los encabezados son los exactos del modelo');
  check(fi.length - 1 === asiento.length, `una fila por línea del asiento (${fi.length - 1} vs ${asiento.length})`);
  check(fi[1][2] === 46203, `la fecha va como serial de Excel (${fi[1][2]})`);
  check(fi[1][5] === 0, 'el importe en moneda local es 0');
  const sumaG = fi.slice(1).reduce((a, f) => a + (typeof f[6] === 'number' ? f[6] : 0), 0);
  check(Math.abs(motor.redondear2(sumaG)) < 0.005, `la columna G del importador suma 0 (${motor.redondear2(sumaG)})`);
  check(String(fi[1][4]).length === 9, `el código de cuenta va completo de 9 dígitos (${fi[1][4]})`);

  console.log(`\n${fallos === 0 ? 'TODO OK' : fallos + ' FALLAS'}`);
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error('ERROR:', e.message, e.stack); process.exit(1); });
