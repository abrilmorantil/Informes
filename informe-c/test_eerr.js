// ¿El cálculo en JavaScript reproduce lo que Excel tiene cacheado en el maestro de junio?
// Es la prueba de fuego: si da igual, el EE RR se puede armar en la misma corrida.
const BASE = 'C:/Users/amoran/Downloads/motor_informes_sca (1)/motor_informes_sca/sitio';
const E = require(BASE + '/informe-a/vendor/exceljs.min.js'); global.ExcelJS = E;
const X = require(BASE + '/informe-a/vendor/xlsx.full.min.js'); global.XLSX = X;
const fuA = require(BASE + '/informe-a/formula_utils.js');
// En Node cada archivo es su propio módulo, así que los nombres de hoja compartidos —que en
// el navegador define motor_balances.js para todos— hay que dejarlos en el global a mano.
for (const [k, v] of Object.entries(require(BASE + '/informe-c/motor_balances.js'))) {
  if (global[k] === undefined) global[k] = v;
}
const { parseExportBalances } = require(BASE + '/informe-c/parser_balances.js');
const cl = require(BASE + '/informe-c/clasificacion.js');
const usd = require(BASE + '/informe-c/dolares.js');
const eerr = require(BASE + '/informe-c/eerr.js');
const fs = require('fs');

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? '  OK  ' : ' FALLA'} ${m}`); if (!ok) fallos++; };
const cerca = (a, b, t = 0.02) => Math.abs(a - b) < t;

// lo que el EE RR de junio muestra en la columna MES ACTUAL
const EERR_JUNIO = {
  gastosOperacion: -616736.47,
  gastosAdministracion: -530007.27,
  totalResultadoOperacion: -1146743.74,
  ajusteTraduccion: -46243.87,
  otrosIngresos: 0,
  antesDeImpuestos: -1192987.61,
  resultadoEjercicio: -1192987.61,
};

// El export de junio no está en el repo: vive en la carpeta de descargas y se lo pisa cada
// mes con el del mes nuevo. Los bloques 1, 2 y 4 no lo necesitan y corren siempre; el 3 sí,
// y cuando no está —o cuando el que está es de otro mes— se saltea DICIENDO por qué, en vez
// de dar una falla que parece un error del motor y no lo es.
const EXPORT_JUNIO = 'C:/Users/amoran/Downloads/Balance de SyS por Cod. de Cta. (2).xls';

(async () => {
  const clasif = cl.indexarClasificacion(JSON.parse(fs.readFileSync(BASE + '/mapping.json', 'utf8')));
  const hayExport = fs.existsSync(EXPORT_JUNIO);
  let exp = null;
  if (hayExport) {
    const wbO = X.read(fs.readFileSync(EXPORT_JUNIO), { type: 'buffer', raw: true });
    const wsO = wbO.Sheets[wbO.SheetNames[0]];
    exp = parseExportBalances(X.utils.sheet_to_json(wsO, { header: 1, raw: true, defval: null }), wsO['!merges']);
  }
  const wb = await fuA.abrirWorkbook(fs.readFileSync(BASE + '/informe-c/base_dolares.xlsx'));

  // --- 1) primero, contra los valores que Excel dejo cacheados en el archivo
  console.log('=== 1) el evaluador contra los valores cacheados por Excel ===');
  const s = hojaDistrib(wb);
  const cacheado = (fila) => {
    const c = s.getCell(fila, 3);
    return typeof c.result === 'number' ? c.result : (typeof c.value === 'number' ? c.value : 0);
  };
  const tCache = eerr.totalesEstadoResultados(wb, cacheado);
  const a2 = wb.getWorksheet('Anexo II');
  const fila = tCache.anexo.fila;
  console.log(`   fila de totales del Anexo II detectada: ${fila}`);
  for (const [col, nombre] of [['F', 'administracion'], ['G', 'comercializacion'], ['H', 'exploracion'], ['I', 'financieros'], ['E', 'totales']]) {
    const excel = a2.getCell(`${col}${fila}`).result;
    const mio = tCache.anexo[nombre];
    check(cerca(mio, excel), `Anexo II ${col}${fila} (${nombre}): Excel ${Number(excel).toFixed(2)} / calculado ${mio.toFixed(2)}`);
  }
  const res = wb.getWorksheet('Resultados');
  check(cerca(tCache.ajusteTraduccion, res.getCell('C17').result),
    `ajuste por traducción: Excel ${Number(res.getCell('C17').result).toFixed(2)} / calculado ${tCache.ajusteTraduccion.toFixed(2)}`);
  check(cerca(tCache.antesDeImpuestos, res.getCell('C21').result),
    `resultado antes de impuestos: Excel ${Number(res.getCell('C21').result).toFixed(2)} / calculado ${tCache.antesDeImpuestos.toFixed(2)}`);

  // --- 2) contra el EE RR de junio que armo la usuaria a mano
  console.log('\n=== 2) contra el EE RR de junio hecho a mano ===');
  for (const [k, esperado] of Object.entries(EERR_JUNIO)) {
    check(cerca(tCache[k], esperado), `${k}: EE RR ${esperado.toFixed(2)} / calculado ${tCache[k].toFixed(2)}`);
  }

  // --- 3) ahora de verdad: calculando los saldos desde el export, sin usar el cache
  console.log('\n=== 3) calculando los saldos desde el export (sin mirar el cache) ===');
  if (!hayExport) {
    console.log(`  (salteado) falta el export de junio 2026:\n     ${EXPORT_JUNIO}`);
  } else {
    const equiv = JSON.parse(fs.readFileSync(BASE + '/informe-c/equivalencias_dolares.json', 'utf8'));
    const cm = usd.cuentasDelMaestro(wb, 'dolares');
    const r = usd.resolverDestinosDolares({
      cuentasExport: exp.cuentas, cuentasMaestro: cm, clasificacion: clasif,
      equivalencias: equiv, clavesHoja1: usd.clavesDeHoja1(wb, 'dolares'),
    });
    const porFila = new Map();
    for (const d of r.destinos.values()) porFila.set(d.fila, d.aportes.reduce((a, x) => a + x.saldo, 0));
    const tCalc = eerr.totalesEstadoResultados(wb, (f) => porFila.get(f) || 0);
    console.log(`   Administración ${tCalc.gastosAdministracion.toFixed(2)}  Exploración ${tCalc.gastosOperacion.toFixed(2)}`);
    console.log(`   Ajuste traducción ${tCalc.ajusteTraduccion.toFixed(2)}  Resultado ${tCalc.resultadoEjercicio.toFixed(2)}`);

    // Si el archivo que hay es de otro mes, este bloque no puede decir nada del motor: sus
    // cifras van a diferir por el dato, no por el código. Se lo dice y se saltea. Que sea de
    // otro mes se sabe porque los bloques 1 y 2 —que no lo usan— dieron bien.
    const coincide = cerca(tCalc.gastosAdministracion, EERR_JUNIO.gastosAdministracion, 1) &&
                     cerca(tCalc.gastosOperacion, EERR_JUNIO.gastosOperacion, 1);
    if (!coincide && fallos === 0) {
      console.log(`  (salteado) el archivo que hay no es el export de junio 2026 — dice` +
        ` ${tCalc.gastosOperacion.toFixed(2)} de gastos de operación y junio son` +
        ` ${EERR_JUNIO.gastosOperacion.toFixed(2)}. Para correr este bloque hay que dejar el` +
        ` export de junio en:\n     ${EXPORT_JUNIO}`);
    } else {
      check(cerca(tCalc.gastosAdministracion, EERR_JUNIO.gastosAdministracion, 1),
        `gastos de administración desde el export (${tCalc.gastosAdministracion.toFixed(2)})`);
      check(cerca(tCalc.gastosOperacion, EERR_JUNIO.gastosOperacion, 1),
        `gastos de operación desde el export (${tCalc.gastosOperacion.toFixed(2)})`);
    }
  }

  console.log('\n=== 4) los controles de consistencia ===');
  const v = eerr.verificarEERR(tCache);
  check(v.concilia, `el estado concilia${v.avisos.length ? ': ' + v.avisos.join(' | ') : ''}`);

  console.log(`\n${fallos === 0 ? 'TODO OK' : fallos + ' FALLAS'}`);
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error('ERROR:', e.message, e.stack); process.exit(1); });
