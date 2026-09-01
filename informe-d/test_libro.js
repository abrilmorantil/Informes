// ¿El libro que genera la app tiene la misma estructura que el prototipo?
const BASE = 'C:/Users/amoran/Downloads/motor_informes_sca (1)/motor_informes_sca/sitio';
const X = require(BASE + '/informe-a/vendor/xlsx.full.min.js');
global.XLSX = X;
const { parseSysBimonetario } = require(BASE + '/informe-d/parser_sys.js');
const motor = require(BASE + '/informe-d/motor_difcambio.js');
const imp = require(BASE + '/informe-d/importador_onvio.js');
const libro = require(BASE + '/informe-d/libro_calculo.js');
const fs = require('fs');

let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? '  OK  ' : ' FALLA'} ${m}`); if (!ok) fallos++; };
const PARAMS = { periodoFin: '2026-06-30', tcCompra: 1473, tcVenta: 1482, numeroAsiento: 1, concepto: 'Ajuste por Conversión' };
const PROTO = BASE + '/Prototipo_Ajuste_Dif_Cambio_06-2026.xlsx';

// Este test necesita dos archivos que NO estan en el repo: el SyS de prueba de junio y el
// prototipo contra el que se compara la estructura. Si falta alguno, no puede decir nada del
// motor, asi que lo dice y sale bien en vez de reventar con un ENOENT.
const SYS = BASE + '/SyS_prueba_06-2026.xls';
const faltan = [SYS, PROTO].filter(f => !fs.existsSync(f));
if (faltan.length) {
  console.log('(salteado) faltan los archivos con los que se compara:');
  faltan.forEach(f => console.log('   ' + f));
  process.exit(0);
}

// El libro que se genera va a una carpeta temporal, NO al repo.
const SALIDA = require('path').join(require('os').tmpdir(), 'Ajuste_Dif_Cambio_prueba.xlsx');

const wbs = X.read(fs.readFileSync(SYS), { type: 'buffer', raw: true });
const { cuentas } = parseSysBimonetario(X.utils.sheet_to_json(wbs.Sheets[wbs.SheetNames[0]], { header: 1, raw: true, defval: null }));
const cfg = motor.configDifCambio();
const { lineasOk, lineasARevisar } = motor.calcularConRevision(cuentas, PARAMS, cfg);
const asiento = motor.armarAsiento(lineasOk, cfg);

const datos = libro.escribirLibroCalculo({
  asiento, params: PARAMS, cfg, revisadas: lineasARevisar, revisadasExcluidas: [],
  titulo: 'Southern Copper Argentina S.R.L. — cierre 2026-06-30',
});
fs.writeFileSync(SALIDA, Buffer.from(datos));

const mio = X.read(fs.readFileSync(SALIDA), { type: 'buffer', cellNF: true, raw: true });
const pro = X.read(fs.readFileSync(PROTO), { type: 'buffer', cellNF: true, raw: true });

console.log('=== 1) las hojas ===');
console.log(`   prototipo: ${pro.SheetNames.join(' | ')}`);
console.log(`   generado : ${mio.SheetNames.join(' | ')}`);
check(JSON.stringify(mio.SheetNames) === JSON.stringify(pro.SheetNames), 'mismas hojas y en el mismo orden');

const filasDe = (wb, hoja) => X.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, defval: null });

console.log('\n=== 2) hoja Cálculo ===');
const cp = filasDe(pro, 'Cálculo'), cm = filasDe(mio, 'Cálculo');
check(JSON.stringify(cm[1]) === JSON.stringify(cp[1]), `mismos encabezados (${(cm[1] || []).length} columnas)`);
console.log(`   encabezados: ${(cm[1] || []).join(' | ')}`);
check(cm.length === cp.length, `misma cantidad de filas (${cm.length} vs ${cp.length})`);
// la fila de control
const ctrlP = cp[cp.length - 1], ctrlM = cm[cm.length - 1];
check(String(ctrlM[7]).startsWith('CONTROL') && ctrlM[9] === 'OK ✓',
  `la fila de control cierra: "${ctrlM[7]}" ${ctrlM[8]} ${ctrlM[9]}`);
check(String(ctrlP[7]).startsWith('CONTROL'), 'el prototipo también la tiene, en la misma columna');

console.log('\n=== 3) hoja Importador ===');
const ip = filasDe(pro, 'Importador'), im = filasDe(mio, 'Importador');
check(JSON.stringify(im[1]) === JSON.stringify(ip[1]), 'mismos encabezados que el prototipo');
check(JSON.stringify(im[1]) === JSON.stringify(imp.HEADERS_IMPORTADOR), 'y son los 12 exactos del modelo 0.xls');
check(im.length === ip.length, `misma cantidad de filas (${im.length} vs ${ip.length})`);
const ultM = im[im.length - 1];
check(String(ultM[5]).trim() === 'Σ =' && Math.abs(ultM[6]) < 0.005, `la fila Σ cierra en 0 (${ultM[6]})`);

console.log('\n=== 4) hoja Parámetros ===');
const pp = filasDe(pro, 'Parámetros'), pm = filasDe(mio, 'Parámetros');
const etiquetas = (f) => f.map(r => (r || [])[1]).filter(x => typeof x === 'string');
console.log(`   generado: ${etiquetas(pm).slice(2).join(' / ')}`);
check(pm[4][2] === 46203, `la fecha de cierre va como serial (${pm[4][2]})`);
check(pm[5][2] === 1473 && pm[6][2] === 1482, `los dos tipos de cambio (${pm[5][2]} / ${pm[6][2]})`);
check(pm[9][2] === cfg.cuentaBalanceo, `la cuenta de balanceo (${pm[9][2]})`);

console.log('\n=== 5) hoja Resumen y notas ===');
const rm = filasDe(mio, 'Resumen y notas');
const balanceo = asiento.find(l => l.esBalanceo);
check(rm[3][2] === asiento.length - 1, `líneas que publican (${rm[3][2]})`);
check(Math.abs(rm[4][2] - balanceo.ajusteUsd) < 0.005, `ajuste a diferencia de cambio (${rm[4][2]})`);
check(Math.abs(rm[5][2]) < 0.005, `control de cierre en 0 (${rm[5][2]})`);
check(rm[6][2] === 0, `cuentas marcadas para revisar (${rm[6][2]})`);

console.log('\n=== 6) el importador para Onvio sigue limpio ===');
const buf = imp.escribirImportador(asiento, PARAMS, 'xls');
const wbi = X.read(buf, { type: 'array', raw: true });
const fi = X.utils.sheet_to_json(wbi.Sheets['Asientos'], { header: 1, raw: true, defval: null });
check(JSON.stringify(fi[0]) === JSON.stringify(imp.HEADERS_IMPORTADOR), 'sus encabezados están en la PRIMERA fila, sin título');
check(fi.length - 1 === asiento.length, `y no lleva fila de control (${fi.length - 1} filas de datos)`);

console.log(`\n${fallos === 0 ? 'TODO OK' : fallos + ' FALLAS'}`);
process.exit(fallos ? 1 : 0);
