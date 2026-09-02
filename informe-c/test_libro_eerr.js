// ¿El EE RR generado tiene el mismo layout y los mismos números que el de junio?
const BASE = 'C:/Users/amoran/Downloads/motor_informes_sca (1)/motor_informes_sca/sitio';
const E = require(BASE + '/informe-a/vendor/exceljs.min.js'); global.ExcelJS = E;
const X = require(BASE + '/informe-a/vendor/xlsx.full.min.js'); global.XLSX = X;
const fuA = require(BASE + '/informe-a/formula_utils.js');
const eerr = require(BASE + '/informe-c/eerr.js');
const libro = require(BASE + '/informe-c/libro_eerr.js');
const fs = require('fs');
let fallos = 0;
const check = (ok, m) => { console.log(`${ok ? '  OK  ' : ' FALLA'} ${m}`); if (!ok) fallos++; };
const cerca = (a, b) => Math.abs(Number(a) - Number(b)) < 0.02;

// mes anterior segun el EE RR de junio (columna C)
const MAYO = { gastosOperacion: -559321.59, gastosAdministracion: -491629.55,
               ajusteTraduccion: 12131.03, otrosIngresos: 0, extraordinarios: 0,
               ingresosOperacion: 0, impuesto: 0 };

// Este test se compara contra el EE RR de junio hecho a mano, que no está en el repo: vive en
// la carpeta de descargas. Si no está, el test no puede decir nada, así que lo dice y sale
// bien, en vez de reventar con un ENOENT que parece que se rompió el motor.
const ORIGINAL = 'C:/Users/amoran/Downloads/EE RR 062026(adic).xls';

// El archivo que se genera para comparar va a una carpeta temporal, NO al repo: escribirlo
// acá dejaba un `EERR_generado.xlsx` suelto cada vez que se corría el test.
const SALIDA = require('path').join(require('os').tmpdir(), 'EERR_generado_test.xlsx');

(async () => {
  if (!fs.existsSync(ORIGINAL)) {
    console.log(`(salteado) falta el EE RR de junio contra el que se compara:\n   ${ORIGINAL}`);
    process.exit(0);
  }
  const wb = await fuA.abrirWorkbook(fs.readFileSync(BASE + '/informe-c/base_dolares.xlsx'));
  const s = hojaDistrib(wb);
  const cacheado = (f) => { const c = s.getCell(f, 3); return typeof c.result === 'number' ? c.result : (typeof c.value === 'number' ? c.value : 0); };
  const actual = eerr.totalesEstadoResultados(wb, cacheado);

  const datos = libro.escribirLibroEERR({ actual, anterior: MAYO, periodoFin: '2026-06-30' });
  fs.writeFileSync(SALIDA, Buffer.from(datos));

  const mio = X.read(fs.readFileSync(SALIDA), { type: 'buffer', raw: true });
  const pro = X.read(fs.readFileSync(ORIGINAL), { type: 'buffer', raw: true });
  const hm = mio.SheetNames[0], hp = pro.SheetNames.find(n => /RESULT/i.test(n));
  console.log(`hoja generada: "${hm}"   |   original: "${hp}"`);
  check(hm === 'RESULT Junio US$', 'el nombre de la hoja sigue el patrón del original');

  const fm = X.utils.sheet_to_json(mio.Sheets[hm], { header: 1, raw: true, defval: null });
  const fp = X.utils.sheet_to_json(pro.Sheets[hp], { header: 1, raw: true, defval: null });

  console.log('\n=== etiquetas de la columna B ===');
  const etq = (f) => f.map(r => (r || [])[1]).map(x => (typeof x === 'string' ? x.trim() : null));
  const em = etq(fm), ep = etq(fp);
  let ok = 0, dif = [];
  for (let i = 0; i < Math.max(em.length, ep.length); i++) {
    const a = em[i], b = ep[i];
    if (!a && !b) continue;
    if (a === b) { ok++; continue; }
    if (a && b && b.startsWith(a.slice(0, 18))) { ok++; continue; }
    dif.push(`fila ${i + 1}: generado ${JSON.stringify(a)} / original ${JSON.stringify(b)}`);
  }
  console.log(`   coinciden: ${ok}`);
  dif.forEach(d => console.log('   ' + d));
  check(dif.length === 0, `todas las etiquetas en su fila (${dif.length} distintas)`);

  console.log('\n=== los importes de la columna E (mes actual) ===');
  for (const [fila, nombre, esperado] of [
    [11, 'Gastos de Operación', -616736.47],
    [13, 'Gastos de Administración', -530007.27],
    [19, 'Ajuste por traducción', -46243.87],
  ]) {
    check(cerca(fm[fila - 1][4], esperado), `${nombre}: ${Number(fm[fila - 1][4]).toFixed(2)} (original ${esperado.toFixed(2)})`);
  }
  console.log('\n=== la columna C (mes anterior) ===');
  check(cerca(fm[10][2], MAYO.gastosOperacion), `gastos de operación de mayo: ${Number(fm[10][2]).toFixed(2)}`);
  check(cerca(fm[12][2], MAYO.gastosAdministracion), `gastos de administración de mayo: ${Number(fm[12][2]).toFixed(2)}`);

  console.log('\n=== los subtotales van como fórmula, igual que el original ===');
  // Se mira el libro ANTES de escribirlo: al releer, SheetJS no devuelve .f en celdas sin
  // valor cacheado, aunque la fórmula sí quede escrita en el XML del archivo.
  const libroObj = libro.construirLibroEERR({ actual, anterior: MAYO, periodoFin: '2026-06-30' });
  const ws = libroObj.Sheets[libroObj.SheetNames[0]];
  const wsP = pro.Sheets[hp];
  for (const dir of ['D11', 'E12', 'E15', 'E22', 'E26']) {
    const c = ws[dir], o = wsP[dir];
    check(!!(c && c.f), `${dir} es fórmula: ${c && c.f ? c.f : '(no)'}` +
      (o && o.f ? `   (original: ${o.f})` : ''));
  }
  console.log(`\n${fallos === 0 ? 'TODO OK' : fallos + ' FALLAS'}`);
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error('ERROR:', e.message, e.stack); process.exit(1); });
