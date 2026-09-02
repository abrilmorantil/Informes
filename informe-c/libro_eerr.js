// Arma el archivo del Estado de Resultados con el mismo layout que
// "EE RR <periodo>(adic).xls", hoja `RESULT <Mes> US$`.
//
// Las columnas C (mes anterior) y E (mes actual) llevan importes; D (movimiento del mes)
// y los subtotales van como FÓRMULA, igual que en el original, para que el archivo siga
// siendo un papel de trabajo vivo y se vea de dónde sale cada número.

const MESES_EERR = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

const FMT_EERR = '#,##0.00;(#,##0.00);\\-';

function _Xe() {
  return typeof XLSX !== "undefined" ? XLSX : require("../informe-a/vendor/xlsx.full.min.js");
}

// "2026-06-30" -> {anio, mes, dia, nombreMes}
function partesPeriodo(periodoFin) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(periodoFin));
  if (!m) throw new Error(`Período inválido: ${periodoFin}. Se espera 2026-06-30.`);
  const anio = +m[1], mes = +m[2], dia = +m[3];
  return { anio, mes, dia, nombreMes: MESES_EERR[mes - 1] };
}

function nombreHojaEERR(periodoFin) {
  const p = partesPeriodo(periodoFin);
  const mes = p.nombreMes.charAt(0).toUpperCase() + p.nombreMes.slice(1);
  return `RESULT ${mes} US$`;
}

// `actual` y `anterior` son objetos de totalesEstadoResultados (anterior puede ser null).
function construirLibroEERR({ actual, anterior, periodoFin, titulo, movimiento }) {
  const X = _Xe();
  const p = partesPeriodo(periodoFin);
  const ws = {};
  let maxC = 4, maxR = 0;
  const set = (col, fila, celda) => {
    if (celda === null || celda === undefined) return;
    ws[X.utils.encode_cell({ c: col, r: fila })] = celda;
    if (col > maxC) maxC = col;
    if (fila > maxR) maxR = fila;
  };
  const txt = (v) => ({ t: "s", v: String(v) });
  const num = (v) => ({ t: "n", v: Number(v) || 0, z: FMT_EERR });
  const frm = (f) => ({ t: "n", f, z: FMT_EERR });
  const A = 0, B = 1, C = 2, D = 3, E = 4;

  // encabezado (filas 1 a 4, combinadas B:E como en el original)
  set(B, 0, txt(titulo || "SOUTHERN COPPER ARGENTINA S.R.L."));
  set(B, 1, txt("ESTADO DE RESULTADOS"));
  // El período es el mes que se está emitiendo, del 1 al último día: "del 01 al 30 de junio
  // de 2026". Decía "del 01 de mayo al 30 de junio", que es otro período.
  set(B, 2, txt(`Período del 01 al ${p.dia} de ${p.nombreMes} de ${p.anio}`));
  set(B, 3, txt("En Dólares"));

  set(C, 6, txt("MES ANTERIOR"));
  set(D, 6, txt("MOVIMIENTO"));
  set(E, 6, txt("MES ACTUAL"));
  set(B, 7, txt("RESULTADO OPERACIONAL"));
  set(D, 7, txt("DEL MES"));

  const ant = anterior || {};
  const v = (o, k) => (o && typeof o[k] === "number" ? o[k] : 0);
  // fila (0-based) de cada concepto, para poder escribir las fórmulas
  const F = { ingresos: 9, gastosOp: 10, subtotal: 11, admin: 12, extra: 13, totalOp: 14,
              correccion: 16, difCambio: 17, ajuste: 18, otros: 19, antesImp: 21,
              impuesto: 23, ejercicio: 25 };
  const xl = (fila) => fila + 1;   // a numeración de Excel

  set(A, F.ingresos, num(67)); set(B, F.ingresos, txt("Ingresos de Operación"));
  if (v(ant, "ingresosOperacion") || v(actual, "ingresosOperacion")) {
    set(C, F.ingresos, num(v(ant, "ingresosOperacion")));
    set(D, F.ingresos, frm(`+E${xl(F.ingresos)}-C${xl(F.ingresos)}`));
    set(E, F.ingresos, num(v(actual, "ingresosOperacion")));
  }

  set(A, F.gastosOp, num(68)); set(B, F.gastosOp, txt("Gastos de Operación"));
  set(C, F.gastosOp, num(v(ant, "gastosOperacion")));
  set(D, F.gastosOp, frm(`+E${xl(F.gastosOp)}-C${xl(F.gastosOp)}`));
  set(E, F.gastosOp, num(v(actual, "gastosOperacion")));

  set(B, F.subtotal, txt("Sub total Resultado de Operación"));
  set(C, F.subtotal, frm(`+C${xl(F.gastosOp)}`));
  set(D, F.subtotal, frm(`+D${xl(F.gastosOp)}`));
  set(E, F.subtotal, frm(`+E${xl(F.gastosOp)}`));

  set(A, F.admin, num(69)); set(B, F.admin, txt("Gastos de Administración "));
  set(C, F.admin, num(v(ant, "gastosAdministracion")));
  set(D, F.admin, frm(`+E${xl(F.admin)}-C${xl(F.admin)}`));
  set(E, F.admin, num(v(actual, "gastosAdministracion")));

  set(B, F.extra, txt("Resultados Extraordinarios (Robo)"));
  if (v(ant, "extraordinarios") || v(actual, "extraordinarios")) {
    set(C, F.extra, num(v(ant, "extraordinarios")));
    set(D, F.extra, frm(`+E${xl(F.extra)}-C${xl(F.extra)}`));
    set(E, F.extra, num(v(actual, "extraordinarios")));
  }

  set(B, F.totalOp, txt("Total Resultado de operación"));
  ["C", "D", "E"].forEach((col, i) => set(C + i, F.totalOp, frm(`SUM(${col}${xl(F.subtotal)}:${col}${xl(F.extra)})`)));

  set(A, F.correccion, num(71)); set(B, F.correccion, txt("Correción Monetaria"));
  set(A, F.difCambio, num(72)); set(B, F.difCambio, txt("Diferencia de Cambio"));

  set(B, F.ajuste, txt("Ajuste por traducción"));
  set(C, F.ajuste, num(v(ant, "ajusteTraduccion")));
  set(D, F.ajuste, frm(`+E${xl(F.ajuste)}-C${xl(F.ajuste)}`));
  set(E, F.ajuste, num(v(actual, "ajusteTraduccion")));

  set(B, F.otros, txt("Otros ingresos y egresos"));
  set(C, F.otros, num(v(ant, "otrosIngresos")));
  set(D, F.otros, frm(`+E${xl(F.otros)}-C${xl(F.otros)}`));
  set(E, F.otros, num(v(actual, "otrosIngresos")));

  set(B, F.antesImp, txt("Resultado antes de Impuestos"));
  ["C", "D", "E"].forEach((col, i) => set(C + i, F.antesImp, frm(`SUM(${col}${xl(F.totalOp)}:${col}${xl(F.otros)})`)));

  // El impuesto se muestra NEGATIVO: la hoja Resultados lo resta (C25 = C21 - C23) y acá
  // en cambio la columna se suma, así que hay que invertirle el signo para que cierre.
  set(A, F.impuesto, num(79)); set(B, F.impuesto, txt("Impuesto a las Ganancias"));
  if (v(ant, "impuesto") || v(actual, "impuesto")) {
    set(C, F.impuesto, num(-v(ant, "impuesto")));
    set(D, F.impuesto, frm(`+E${xl(F.impuesto)}-C${xl(F.impuesto)}`));
    set(E, F.impuesto, num(-v(actual, "impuesto")));
  }

  set(B, F.ejercicio, txt(`Resultado del Ejercicio al ${p.dia} de ${p.nombreMes} ${p.anio}`));
  ["C", "D", "E"].forEach((col, i) => set(C + i, F.ejercicio, frm(`SUM(${col}${xl(F.antesImp)}:${col}${xl(F.impuesto)})`)));

  ws["!ref"] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: maxC, r: maxR } });
  ws["!cols"] = [7.5, 46.67, 14.83, 14.83, 14.83].map(w => ({ wch: w }));
  ws["!merges"] = [0, 1, 2, 3].map(r => ({ s: { c: B, r }, e: { c: E, r } }));

  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws, nombreHojaEERR(periodoFin));
  if (movimiento && movimiento.length) {
    X.utils.book_append_sheet(wb, hojaMovimiento(movimiento), nombreHojaMovimiento(periodoFin));
  }
  return wb;
}

// "MOVI 07-2026": el balance de comprobación en dólares que acompaña al estado de resultados.
//
// Es la hoja que Abril venía armando a mano. Una fila por cuenta del balance, con el saldo del
// mes pasado, lo que se movió y cómo queda; agrupadas por capítulo y con un TOTAL que tiene que
// cerrar en cero por las tres columnas.
//
// Van TODAS las cuentas del maestro, muevan o no. Es un balance de comprobación: una cuenta que
// no figura no se distingue de una que quedó en cero por error — y en agosto se perdieron
// 1.915.260,00 justamente porque dos cuentas nuevas no aparecían en ningún lado.
function nombreHojaMovimiento(periodoFin) {
  const p = partesPeriodo(periodoFin);
  return `MOVI ${String(p.mes).padStart(2, "0")}-${p.anio}`;
}

const ORDEN_CAPITULOS = ["ACTIVO", "PASIVO", "PATRIMONIO NETO", "RESULTADOS"];

function hojaMovimiento(movimiento) {
  const X = _Xe();
  const ws = {};
  let maxR = 0;
  const set = (c, r, celda) => {
    if (celda === null || celda === undefined) return;
    ws[X.utils.encode_cell({ c, r })] = celda;
    if (r > maxR) maxR = r;
  };
  const txt = (v, bold) => ({ t: "s", v, ...(bold ? { s: { font: { bold: true } } } : {}) });
  const num = (v) => ({ t: "n", v: Math.round(v * 100) / 100, z: FMT_EERR });

  set(0, 0, txt("Cuenta Contable", true));
  set(1, 0, txt("Saldo anterior", true));
  set(2, 0, txt("Movimiento del Mes Dolares", true));
  set(3, 0, txt("Saldo Final", true));

  // por capítulo, y dentro de cada uno en el orden en que están en el balance
  const porCap = new Map();
  for (const m of movimiento) {
    const k = m.capitulo || "OTRAS";
    if (!porCap.has(k)) porCap.set(k, []);
    porCap.get(k).push(m);
  }
  const capitulos = [...ORDEN_CAPITULOS.filter(c => porCap.has(c)),
                     ...[...porCap.keys()].filter(c => ORDEN_CAPITULOS.indexOf(c) < 0)];

  let r = 1;
  const filasDeDatos = [];
  for (const cap of capitulos) {
    set(0, r, txt(cap, true));
    r++;
    for (const m of porCap.get(cap).sort((a, b) => a.fila - b.fila)) {
      set(0, r, txt(`${m.codigo} - ${m.nombre}`));
      set(1, r, num(m.anterior));
      set(2, r, num(m.movimiento));
      set(3, r, num(m.final));
      filasDeDatos.push(r);
      r++;
    }
  }

  // El total suma SÓLO las filas de cuenta, salteando los títulos de capítulo: por eso va como
  // una lista de sumandos y no como un rango. Un rango que abarcara los títulos sumaría texto.
  r++;
  set(0, r, txt("TOTAL", true));
  for (const [c, letra] of [[1, "B"], [2, "C"], [3, "D"]]) {
    set(c, r, {
      t: "n", z: FMT_EERR,
      f: filasDeDatos.map(x => `${letra}${x + 1}`).join("+"),
      v: Math.round(movimiento.reduce((a, m) => a + (c === 1 ? m.anterior : c === 2 ? m.movimiento : m.final), 0) * 100) / 100,
    });
  }

  ws["!ref"] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 3, r: maxR } });
  ws["!cols"] = [52, 16, 26, 16].map(w => ({ wch: w }));
  ws["!autofilter"] = { ref: X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 3, r: maxR } }) };
  return ws;
}

function escribirLibroEERR(datos) {
  const X = _Xe();
  return X.write(construirLibroEERR(datos), { bookType: "xlsx", type: "array" });
}

// Lo que se guarda por corrida para que el mes siguiente tenga su "MES ANTERIOR".
function snapshotEERR(totales, periodoFin) {
  return {
    periodo: periodoFin,
    ingresosOperacion: totales.ingresosOperacion,
    gastosOperacion: totales.gastosOperacion,
    gastosAdministracion: totales.gastosAdministracion,
    extraordinarios: totales.extraordinarios,
    ajusteTraduccion: totales.ajusteTraduccion,
    otrosIngresos: totales.otrosIngresos,
    impuesto: totales.impuesto,
    antesDeImpuestos: totales.antesDeImpuestos,
    resultadoEjercicio: totales.resultadoEjercicio,
  };
}

if (typeof module !== "undefined") {
  module.exports = {
    MESES_EERR, partesPeriodo, nombreHojaEERR, nombreHojaMovimiento, hojaMovimiento,
    construirLibroEERR, escribirLibroEERR, snapshotEERR,
  };
}
