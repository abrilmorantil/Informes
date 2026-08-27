// Las columnas de total por mes de "Dist.de gastos".
//
// El archivo tiene una columna por mes (TOTAL ENERO … TOTAL DICIEMBRE) más
// TOTAL AÑO. Los meses ya cerrados guardan un número fijo; el mes en curso es el
// único "vivo", con la fórmula =D<fila>, que sigue a la columna MOVIMIENTO MES.
//
// Por eso cargar un mes sin mover esto escribe el mes nuevo encima del total del
// mes anterior, y el TOTAL AÑO queda mal.

const NOMBRES_MES = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
];

function nombreMes(nro) {
  return NOMBRES_MES[nro - 1];
}

// "2026-06" -> {anio: 2026, mes: 6}
function parsearPeriodo(p) {
  const m = /^(\d{4})-(\d{1,2})$/.exec(String(p || "").trim());
  if (!m) throw new Error(`Período inválido: ${p}. Se espera algo como "2026-06".`);
  const mes = parseInt(m[2], 10);
  if (mes < 1 || mes > 12) throw new Error(`Mes inválido en el período: ${p}.`);
  return { anio: parseInt(m[1], 10), mes };
}

function formatearPeriodo({ anio, mes }) {
  return `${anio}-${String(mes).padStart(2, "0")}`;
}

function periodoSiguiente(p) {
  const { anio, mes } = parsearPeriodo(p);
  return mes === 12 ? formatearPeriodo({ anio: anio + 1, mes: 1 })
                    : formatearPeriodo({ anio, mes: mes + 1 });
}

function etiquetaPeriodo(p) {
  const { anio, mes } = parsearPeriodo(p);
  return `${nombreMes(mes)} ${anio}`;
}

// Las filas de categoría son las que suman su propia fila en la columna D
// (=SUM(F10:V10)+E10, el movimiento del mes).
//
// El rango tiene que ser de la MISMA fila: así queda afuera la fila de totales,
// que suma en vertical (=SUM(D8:D101)) y lleva su propia fórmula en las columnas
// de mes (SUBTOTAL), que no hay que pisar.
const SUM_HORIZONTAL = /SUM\([A-Z]{1,3}(\d+):[A-Z]{1,3}(\d+)\)/i;

function filasDeCategoria(wsDist) {
  const filas = [];
  for (let r = 8; r <= wsDist.rowCount; r++) {
    const v = wsDist.getCell(r, 4).value;
    if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
    const m = SUM_HORIZONTAL.exec(v.formula);
    if (m && parseInt(m[1], 10) === r && parseInt(m[2], 10) === r) filas.push(r);
  }
  return filas;
}

function columnaDeMes(wsDist, nroMes) {
  const buscado = nombreMes(nroMes);
  for (let c = 1; c <= wsDist.columnCount; c++) {
    const t6 = normTexto(wsDist.getCell(6, c).value);
    const t7 = normTexto(wsDist.getCell(7, c).value);
    if (t6.startsWith("TOTAL") && t7 === buscado) return c;
  }
  return null;
}

// Devuelve qué meses están "vivos" (apuntando a la columna D). En un archivo
// sano hay exactamente uno: el mes en curso.
function mesesVivos(wsDist) {
  const filas = filasDeCategoria(wsDist);
  const vivos = [];
  for (let nro = 1; nro <= 12; nro++) {
    const c = columnaDeMes(wsDist, nro);
    if (c === null) continue;
    let n = 0;
    for (const r of filas) {
      const v = wsDist.getCell(r, c).value;
      if (v && typeof v === "object" && typeof v.formula === "string" &&
          /^\s*\+?\s*D\d+\s*$/i.test(v.formula)) {
        n++;
      }
    }
    if (n > 0) vivos.push({ mes: nro, nombre: nombreMes(nro), celdas: n });
  }
  return vivos;
}

// Pone =D<fila> en la columna del mes, para que siga al movimiento del mes en curso.
function activarColumnaMes(wsDist, nroMes, log = () => {}) {
  const c = columnaDeMes(wsDist, nroMes);
  if (c === null) throw new Error(`No encontré la columna "TOTAL ${nombreMes(nroMes)}" en Dist.de gastos.`);
  const filas = filasDeCategoria(wsDist);
  for (const r of filas) wsDist.getCell(r, c).value = { formula: `D${r}` };

  // En el maestro los meses que todavía no se usaron vienen OCULTOS, y activarlos sólo
  // les escribía la fórmula: la columna quedaba con los importes correctos pero sin
  // verse, y había que mostrarla a mano en Excel cada mes. Pasó con julio —en el
  // informe presentado alguien la mostró— y volvería a pasar con agosto, que también
  // está oculto. El mes que se está cargando es justamente el que hay que mirar.
  const columna = wsDist.getColumn(c);
  if (columna.hidden) {
    columna.hidden = false;
    log(`  La columna TOTAL ${nombreMes(nroMes)} estaba oculta en el archivo: se muestra.`);
  }

  log(`  Columna TOTAL ${nombreMes(nroMes)} activada (${filas.length} filas siguen al movimiento del mes).`);
  return filas.length;
}

// Reemplaza la fórmula por el número que calculó Excel, para que el mes quede
// cerrado y no cambie cuando se cargue el mes siguiente.
function congelarColumnaMes(wsDist, nroMes, log = () => {}) {
  const c = columnaDeMes(wsDist, nroMes);
  if (c === null) throw new Error(`No encontré la columna "TOTAL ${nombreMes(nroMes)}" en Dist.de gastos.`);
  const filas = filasDeCategoria(wsDist);

  // ExcelJS no devuelve el resultado cuando la celda vale cero, así que no se puede
  // distinguir "da cero" de "nunca se calculó" mirando celda por celda. Lo que sí
  // distingue es la columna entera: si Excel la calculó alguna vez, al menos una
  // categoría tiene un importe distinto de cero. Si no hay ninguno, el archivo nunca
  // pasó por Excel y congelarlo dejaría todo en cero sin que se note.
  const valores = [];
  let algunoCalculado = false;
  for (const r of filas) {
    const v = wsDist.getCell(r, c).value;
    if (v && typeof v === "object" && v.formula !== undefined) {
      const n = typeof v.result === "number" ? v.result : 0;
      if (n !== 0) algunoCalculado = true;
      valores.push([r, n]);
    } else if (typeof v === "number") {
      if (v !== 0) algunoCalculado = true;
      valores.push([r, v]);           // ya estaba congelado
    } else {
      valores.push([r, 0]);
    }
  }

  if (!algunoCalculado) {
    throw new Error(
      `No puedo cerrar ${nombreMes(nroMes)}: la columna está entera en cero, lo que pasa ` +
      `cuando el archivo no se abrió y guardó en Excel (Excel es el que calcula los ` +
      `importes). Abrí el borrador en Excel, guardalo, y volvé a subirlo. NO se guardó nada.`
    );
  }

  for (const [r, valor] of valores) wsDist.getCell(r, c).value = valor;
  log(`  Columna TOTAL ${nombreMes(nroMes)} congelada (${valores.length} valores fijos).`);
  return valores.length;
}

function normTexto(s) {
  if (s === null || s === undefined) return "";
  if (typeof s === "object") {
    if (s.richText) return normTexto(s.richText.map(t => t.text).join(""));
    if (s.result !== undefined) return normTexto(s.result);
    return "";
  }
  return String(s).replace(/\s+/g, " ").trim().toUpperCase();
}

if (typeof module !== "undefined") {
  module.exports = {
    NOMBRES_MES, nombreMes, parsearPeriodo, formatearPeriodo, periodoSiguiente,
    etiquetaPeriodo, filasDeCategoria, columnaDeMes, mesesVivos,
    activarColumnaMes, congelarColumnaMes,
  };
}
