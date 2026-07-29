// Lee los saldos de cierre desde un BALCOMPROBDOLARES ya terminado, para sembrar la
// memoria del informe sin tener que cargar la columna a mano.
//
// Las columnas se ubican por su ENCABEZADO, no por posición: así sirve tanto para el
// archivo que genera la app como para el que se venía armando a mano, que puede tener
// las columnas corridas.
//
// El saldo final se toma del valor que Excel dejó calculado. Si el archivo nunca pasó por
// Excel esa celda viene vacía, así que se reconstruye como saldo anterior + (debe − haber),
// que es exactamente la fórmula del archivo.

const RE_CUENTA_B = /^\s*(\d{6,})\s*-\s*(.+?)\s*$/;

function _textoB(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined) return _textoB(v.result);
    return "";
  }
  return String(v);
}

function _numeroB(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/\s/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", "."));
    if (isFinite(n)) return n;
  }
  return null;
}

const _norm = (t) => _textoB(t).normalize("NFD").replace(/[̀-ͯ]/g, "")
  .toUpperCase().replace(/\s+/g, " ").trim();

const MESES_B = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
                 "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// El informe hecho a mano no rotula las columnas "Saldo anterior" y "Saldo Final": las
// llama por su fecha ("Saldo al 30 de Abril de 2026"). Se les saca la fecha para poder
// ordenarlas — la más nueva es el saldo final, la más vieja el anterior.
function _fechaDeEncabezado(t) {
  let m = /^SALDO AL (\d{1,2}) DE ([A-Z]+)(?: DE)? (\d{4})/.exec(t);
  if (m) {
    let mes = MESES_B.indexOf(m[2]);
    if (mes < 0 && m[2] === "SETIEMBRE") mes = 8;
    if (mes >= 0) return { orden: +m[3] * 10000 + (mes + 1) * 100 + +m[1], texto: t };
  }
  m = /^SALDO AL (\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(t);
  if (m) return { orden: +m[3] * 10000 + +m[2] * 100 + +m[1], texto: t };
  return null;
}

// Busca la fila de encabezados y devuelve en qué columna quedó cada cosa.
function ubicarColumnasBalcomp(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 40); r++) {
    const cols = {};
    const porFecha = [];
    for (let c = 1; c <= Math.min(ws.columnCount, 30); c++) {
      const t = _norm(ws.getCell(r, c).value);
      if (!t) continue;
      if (/^CUENTA/.test(t)) { if (!cols.cuenta) cols.cuenta = c; }
      else if (/^SALDO ANTERIOR/.test(t)) cols.anterior = c;
      else if (/^SALDO FINAL/.test(t)) cols.final = c;
      else if (/^DEBITO/.test(t)) cols.debe = c;
      else if (/^CREDITO/.test(t)) cols.haber = c;
      else if (/^MOVIMIENTO/.test(t)) cols.movimiento = c;
      else {
        const f = _fechaDeEncabezado(t);
        if (f) porFecha.push({ col: c, ...f });
      }
    }
    // las columnas rotuladas con fecha completan lo que no vino rotulado a secas
    if (porFecha.length) {
      porFecha.sort((a, b) => a.orden - b.orden);
      const ultima = porFecha[porFecha.length - 1];
      if (!cols.final) { cols.final = ultima.col; cols.etiquetaFinal = ultima.texto; }
      if (!cols.anterior && porFecha.length > 1) cols.anterior = porFecha[0].col;
    }
    // hace falta el saldo final, o el anterior más los movimientos para reconstruirlo.
    // Débitos y créditos SOLOS no alcanzan: darían el movimiento del mes, no el saldo.
    if (cols.cuenta && (cols.final || (cols.anterior && cols.debe && cols.haber))) {
      return { fila: r, ...cols };
    }
  }
  return null;
}

// Devuelve { saldos: {codigo: saldoFinal}, cuentas, avisos, columnas }.
function leerSaldosDeBalcomp(wb) {
  const avisos = [];
  for (const ws of wb.worksheets) {
    const cols = ubicarColumnasBalcomp(ws);
    if (!cols) continue;

    const saldos = {};
    const filaDe = {};
    let filas = 0, reconstruidos = 0, sinDato = 0;
    const repetidos = [];
    for (let r = cols.fila + 1; r <= ws.rowCount; r++) {
      const texto = _textoB(ws.getCell(r, cols.cuenta).value).trim();
      const m = RE_CUENTA_B.exec(texto);
      if (!m) continue;                       // títulos de capítulo, TOTAL, filas vacías

      let final = cols.final ? _numeroB(ws.getCell(r, cols.final)) : null;
      if (final === null) {
        // la celda es una fórmula sin resultado guardado (el archivo no pasó por Excel):
        // se rehace con la misma cuenta que hace el archivo
        if (!cols.anterior || (!cols.debe && !cols.haber)) { sinDato++; continue; }
        const ant = _numeroB(ws.getCell(r, cols.anterior)) || 0;
        const d = cols.debe ? (_numeroB(ws.getCell(r, cols.debe)) || 0) : 0;
        const h = cols.haber ? (_numeroB(ws.getCell(r, cols.haber)) || 0) : 0;
        final = ant + (d - h);
        reconstruidos++;
      }
      filas++;
      final = Math.round(final * 100) / 100;
      // un código repetido pisaría al anterior sin que se note: hay que avisarlo
      if (Object.prototype.hasOwnProperty.call(saldos, m[1])) {
        repetidos.push({ codigo: m[1], nombre: m[2], filas: [filaDe[m[1]], r],
                         valores: [saldos[m[1]], final] });
      }
      saldos[m[1]] = final;
      filaDe[m[1]] = r;
    }

    const n = Object.keys(saldos).length;
    if (!n) continue;

    if (cols.etiquetaFinal) {
      avisos.push(`El saldo de cada cuenta se tomó de la columna "${cols.etiquetaFinal}". ` +
                  `Si no es esa, no importes.`);
    }
    if (reconstruidos) {
      avisos.push(`${reconstruidos} de ${filas} filas no traían el saldo final calculado, ` +
                  `así que se rehizo como saldo anterior + (débitos − créditos), que es la ` +
                  `fórmula del propio archivo.`);
    }
    for (const d of repetidos) {
      const iguales = Math.abs(d.valores[0] - d.valores[1]) < 0.005;
      avisos.push(`El código ${d.codigo} ("${d.nombre}") está en las filas ` +
        `${d.filas.join(" y ")}${iguales
          ? `, con el mismo saldo (${d.valores[1]}). Se guardó una sola vez.`
          : `, con saldos distintos (${d.valores[0]} y ${d.valores[1]}). Se guardó el de la ` +
            `fila ${d.filas[1]}: revisalo antes de importar.`}`);
    }
    if (sinDato) {
      avisos.push(`${sinDato} fila(s) quedaron afuera porque no se les pudo determinar el saldo.`);
    }
    return { saldos, cuentas: n, filas, repetidos, avisos, columnas: cols, hoja: ws.name };
  }
  return { saldos: {}, cuentas: 0, filas: 0, repetidos: [], hoja: null, columnas: null,
           avisos: ["No encontré la tabla del balance. Hace falta una columna de cuenta y una " +
                    "de saldo final — puede llamarse 'Saldo Final' o con su fecha, como " +
                    "'Saldo al 31 de Mayo de 2026'. Con débitos y créditos solos no alcanza: " +
                    "darían el movimiento del mes, no el saldo."] };
}

// ExcelJS no lee `.xls` (el formato viejo), y el informe del mes pasado bien puede estar
// guardado así. SheetJS sí lo lee, pero devuelve otra forma de libro, así que se la disfraza
// de la que espera `leerSaldosDeBalcomp`. De un `.xls` sólo vienen los valores calculados,
// que es justo lo que hace falta acá.
function _adaptarSheetJs(libro) {
  const hojas = libro.SheetNames.map(nombre => {
    const hoja = libro.Sheets[nombre];
    const r = XLSX.utils.decode_range(hoja["!ref"] || "A1");
    return {
      name: nombre,
      rowCount: r.e.r + 1,
      columnCount: r.e.c + 1,
      getCell(fila, col) {
        const c = hoja[XLSX.utils.encode_cell({ r: fila - 1, c: col - 1 })];
        return { value: c ? (c.v !== undefined ? c.v : null) : null };
      },
    };
  });
  return { worksheets: hojas };
}

// Abre el archivo que haya subido la usuaria, sea .xlsx o .xls.
async function abrirLibroDeSaldos(buffer, nombreArchivo) {
  if (/\.xls$/i.test(nombreArchivo || "")) {
    return _adaptarSheetJs(XLSX.read(buffer, { type: "array" }));
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

if (typeof module !== "undefined") {
  module.exports = { leerSaldosDeBalcomp, ubicarColumnasBalcomp, abrirLibroDeSaldos, RE_CUENTA_B };
}
