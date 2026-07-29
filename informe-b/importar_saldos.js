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

// Busca la fila de encabezados y devuelve en qué columna quedó cada cosa.
function ubicarColumnasBalcomp(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 40); r++) {
    const cols = {};
    for (let c = 1; c <= Math.min(ws.columnCount, 20); c++) {
      const t = _norm(ws.getCell(r, c).value);
      if (!t) continue;
      if (/^CUENTA/.test(t)) cols.cuenta = c;
      else if (/^SALDO ANTERIOR/.test(t)) cols.anterior = c;
      else if (/^DEBITO/.test(t)) cols.debe = c;
      else if (/^CREDITO/.test(t)) cols.haber = c;
      else if (/^SALDO FINAL/.test(t)) cols.final = c;
      else if (/^MOVIMIENTO/.test(t)) cols.movimiento = c;
    }
    if (cols.cuenta && (cols.final || (cols.debe && cols.haber))) {
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
    let reconstruidos = 0, sinDato = 0;
    for (let r = cols.fila + 1; r <= ws.rowCount; r++) {
      const texto = _textoB(ws.getCell(r, cols.cuenta).value).trim();
      const m = RE_CUENTA_B.exec(texto);
      if (!m) continue;                       // títulos de capítulo, TOTAL, filas vacías

      let final = cols.final ? _numeroB(ws.getCell(r, cols.final)) : null;
      if (final === null) {
        // el archivo no pasó por Excel: se rehace la cuenta con la misma fórmula
        const ant = cols.anterior ? (_numeroB(ws.getCell(r, cols.anterior)) || 0) : 0;
        const d = cols.debe ? (_numeroB(ws.getCell(r, cols.debe)) || 0) : 0;
        const h = cols.haber ? (_numeroB(ws.getCell(r, cols.haber)) || 0) : 0;
        if (!cols.debe && !cols.haber) { sinDato++; continue; }
        final = ant + (d - h);
        reconstruidos++;
      }
      saldos[m[1]] = Math.round(final * 100) / 100;
    }

    const n = Object.keys(saldos).length;
    if (!n) continue;
    if (reconstruidos) {
      avisos.push(`${reconstruidos} de ${n} saldos se recalcularon como saldo anterior + ` +
                  `(débitos − créditos), porque el archivo no traía el resultado guardado. ` +
                  `Si lo abrís y guardás en Excel antes de subirlo, se toman tal cual están.`);
    }
    if (sinDato) avisos.push(`${sinDato} fila(s) sin importes que no se pudieron leer.`);
    return { saldos, cuentas: n, avisos, columnas: cols, hoja: ws.name };
  }
  return { saldos: {}, cuentas: 0, hoja: null, columnas: null,
           avisos: ["No encontré la tabla del balance: hace falta una columna de cuenta y, " +
                    "o bien 'Saldo Final', o bien las de débitos y créditos."] };
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
