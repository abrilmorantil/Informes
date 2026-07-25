// Ni openpyxl ni ExcelJS reacomodan las fórmulas cuando se inserta una fila (a
// diferencia de Excel). Estas funciones simulan ese comportamiento a mano para
// todas las fórmulas del archivo que dependan de la hoja donde se insertó la fila.

// 'Sumas y Saldos'!$E$130  |  'Sumas y Saldos'!E130:E193  |  'Sumas y Saldos'!E130
const REF_RE = /('Sumas y Saldos'!)\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g;

function shiftRow(row, insertBeforeRow) {
  const n = parseInt(row, 10);
  return n >= insertBeforeRow ? n + 1 : n;
}

// Mueve hacia abajo toda referencia a 'Sumas y Saldos' con fila >= insertBeforeRow.
// Un rango que "contiene" la fila insertada se expande, igual que haría Excel.
function shiftFormula(formula, insertBeforeRow) {
  return formula.replace(REF_RE, (m, prefix, col1, row1, col2, row2) => {
    const nueva1 = shiftRow(row1, insertBeforeRow);
    if (col2 !== undefined) {
      return `${prefix}${col1}${nueva1}:${col2}${shiftRow(row2, insertBeforeRow)}`;
    }
    return `${prefix}${col1}${nueva1}`;
  });
}

// Dentro de la propia hoja 'Sumas y Saldos' las fórmulas se escriben sin el nombre
// de la hoja (`A367`, `SUM(C7:C398)`), así que no las agarra la expresión de arriba.
// También hay que correrlas: si no, las filas que bajan siguen apuntando a la fila
// que ocupaban antes y cada cuenta termina mostrando el importe de otra.
//
// Antes de tocar nada se guardan aparte los textos entre comillas ("1_") y las
// referencias que sí llevan nombre de hoja (SyS!$B:$P), para no confundirlos con
// referencias locales.
const TEXTO_ENTRE_COMILLAS = /"[^"]*"/g;
const REF_CON_HOJA = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?/g;
// se excluye lo que va seguido de "(" para no confundir un nombre de función con una celda
const REF_LOCAL = /(?<![A-Z0-9_$!.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\s*\()/g;

// Marcador imposible de encontrar en una fórmula. REF_LOCAL no lo toma porque
// exige letras delante de los dígitos.
const MARCA = String.fromCharCode(1);
const MARCA_RE = new RegExp(`${MARCA}(\\d+)${MARCA}`, "g");

function shiftFormulaLocal(formula, insertBeforeRow) {
  const guardados = [];
  const guardar = (m) => `${MARCA}${guardados.push(m) - 1}${MARCA}`;

  let f = formula.replace(TEXTO_ENTRE_COMILLAS, guardar).replace(REF_CON_HOJA, guardar);
  f = f.replace(REF_LOCAL, (m, d1, col, d2, row) =>
    `${d1}${col}${d2}${shiftRow(row, insertBeforeRow)}`);
  return f.replace(MARCA_RE, (_, i) => guardados[Number(i)]);
}

const HOJA_BALANCE = "Sumas y Saldos";

// Recorre TODAS las hojas y reacomoda las fórmulas afectadas por la inserción.
// Devuelve la cantidad de celdas modificadas.
function shiftAllFormulas(wb, insertBeforeRow) {
  let modificadas = 0;

  wb.worksheets.forEach(ws => {
    const esHojaBalance = ws.name === HOJA_BALANCE;

    // Las fórmulas "shared" no guardan texto propio: apuntan a una celda maestra.
    // Si alguna dependiera de las filas movidas, reescribir solo la maestra dejaría
    // a las hijas apuntando a filas viejas, así que se corta en vez de generar un
    // archivo con los totales silenciosamente mal.
    const maestras = {};
    ws.eachRow(row => row.eachCell(cell => {
      if (cell.formula) maestras[cell.address] = cell.formula;
    }));

    ws.eachRow(row => row.eachCell(cell => {
      const v = cell.value;
      if (!v || typeof v !== "object") return;

      if (v.sharedFormula) {
        const maestra = maestras[v.sharedFormula];
        if (maestra && (esHojaBalance || maestra.includes(HOJA_BALANCE))) {
          throw new Error(
            `La celda ${ws.name}!${cell.address} usa una fórmula compartida que depende ` +
            `de '${HOJA_BALANCE}'. Este motor no sabe reacomodar ese caso y no puede ` +
            `garantizar los totales. NO se generó ningún archivo.`
          );
        }
        return;
      }

      if (typeof v.formula !== "string") return;

      let nueva = v.formula;
      if (nueva.includes(HOJA_BALANCE)) nueva = shiftFormula(nueva, insertBeforeRow);
      // dentro de la hoja del balance, además, las referencias locales
      if (esHojaBalance) nueva = shiftFormulaLocal(nueva, insertBeforeRow);

      if (nueva !== v.formula) {
        // Se descarta el resultado cacheado: quedó viejo tras el reacomodo, y Excel
        // lo recalcula al abrir (ver fullCalcOnLoad en motor.js).
        cell.value = { formula: nueva };
        modificadas++;
      }
    }));
  });

  return modificadas;
}

// Inserta una fila vacía en 'Sumas y Saldos', corriendo los datos hacia abajo, y
// reacomoda las fórmulas de todo el archivo que dependían de las filas movidas.
//
// Importante: el reacomodo va ANTES de mover las celdas. Así una fórmula que está
// en la fila 367 y dice A367 pasa a decir A368 y recién después baja a la fila 368,
// que es exactamente lo que hace Excel con una referencia relativa.
function insertRowSumasYSaldos(wsSs, wb, insertAtRow) {
  const modificadas = shiftAllFormulas(wb, insertAtRow);
  wsSs.spliceRows(insertAtRow, 0, []);
  return modificadas;
}

if (typeof module !== "undefined") {
  module.exports = {
    shiftFormula, shiftFormulaLocal, shiftAllFormulas, insertRowSumasYSaldos, REF_RE,
  };
}
