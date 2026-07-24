// Port de formula_utils.py.
// Ni openpyxl ni ExcelJS reacomodan las fórmulas cuando se inserta una fila (a
// diferencia de Excel). Estas funciones simulan ese comportamiento a mano para
// cualquier fórmula del workbook que apunte a la hoja donde se insertó la fila.

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

// Recorre TODAS las hojas y reacomoda las fórmulas que referencian
// 'Sumas y Saldos'. Devuelve la cantidad de celdas modificadas.
function shiftAllFormulas(wb, insertBeforeRow) {
  let modificadas = 0;

  wb.worksheets.forEach(ws => {
    // Las fórmulas "shared" no guardan texto propio: apuntan a una celda maestra.
    // Si alguna derivara de una maestra que referencia 'Sumas y Saldos', reescribir
    // solo la maestra dejaría a las hijas apuntando a filas viejas, así que se corta
    // en vez de generar un archivo con totales silenciosamente mal.
    const maestras = {};
    ws.eachRow(row => row.eachCell(cell => {
      if (cell.formula) maestras[cell.address] = cell.formula;
    }));

    ws.eachRow(row => row.eachCell(cell => {
      const v = cell.value;
      if (!v || typeof v !== "object") return;

      if (v.sharedFormula) {
        const maestra = maestras[v.sharedFormula];
        if (maestra && maestra.includes("Sumas y Saldos")) {
          throw new Error(
            `La celda ${ws.name}!${cell.address} usa una fórmula compartida que ` +
            `referencia 'Sumas y Saldos'. Este motor no sabe reacomodar ese caso ` +
            `y no puede garantizar los totales. NO se generó ningún archivo.`
          );
        }
        return;
      }

      if (typeof v.formula === "string" && v.formula.includes("Sumas y Saldos")) {
        const nueva = shiftFormula(v.formula, insertBeforeRow);
        if (nueva !== v.formula) {
          // Se descarta el resultado cacheado: quedó viejo tras el reacomodo y
          // Excel lo recalcula al abrir (ver fullCalcOnLoad en motor.js).
          cell.value = { formula: nueva };
          modificadas++;
        }
      }
    }));
  });

  return modificadas;
}

// Inserta una fila vacía en 'Sumas y Saldos', corriendo los datos hacia abajo, y
// reacomoda las fórmulas de todo el workbook que dependían de las filas movidas.
function insertRowSumasYSaldos(wsSs, wb, insertAtRow) {
  wsSs.spliceRows(insertAtRow, 0, []);
  return shiftAllFormulas(wb, insertAtRow);
}

if (typeof module !== "undefined") {
  module.exports = { shiftFormula, shiftAllFormulas, insertRowSumasYSaldos, REF_RE };
}
