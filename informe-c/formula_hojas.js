// Versión parametrizada de las utilidades de fórmulas del Informe A: acá hay que
// insertar filas en más de una hoja (SALDOS y Activo y Pasivo), así que la hoja no
// puede estar fija en el código como allá.
//
// Igual que allá: ni ExcelJS ni openpyxl reacomodan las fórmulas al insertar una
// fila (Excel sí), y hay que simularlo a mano para todas las fórmulas del archivo
// que dependan de la hoja donde se insertó.

function escaparRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const TEXTO_ENTRE_COMILLAS = /"[^"]*"/g;
const REF_CON_HOJA = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?/g;
const REF_LOCAL = /(?<![A-Z0-9_$!.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\s*\()/g;
const MARCA = String.fromCharCode(1);
const MARCA_RE = new RegExp(`${MARCA}(\\d+)${MARCA}`, "g");

function crearShifters(nombreHoja) {
  // 'SALDOS'!$G$3 | SALDOS!G3:G100 | SALDOS!G3 — con o sin comillas.
  //
  // Los `$` se capturan y se devuelven tal cual: reescribir `Hoja1!$A$2:$E$377` como
  // `Hoja1!A2:E431` da el mismo número, pero deja de ser una referencia absoluta. Eso
  // rompía cosas río abajo (normalizarRangosVlookup busca los rangos por su `$` para
  // reparar los que están cortos) y volvía frágil cualquier fórmula que después se
  // copie. Insertar una fila SÍ corre el número aunque la referencia sea absoluta:
  // el `$` manda al copiar, no al insertar.
  const REF_HOJA_RE = new RegExp(
    `('?${escaparRegex(nombreHoja)}'?!)(\\$?)([A-Z]{1,3})(\\$?)(\\d+)` +
    `(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?`, "g"
  );

  const shiftRow = (row, insertBeforeRow) => {
    const n = parseInt(row, 10);
    return n >= insertBeforeRow ? n + 1 : n;
  };

  function shiftFormula(formula, insertBeforeRow) {
    return formula.replace(REF_HOJA_RE, (m, prefix, dc1, col1, df1, row1, dc2, col2, df2, row2) => {
      const uno = `${prefix}${dc1}${col1}${df1}${shiftRow(row1, insertBeforeRow)}`;
      if (col2 === undefined) return uno;
      return `${uno}:${dc2}${col2}${df2}${shiftRow(row2, insertBeforeRow)}`;
    });
  }

  function shiftFormulaLocal(formula, insertBeforeRow) {
    const guardados = [];
    const guardar = (m) => `${MARCA}${guardados.push(m) - 1}${MARCA}`;
    let f = formula.replace(TEXTO_ENTRE_COMILLAS, guardar).replace(REF_CON_HOJA, guardar);
    f = f.replace(REF_LOCAL, (m, d1, col, d2, row) =>
      `${d1}${col}${d2}${shiftRow(row, insertBeforeRow)}`);
    return f.replace(MARCA_RE, (_, i) => guardados[Number(i)]);
  }

  return { shiftFormula, shiftFormulaLocal, REF_HOJA_RE };
}

// Recorre TODAS las hojas y reacomoda las fórmulas afectadas por insertar una fila
// en `nombreHoja`. Presupone que las fórmulas compartidas ya fueron materializadas
// (abrirWorkbook del Informe A lo garantiza), así que acá no existen clones.
function shiftAllFormulasEn(wb, nombreHoja, insertBeforeRow) {
  const { shiftFormula, shiftFormulaLocal } = crearShifters(nombreHoja);
  const nombreRe = new RegExp(`'?${escaparRegex(nombreHoja)}'?!`);
  let modificadas = 0;

  wb.worksheets.forEach(ws => {
    const esLaHoja = ws.name === nombreHoja;
    ws.eachRow(row => row.eachCell(cell => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;

      let nueva = v.formula;
      if (nombreRe.test(nueva)) nueva = shiftFormula(nueva, insertBeforeRow);
      if (esLaHoja) nueva = shiftFormulaLocal(nueva, insertBeforeRow);

      if (nueva !== v.formula) {
        // resultado cacheado descartado a propósito: Excel recalcula al abrir
        cell.value = { formula: nueva };
        modificadas++;
      }
    }));
  });

  return modificadas;
}

function insertRowEn(wb, nombreHoja, insertAtRow) {
  const ws = wb.getWorksheet(nombreHoja);
  if (!ws) throw new Error(`El archivo no tiene la hoja '${nombreHoja}'.`);
  const modificadas = shiftAllFormulasEn(wb, nombreHoja, insertAtRow);
  ws.spliceRows(insertAtRow, 0, []);
  return modificadas;
}

if (typeof module !== "undefined") {
  module.exports = { crearShifters, shiftAllFormulasEn, insertRowEn };
}
