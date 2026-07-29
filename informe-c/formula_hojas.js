// Versión parametrizada de las utilidades de fórmulas del Informe A: acá hay que
// insertar filas en más de una hoja (SALDOS y Activo y Pasivo), así que la hoja no
// puede estar fija en el código como allá.
//
// Igual que allá: ni ExcelJS ni openpyxl reacomodan las fórmulas al insertar una
// fila (Excel sí), y hay que simularlo a mano para todas las fórmulas del archivo
// que dependan de la hoja donde se insertó.

// Los nombres llevan prefijo FH_ a propósito: estos scripts se cargan como <script>
// sueltos y comparten UN solo ámbito global con los del Informe A, que declara constantes
// con estos mismos nombres. Repetir un `const` en el ámbito global es un SyntaxError que
// tumba el archivo entero (y con él insertRowEn) sin que se vea el motivo.
function escaparRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const FH_TEXTO_ENTRE_COMILLAS = /"[^"]*"/g;
const FH_REF_CON_HOJA = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?/g;
const FH_REF_LOCAL = /(?<![A-Z0-9_$!.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\s*\()/g;
const FH_MARCA = String.fromCharCode(1);
const FH_MARCA_RE = new RegExp(`${FH_MARCA}(\\d+)${FH_MARCA}`, "g");

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
    const guardar = (m) => `${FH_MARCA}${guardados.push(m) - 1}${FH_MARCA}`;
    let f = formula.replace(FH_TEXTO_ENTRE_COMILLAS, guardar).replace(FH_REF_CON_HOJA, guardar);
    f = f.replace(FH_REF_LOCAL, (m, d1, col, d2, row) =>
      `${d1}${col}${d2}${shiftRow(row, insertBeforeRow)}`);
    return f.replace(FH_MARCA_RE, (_, i) => guardados[Number(i)]);
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

// Borrar una fila es la operación inversa: lo que está debajo sube uno. Es MUCHO más
// delicada que insertar, porque cualquier fórmula que apunte a la fila borrada queda
// en #REF! y eso se propaga a los estados. Por eso primero se comprueba que nadie la
// referencie y, si alguien lo hace, se corta sin tocar el archivo.
function borrarFilaEn(wb, nombreHoja, filaBorrada) {
  const ws = wb.getWorksheet(nombreHoja);
  if (!ws) throw new Error(`El archivo no tiene la hoja '${nombreHoja}'.`);

  const reHoja = new RegExp(`'?${escaparRegex(nombreHoja)}'?!\\$?[A-Z]{1,3}\\$?${filaBorrada}(?!\\d)`);
  const reLocal = new RegExp(`(?<![A-Z0-9_$!.])\\$?[A-Z]{1,3}\\$?${filaBorrada}(?!\\d)`);
  const culpables = [];
  wb.worksheets.forEach(w => {
    const esLaHoja = w.name === nombreHoja;
    w.eachRow((row, r) => row.eachCell(cell => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      if (esLaHoja && r === filaBorrada) return;        // la propia fila se va con ella
      const f = v.formula;
      // los rangos se achican solos; lo que rompe es una referencia suelta a esa fila
      const sinRangos = f.replace(/\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+/g, " ");
      if (reHoja.test(sinRangos) || (esLaHoja && reLocal.test(sinRangos))) {
        culpables.push(`${w.name}!${cell.address} = ${f}`);
      }
    }));
  });
  if (culpables.length) {
    throw new Error(
      `No borro ${nombreHoja}!${filaBorrada}: hay ${culpables.length} fórmula(s) que la ` +
      `referencian y quedarían en #REF!. ` + culpables.slice(0, 4).join(" | ")
    );
  }

  const nombreRe = new RegExp(`'?${escaparRegex(nombreHoja)}'?!`);
  let modificadas = 0;
  wb.worksheets.forEach(w => {
    const esLaHoja = w.name === nombreHoja;
    w.eachRow(row => row.eachCell(cell => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      let nueva = v.formula;
      if (nombreRe.test(nueva)) nueva = bajarUno(nueva, nombreHoja, filaBorrada, false);
      if (esLaHoja) nueva = bajarUno(nueva, nombreHoja, filaBorrada, true);
      if (nueva === v.formula) return;
      // El resultado cacheado SÍ se conserva, al revés que al insertar: tras borrar una
      // fila, `G258` pasa a decir `G257` pero apunta al mismo contenido de siempre, así
      // que el valor sigue siendo el correcto. Descartarlo dejaría el maestro sin los
      // resultados que el motor usa para leer el nombre de las cuentas que lo toman de
      // una fórmula (`SALDOS!E426 = +Hoja1!A267`), y esas cuentas desaparecerían.
      cell.value = v.result === undefined ? { formula: nueva } : { formula: nueva, result: v.result };
      modificadas++;
    }));
  });

  ws.spliceRows(filaBorrada, 1);
  return modificadas;
}

// Resta uno a toda referencia a filas POSTERIORES a la borrada, dentro de `nombreHoja`.
function bajarUno(formula, nombreHoja, filaBorrada, local) {
  const menos = (row) => {
    const n = parseInt(row, 10);
    return n > filaBorrada ? n - 1 : n;
  };
  if (!local) {
    const RE = new RegExp(
      `('?${escaparRegex(nombreHoja)}'?!)(\\$?)([A-Z]{1,3})(\\$?)(\\d+)` +
      `(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?`, "g");
    return formula.replace(RE, (m, pre, dc1, c1, df1, r1, dc2, c2, df2, r2) => {
      const uno = `${pre}${dc1}${c1}${df1}${menos(r1)}`;
      if (c2 === undefined) return uno;
      return `${uno}:${dc2}${c2}${df2}${menos(r2)}`;
    });
  }
  const guardados = [];
  const guardar = (m) => `${FH_MARCA}${guardados.push(m) - 1}${FH_MARCA}`;
  let f = formula.replace(FH_TEXTO_ENTRE_COMILLAS, guardar).replace(FH_REF_CON_HOJA, guardar);
  f = f.replace(FH_REF_LOCAL, (m, d1, col, d2, row) => `${d1}${col}${d2}${menos(row)}`);
  return f.replace(FH_MARCA_RE, (_, i) => guardados[Number(i)]);
}

if (typeof module !== "undefined") {
  module.exports = { crearShifters, shiftAllFormulasEn, insertRowEn, borrarFilaEn };
}
