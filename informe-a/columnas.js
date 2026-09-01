// Insertar una COLUMNA y que las fórmulas queden apuntando a donde tienen que apuntar.
//
// Es el equivalente horizontal de lo que `formula_utils.js` hace con las filas, y hace falta
// por lo mismo: ni ExcelJS ni openpyxl reacomodan las fórmulas al insertar, a diferencia de
// Excel. Sin esto, meter una columna de centro de costo en `Dist.de gastos` deja las 118
// fórmulas de la zona de meses (`SUM(Z8:AK8)` y las columnas de cada mes) apuntando una
// columna a la izquierda, y cada mes muestra el importe del mes anterior.
//
// Medido sobre el archivo real, insertando en `Z` (la primera columna después de los centros
// de costo): 0 fórmulas de otras hojas miran esa zona —`Gastos Acumulados` lee las columnas
// de centro de costo, que están antes del corte— y 118 fórmulas de la propia hoja sí.
//
// Los nombres llevan prefijo `IC_` a propósito: el sitio carga los .js como <script> sueltos,
// en UN solo ámbito global, así que dos const con el mismo nombre en archivos distintos son
// un SyntaxError que tumba el segundo archivo entero. Ya pasó con `formula_hojas.js`.

const IC_TEXTO = /"[^"]*"/g;
const IC_MARCA = String.fromCharCode(2);
const IC_MARCA_RE = new RegExp(`${IC_MARCA}(\\d+)${IC_MARCA}`, "g");
// se excluye lo que va seguido de "(" para no confundir un nombre de función con una celda
const IC_REF_LOCAL = /(?<![A-Z0-9_$!.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\s*\()/g;

function icIndice(col) {
  return String(col).toUpperCase().split("")
    .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
}

function icLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Escapa el nombre de la hoja para meterlo en una expresión regular.
function icEscapar(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Una columna se corre si está en el punto de inserción o a su derecha. Un rango que
// "contiene" el corte se expande, igual que haría Excel: su fin corre y su principio no.
function icCorrer(col, desde) {
  return icIndice(col) >= desde ? icLetra(icIndice(col) + 1) : col;
}

// Referencias que llevan el nombre de la hoja: 'Dist.de gastos'!Z8  |  'Dist.de gastos'!$Z$8:$AK$8
function icShiftConHoja(formula, nombreHoja, desde) {
  const n = icEscapar(nombreHoja);
  const re = new RegExp(
    `((?:'${n}'|${n})!)(\\$?)([A-Z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?`, "g");
  return String(formula).replace(re, (m, pre, d1, c1, d2, r1, d3, c2, d4, r2) => {
    const a = `${pre}${d1}${icCorrer(c1, desde)}${d2}${r1}`;
    if (c2 === undefined) return a;
    return `${a}:${d3}${icCorrer(c2, desde)}${d4}${r2}`;
  });
}

// Referencias sin nombre de hoja, que son las de adentro de la propia hoja. Antes de tocar
// nada se guardan aparte los textos entre comillas y las referencias que SÍ llevan hoja,
// para no confundirlos con referencias locales.
function icShiftLocal(formula, desde) {
  const guardados = [];
  const guardar = (m) => `${IC_MARCA}${guardados.push(m) - 1}${IC_MARCA}`;
  const conHoja = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?/g;

  let f = String(formula).replace(IC_TEXTO, guardar).replace(conHoja, guardar);
  f = f.replace(IC_REF_LOCAL, (m, d1, col, d2, fila) =>
    `${d1}${icCorrer(col, desde)}${d2}${fila}`);
  return f.replace(IC_MARCA_RE, (_, i) => guardados[Number(i)]);
}

// Corre TODAS las fórmulas del libro para una columna insertada en `hoja`.
function icShiftTodas(wb, nombreHoja, desde) {
  let tocadas = 0;
  for (const ws of wb.worksheets) {
    const propia = ws.name === nombreHoja;
    // Se recorren SOLO las celdas que existen. Con getCell(r,c) sobre todo el rectángulo,
    // ExcelJS materializa cada celda vacía que toca: en "Sumas y Saldos" son 386 x 253 y el
    // alta hace tres inserciones, así que la corrida pasaba de segundos a minutos y encima
    // engordaba el archivo con celdas que no existían.
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const f = cell.formula;
        if (!f) return;
        let nueva = icShiftConHoja(f, nombreHoja, desde);
        if (propia) nueva = icShiftLocal(nueva, desde);
        if (nueva === f) return;
        // El resultado guardado se conserva: no cambia por mover una referencia, y perderlo
        // deja la celda vacía para todo lo que lea el archivo sin abrirlo en Excel.
        const res = cell.value && typeof cell.value === "object" ? cell.value.result : undefined;
        cell.value = res === undefined ? { formula: nueva } : { formula: nueva, result: res };
        tocadas++;
      });
    });
  }
  return tocadas;
}

// Inserta una columna vacía en `colDestino` (índice, 1 = A). Todo lo que estaba ahí y a su
// derecha se corre una posición, y las fórmulas del libro se reacomodan.
//
// No se usa `spliceColumns` de ExcelJS: mueve los valores pero no los estilos ni el ancho, y
// además deja el texto de las fórmulas tal cual, que es lo que importa acá. Moviendo a mano
// queda claro qué se copia.
function insertColumnEn(wb, nombreHoja, colDestino) {
  const ws = wb.getWorksheet(nombreHoja);
  if (!ws) throw new Error(`El archivo no tiene la hoja '${nombreHoja}'.`);
  const ultima = ws.columnCount;

  // Sólo las filas que existen: recorrer el rectángulo entero materializa celdas vacías,
  // que es lento y además las escribe en el archivo.
  const filas = [];
  ws.eachRow({ includeEmpty: false }, (row, r) => filas.push(r));

  // De derecha a izquierda, para no pisar lo que todavía no se movió.
  for (let c = ultima; c >= colDestino; c--) {
    const origen = ws.getColumn(c);
    const destino = ws.getColumn(c + 1);
    for (const r of filas) {
      const de = ws.getCell(r, c);
      if (de.type === ExcelJS.ValueType.Null && !de.style) continue;
      const a = ws.getCell(r, c + 1);
      a.value = de.value;
      a.style = de.style;
    }
    destino.width = origen.width;
    destino.hidden = origen.hidden;
  }

  // La columna nueva queda vacía y visible, con el ancho de la que estaba ahí.
  for (const r of filas) ws.getCell(r, colDestino).value = null;
  ws.getColumn(colDestino).hidden = false;

  const tocadas = icShiftTodas(wb, nombreHoja, colDestino);
  return { columna: icLetra(colDestino), formulasReacomodadas: tocadas };
}

if (typeof module !== "undefined") {
  module.exports = { insertColumnEn, icShiftConHoja, icShiftLocal, icLetra, icIndice };
}
