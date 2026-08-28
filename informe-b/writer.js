// En el navegador, ExcelJS ya viene cargado como variable global por el
// script del CDN. En Node (para test) se toma vía require.
if (typeof ExcelJS === "undefined" && typeof require !== "undefined") {
  var ExcelJS = require("exceljs");
}

// `saldosAnteriores` es {codigo: saldo} de la corrida del mes pasado. Si viene, la
// columna "Saldo anterior" sale ya cargada; si falta una cuenta, esa celda queda
// amarilla para completarla a mano, que es como funcionaba todo antes.
// `filasExport` son las filas crudas del .xls de SISE tal como se subio. Van al final,
// como una hoja mas, para que el informe se explique solo: resultado, desglose y origen en
// el mismo archivo, sin depender de encontrar el export de ese mes seis meses despues.
// `exportInfo` es { filas, colDebe, colHaber }: las filas crudas del .xls de SISE tal como
// se subio, y en que columnas (0 = A) estan los importes en dolares. Con eso el archivo se
// arma con FORMULAS en vez de con numeros copiados:
//
//   hoja "Export de Onvio"  ->  hoja "De donde sale cada saldo"  ->  el informe
//
// Cada eslabon apunta al anterior, asi que los importes los recalcula Excel. La app solo
// decide QUE cuenta va a QUE fila —que es lo que se configura en el panel—; la aritmetica
// deja de ser algo que haya que creerle. Es el mismo criterio del Balance USD, donde el
// motor escribe una sola hoja y el archivo calcula el resto.
async function writeOutputXlsx(lineas, periodo, saldosAnteriores, exportInfo) {
  const expo = exportInfo && exportInfo.filas ? exportInfo : null;
  const filasExport = expo ? expo.filas : null;
  const HOJA_EXPORT = "Export de Onvio";
  const HOJA_DET = "De dónde sale cada saldo";
  const colLetra = (i) => { let n = i + 1, s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  const colExpD = expo ? colLetra(expo.colDebe) : null;
  const colExpH = expo ? colLetra(expo.colHaber) : null;
  const previos = saldosAnteriores || {};
  const hayPrevios = Object.keys(previos).length > 0;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  const bold = { bold: true };
  const yellowFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3B0" } };
  const numFmt = "#,##0.00";

  ws.getCell("B1").value = "SOUTHERN COPPER ARGENTINA S.R.L.";
  ws.getCell("B1").font = bold;
  ws.getCell("B2").value = "BALANCE DE COMPROBACION - DOLARES";
  ws.getCell("B3").value = periodo;
  ws.getCell("B4").value = hayPrevios
    ? "Columna C (Saldo anterior): la trae la app del mes registrado anterior. Lo que quede en amarillo es una cuenta sin saldo guardado."
    : "Columna C (Saldo anterior, en amarillo): pegar a mano del informe del mes pasado.";
  ws.getCell("B4").font = { italic: true, size: 9 };

  // La columna A lleva la cuenta REAL, la que viene de Onvio, y va oculta. La B lleva la
  // que ve el cliente, que usa el plan de cuentas viejo. Los dos codigos conviven porque no
  // hay ninguna regla que los relacione: fueron reasignados, no reformateados. La
  // equivalencia esta declarada en mapping.json (campo `cliente`).
  ws.getCell(5, 1).value = "Cuenta real (Onvio)";
  ws.getCell(5, 1).font = bold;
  ws.getCell(5, 1).border = { bottom: { style: "thin" } };

  const headers = ["Cuenta Contable", "Saldo anterior", "Debitos del Mes Dolares",
    "Creditos del mes Dolares", "Movimiento del Mes Dolares", "Saldo Final"];
  headers.forEach((h, j) => {
    const c = ws.getCell(5, 2 + j);
    c.value = h;
    c.font = bold;
    c.border = { bottom: { style: "thin" } };
  });

  // Una fila marcada `ocultar_si_cero` es una cuenta que existe en el plan pero que casi
  // nunca mueve, y que el informe que se venia entregando no mostraba. Se saltea SOLO
  // mientras este en cero: si algun mes trae movimiento vuelve a aparecer sola, asi no hay
  // forma de que un importe quede afuera sin que se vea. Mismo criterio que las columnas de
  // centro de costo del Balance USD.
  const visibles = lineas.filter(l => !(l.ocultar_si_cero && Math.abs(l.debe) < 0.005 && Math.abs(l.haber) < 0.005));

  let row = 6;
  let currentCat = null;
  for (const l of visibles) {
    if (l.category !== currentCat) {
      currentCat = l.category;
      ws.getCell(row, 2).value = currentCat;
      ws.getCell(row, 2).font = bold;
      row++;
    }
    // Una cuenta madre o un rango de proveedores se alimentan de VARIAS cuentas reales, asi
    // que no hay una sola para poner en A: queda vacia, y el desglose va en la hoja
    // "Detalle Subcuentas". Es como venia armado el informe a mano.
    const cli = l.cliente || { code: l.code, description: l.description };
    if (l.type !== "parent" && l.type !== "range") {
      ws.getCell(row, 1).value = `${l.code} - ${l.description}`;
    }
    ws.getCell(row, 2).value = `${cli.code} - ${cli.description}`;
    const cCell = ws.getCell(row, 3);
    const previo = previos[l.code];
    if (typeof previo === "number") cCell.value = round2(previo);
    else cCell.fill = yellowFill;      // sin saldo guardado: se completa a mano
    cCell.numFmt = numFmt;
    // Los importes salen de sumar el detalle, no de copiar lo que calculo la app. El
    // criterio es el codigo del cliente, que es la clave de la primera columna de esa hoja.
    const clave = String(cli.code);
    const d = ws.getCell(row, 4);
    d.value = expo
      ? { formula: `SUMIF('${HOJA_DET}'!$A:$A,"${clave}",'${HOJA_DET}'!$D:$D)`, result: round2(l.debe) }
      : round2(l.debe);
    d.numFmt = numFmt;
    const e = ws.getCell(row, 5);
    e.value = expo
      ? { formula: `SUMIF('${HOJA_DET}'!$A:$A,"${clave}",'${HOJA_DET}'!$E:$E)`, result: round2(l.haber) }
      : round2(l.haber);
    e.numFmt = numFmt;
    const mov = ws.getCell(row, 6); mov.value = { formula: `D${row}-E${row}` }; mov.numFmt = numFmt;
    const sf = ws.getCell(row, 7); sf.value = { formula: `C${row}+F${row}` }; sf.numFmt = numFmt;
    row++;
  }

  const totalRow = row + 1;
  ws.getCell(totalRow, 2).value = "TOTAL";
  ws.getCell(totalRow, 2).font = bold;
  for (const col of [3, 4, 5, 6, 7]) {
    const letter = String.fromCharCode(64 + col);
    const c = ws.getCell(totalRow, col);
    c.value = { formula: `SUM(${letter}6:${letter}${row - 1})` };
    c.font = bold;
    c.numFmt = numFmt;
  }

  ws.columns = [
    { width: 46 }, { width: 55 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
  ];
  ws.getColumn(1).hidden = true;      // la numeracion real: esta ahi, pero el cliente no la ve

  // --- Segunda hoja: de donde sale cada saldo ---
  // Antes esta hoja solo desglosaba las cuentas madre: se podia rastrear el 6,5% del
  // importe. El resto —incluidos los 118 proveedores metidos en una sola fila— no tenia
  // forma de verificarse. Ahora hay una fila por cada cuenta de Onvio que entra al
  // informe, y su total tiene que dar el mismo total que el informe: eso la convierte en
  // un control, no en un listado.
  const ws2 = wb.addWorksheet("De dónde sale cada saldo");
  ws2.getCell("A1").value = "De dónde sale cada saldo del informe";
  ws2.getCell("A1").font = { bold: true, size: 13 };
  ws2.getCell("A2").value =
    "Una fila por cada cuenta de Onvio que alimenta el informe. Filtrá por la primera " +
    "columna para ver de qué cuentas sale cualquier renglón. El total tiene que coincidir " +
    "con el del informe.";
  ws2.getCell("A2").font = { italic: true, size: 9 };

  // La columna A lleva SOLO el codigo del cliente: es la clave por la que el informe suma
  // esta hoja, y una clave tiene que ser exacta. El nombre va aparte, para leer.
  const headers2 = ["Fila del informe", "Nombre de la fila", "Cuenta de Onvio",
    "Debitos del Mes Dolares", "Creditos del mes Dolares", "Movimiento", "¿Movió este mes?"];
  headers2.forEach((h, j) => {
    const c = ws2.getCell(4, 1 + j);
    c.value = h; c.font = bold; c.border = { bottom: { style: "thin" } };
  });

  let r2 = 5;
  for (const l of visibles) {
    const cli = l.cliente || { code: l.code, description: l.description };
    for (const h of (l.detalle || [])) {
      ws2.getCell(r2, 1).value = String(cli.code);
      ws2.getCell(r2, 2).value = cli.description;
      ws2.getCell(r2, 3).value = `${h.code} - ${h.description}`;
      // Si la cuenta vino en el export, el importe APUNTA a su fila en la hoja del export.
      // Si no vino, no hay a donde apuntar y va un 0: esa cuenta esta declarada pero quieta.
      const d = ws2.getCell(r2, 4);
      d.value = (expo && h.fila)
        ? { formula: `'${HOJA_EXPORT}'!${colExpD}${h.fila}`, result: round2(h.debe) }
        : round2(h.debe);
      d.numFmt = numFmt;
      const e2 = ws2.getCell(r2, 5);
      e2.value = (expo && h.fila)
        ? { formula: `'${HOJA_EXPORT}'!${colExpH}${h.fila}`, result: round2(h.haber) }
        : round2(h.haber);
      e2.numFmt = numFmt;
      const m = ws2.getCell(r2, 6); m.value = { formula: `D${r2}-E${r2}` }; m.numFmt = numFmt;
      ws2.getCell(r2, 7).value = (Math.abs(h.debe) > 0.005 || Math.abs(h.haber) > 0.005) ? "sí" : "no";
      r2++;
    }
  }
  const totalDet = r2 + 1;
  ws2.getCell(totalDet, 1).value = "TOTAL";
  ws2.getCell(totalDet, 1).font = bold;
  for (const col of [4, 5, 6]) {
    const letter = String.fromCharCode(64 + col);
    const c = ws2.getCell(totalDet, col);
    c.value = { formula: `SUM(${letter}5:${letter}${r2 - 1})` };
    c.font = bold; c.numFmt = numFmt;
  }
  ws2.getCell(totalDet + 1, 1).value = "Tiene que dar lo mismo que el TOTAL del informe.";
  ws2.getCell(totalDet + 1, 1).font = { italic: true, size: 9 };
  ws2.autoFilter = { from: { row: 4, column: 1 }, to: { row: r2 - 1, column: 7 } };
  ws2.views = [{ state: "frozen", ySplit: 4 }];
  ws2.columns = [{ width: 14 }, { width: 40 }, { width: 48 }, { width: 18 }, { width: 18 },
                 { width: 16 }, { width: 15 }];

  // --- Tercera hoja: el export tal como se subio ---
  if (filasExport && filasExport.length) {
    const ws3 = wb.addWorksheet(HOJA_EXPORT);
    filasExport.forEach((fila, i) => {
      (fila || []).forEach((v, j) => {
        if (v !== null && v !== undefined && v !== "") ws3.getCell(i + 1, j + 1).value = v;
      });
    });
    ws3.columns = [{ width: 52 }].concat(Array(9).fill({ width: 15 }));
  }

  // El movimiento y el saldo final son formulas, y ExcelJS las escribe SIN resultado: el
  // archivo recien generado tiene esas dos columnas en blanco hasta que Excel las calcula.
  // Con esto Excel recalcula todo al abrirlo, en vez de depender de que lo haga solo.
  // El motor del Informe A ya lo hacia; este no, y por eso el archivo de julio llegaba con
  // el saldo final vacio.
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  return wb;
}

function round2(n) { return Math.round(n * 100) / 100; }

if (typeof module !== "undefined") {
  module.exports = { writeOutputXlsx };
}
