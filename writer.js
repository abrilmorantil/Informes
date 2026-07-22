// En el navegador, ExcelJS ya viene cargado como variable global por el
// script del CDN. En Node (para test) se toma vía require.
if (typeof ExcelJS === "undefined" && typeof require !== "undefined") {
  var ExcelJS = require("exceljs");
}

async function writeOutputXlsx(lineas, periodo) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  const bold = { bold: true };
  const yellowFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3B0" } };
  const numFmt = "#,##0.00";

  ws.getCell("B1").value = "SOUTHERN COPPER ARGENTINA S.R.L.";
  ws.getCell("B1").font = bold;
  ws.getCell("B2").value = "BALANCE DE COMPROBACION - DOLARES";
  ws.getCell("B3").value = periodo;
  ws.getCell("B4").value = "Columna C (Saldo anterior, en amarillo): pegar a mano del informe del mes pasado.";
  ws.getCell("B4").font = { italic: true, size: 9 };

  const headers = ["Cuenta Contable", "Saldo anterior", "Debitos del Mes Dolares",
    "Creditos del mes Dolares", "Movimiento del Mes Dolares", "Saldo Final"];
  headers.forEach((h, j) => {
    const c = ws.getCell(5, 2 + j);
    c.value = h;
    c.font = bold;
    c.border = { bottom: { style: "thin" } };
  });

  let row = 6;
  let currentCat = null;
  for (const l of lineas) {
    if (l.category !== currentCat) {
      currentCat = l.category;
      ws.getCell(row, 2).value = currentCat;
      ws.getCell(row, 2).font = bold;
      row++;
    }
    ws.getCell(row, 2).value = `${l.code} - ${l.description}`;
    const cCell = ws.getCell(row, 3);
    cCell.fill = yellowFill;
    cCell.numFmt = numFmt;
    const d = ws.getCell(row, 4); d.value = round2(l.debe); d.numFmt = numFmt;
    const e = ws.getCell(row, 5); e.value = round2(l.haber); e.numFmt = numFmt;
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
    {}, { width: 55 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
  ];

  // --- Segunda hoja: detalle de subcuentas ---
  const ws2 = wb.addWorksheet("Detalle Subcuentas");
  const headers2 = ["Cuenta Madre", "Subcuenta", "Debitos del Mes Dolares", "Creditos del mes Dolares"];
  headers2.forEach((h, j) => {
    const c = ws2.getCell(1, 1 + j);
    c.value = h; c.font = bold; c.border = { bottom: { style: "thin" } };
  });
  let r2 = 2;
  for (const l of lineas) {
    for (const h of (l.detalle || [])) {
      ws2.getCell(r2, 1).value = `${l.code} - ${l.description}`;
      ws2.getCell(r2, 2).value = `${h.code} - ${h.description}`;
      ws2.getCell(r2, 3).value = round2(h.debe);
      ws2.getCell(r2, 4).value = round2(h.haber);
      r2++;
    }
  }
  ws2.columns = [{ width: 45 }, { width: 45 }, { width: 18 }, { width: 18 }];

  return wb;
}

function round2(n) { return Math.round(n * 100) / 100; }

if (typeof module !== "undefined") {
  module.exports = { writeOutputXlsx };
}
