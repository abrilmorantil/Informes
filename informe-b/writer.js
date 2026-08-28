// En el navegador, ExcelJS ya viene cargado como variable global por el
// script del CDN. En Node (para test) se toma vía require.
if (typeof ExcelJS === "undefined" && typeof require !== "undefined") {
  var ExcelJS = require("exceljs");
}

// `saldosAnteriores` es {codigo: saldo} de la corrida del mes pasado. Si viene, la
// columna "Saldo anterior" sale ya cargada; si falta una cuenta, esa celda queda
// amarilla para completarla a mano, que es como funcionaba todo antes.
async function writeOutputXlsx(lineas, periodo, saldosAnteriores) {
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
    if (l.type !== "parent" && l.type !== "range") {
      ws.getCell(row, 1).value = `${l.code} - ${l.description}`;
    }
    const cli = l.cliente || { code: l.code, description: l.description };
    ws.getCell(row, 2).value = `${cli.code} - ${cli.description}`;
    const cCell = ws.getCell(row, 3);
    const previo = previos[l.code];
    if (typeof previo === "number") cCell.value = round2(previo);
    else cCell.fill = yellowFill;      // sin saldo guardado: se completa a mano
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
    { width: 46 }, { width: 55 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
  ];
  ws.getColumn(1).hidden = true;      // la numeracion real: esta ahi, pero el cliente no la ve

  // --- Segunda hoja: detalle de subcuentas ---
  const ws2 = wb.addWorksheet("Detalle Subcuentas");
  const headers2 = ["Cuenta Madre", "Subcuenta", "Debitos del Mes Dolares", "Creditos del mes Dolares"];
  headers2.forEach((h, j) => {
    const c = ws2.getCell(1, 1 + j);
    c.value = h; c.font = bold; c.border = { bottom: { style: "thin" } };
  });
  let r2 = 2;
  for (const l of visibles) {
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
