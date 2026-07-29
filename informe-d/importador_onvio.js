// Genera el archivo importador de Onvio (hoja "Asientos"), con la misma estructura que
// el modelo 0.xls. Sección 8 de la especificación.
//
// Todas las columnas y sus tipos se copiaron del modelo real:
//   A Número de asiento   numérico, el mismo en todas las filas
//   B Número de Pase      numérico, correlativo 1,2,3…
//   C Fecha               NÚMERO (serial de Excel, base 1899-12-30) con formato de fecha
//   D Concepto            "Ajuste por Conversión"
//   E Código de cuenta    numérico, el código completo de 9 dígitos
//   F Importe moneda local   0  (el asiento no toca pesos)
//   G Importe moneda ext.    el ajuste en USD CON SIGNO (positivo Debe MEP, negativo Haber)
//   H..L                  vacías
//
// El importador usa un único importe firmado en G; el export de la Grilla, en cambio,
// separa Debe MEP / Haber MEP. La equivalencia es G = DebeMEP − HaberMEP.

const HEADERS_IMPORTADOR = [
  "Número de asiento", "Número de Pase", "Fecha", "Concepto",
  "Código de cuenta", "Importe en moneda local",
  "Importe en moneda ext.present.", "Leyenda",
  "Código de centro de costos", "Porcentaje de distribución",
  "Imp.mon.local dist.C.Costos", "Imp.mon.present.dist.C.Costos",
];

// Formatos tomados del modelo, para que el archivo se vea igual al que usan hoy.
const FMT_FECHA = "m/d/yy";
const FMT_IMPORTE = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';

// serial = días desde 1899-12-30. Verificado: 2026-06-30 -> 46203.
function serialExcel(fecha) {
  const d = fecha instanceof Date ? fecha : new Date(fecha + "T00:00:00");
  const base = Date.UTC(1899, 11, 30);
  const x = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((x - base) / 86400000);
}

// Devuelve la matriz de celdas del importador, ya lista para escribir o para revisar.
function filasImportador(asiento, params) {
  const serial = serialExcel(params.periodoFin);
  const concepto = params.concepto || "Ajuste por Conversión";
  const numero = params.numeroAsiento === undefined || params.numeroAsiento === null
    ? 1 : params.numeroAsiento;
  return (asiento || []).map((l, i) => ({
    numeroAsiento: numero,
    numeroPase: i + 1,
    fechaSerial: serial,
    concepto,
    codigo: l.codigo,
    importeLocal: 0,
    importeExtranjero: l.ajusteUsd,
  }));
}

// Construye el libro con SheetJS. `XLSX` llega como global en el navegador.
function construirLibroImportador(asiento, params) {
  const X = typeof XLSX !== "undefined" ? XLSX : require("../informe-a/vendor/xlsx.full.min.js");
  const filas = filasImportador(asiento, params);

  const ws = {};
  const escribir = (col, fila, celda) => { ws[X.utils.encode_cell({ c: col, r: fila }) ] = celda; };

  HEADERS_IMPORTADOR.forEach((h, c) => escribir(c, 0, { t: "s", v: h }));

  filas.forEach((f, i) => {
    const r = i + 1;
    escribir(0, r, { t: "n", v: f.numeroAsiento });
    escribir(1, r, { t: "n", v: f.numeroPase });
    escribir(2, r, { t: "n", v: f.fechaSerial, z: FMT_FECHA });
    escribir(3, r, { t: "s", v: f.concepto });
    // el código va como número, igual que en el modelo (son 9 dígitos sin ceros a la izquierda)
    escribir(4, r, { t: "n", v: Number(f.codigo) });
    escribir(5, r, { t: "n", v: f.importeLocal, z: FMT_IMPORTE });
    escribir(6, r, { t: "n", v: f.importeExtranjero, z: FMT_IMPORTE });
    // H..L quedan vacías a propósito
  });

  ws["!ref"] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: HEADERS_IMPORTADOR.length - 1, r: filas.length } });
  ws["!cols"] = [{ wch: 17 }, { wch: 15 }, { wch: 11 }, { wch: 22 }, { wch: 16 },
                 { wch: 22 }, { wch: 28 }, { wch: 10 }, { wch: 22 }, { wch: 22 }, { wch: 26 }, { wch: 28 }];

  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws, "Asientos");
  return wb;
}

// tipo: "xls" (igual al modelo) o "xlsx".
function escribirImportador(asiento, params, tipo) {
  const X = typeof XLSX !== "undefined" ? XLSX : require("../informe-a/vendor/xlsx.full.min.js");
  const wb = construirLibroImportador(asiento, params);
  return X.write(wb, { bookType: tipo || "xls", type: "array" });
}

if (typeof module !== "undefined") {
  module.exports = {
    HEADERS_IMPORTADOR, serialExcel, filasImportador,
    construirLibroImportador, escribirImportador,
  };
}
