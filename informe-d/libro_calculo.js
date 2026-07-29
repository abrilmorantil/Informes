// Genera el libro de trabajo con la misma estructura que
// Prototipo_Ajuste_Dif_Cambio_06-2026.xlsx: cuatro hojas —Parámetros, Resumen y notas,
// Cálculo e Importador—, para revisar el asiento y archivarlo.
//
// Es distinto del archivo que se le da a Onvio: ese va aparte, limpio y con los
// encabezados en la primera fila (ver importador_onvio.js). Acá la hoja "Importador" es
// una previsualización, con su título arriba y la fila de control de cierre abajo.

const FMT_MONEDA = '#,##0.00;\\(#,##0.00\\);\\-';
const FMT_TC = "#,##0.0000";
const FMT_FECHA_LARGA = "dd/mm/yyyy";
const FMT_ENTERO = "0";
const FMT_UMBRAL = "0.000";

const COLS_CALCULO = [
  "Código", "Denominación", "Sección", "Tipo", "Saldo Pesos ($)", "Saldo USD Libros",
  "T.C. Aplicado", "USD Teórico (pesos/TC)", "Ajuste USD", "¿Publica?", "¿Revisar?",
];

function _X() {
  return typeof XLSX !== "undefined" ? XLSX : require("../informe-a/vendor/xlsx.full.min.js");
}

// Ayuda para escribir celdas sueltas en una hoja que se arma a mano.
function _hoja() {
  const X = _X();
  const ws = {};
  let maxC = 0, maxR = 0;
  return {
    ws,
    set(c, r, celda) {
      if (celda === null || celda === undefined) return;
      ws[X.utils.encode_cell({ c, r })] = celda;
      if (c > maxC) maxC = c;
      if (r > maxR) maxR = r;
    },
    cerrar(anchos) {
      ws["!ref"] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: maxC, r: maxR } });
      if (anchos) ws["!cols"] = anchos.map(w => ({ wch: w }));
      return ws;
    },
  };
}

const txt = (v) => ({ t: "s", v: String(v) });
const num = (v, z) => (z ? { t: "n", v: Number(v) || 0, z } : { t: "n", v: Number(v) || 0 });

// ---------------------------------------------------------------- Parámetros
function hojaParametros(params, cfg, titulo) {
  const h = _hoja();
  h.set(1, 1, txt("Ajuste por Conversión – Diferencia de Cambio USD"));
  h.set(1, 2, txt(titulo || "Southern Copper Argentina S.R.L."));
  h.set(1, 4, txt("Fecha de cierre"));            h.set(2, 4, num(serialExcel(params.periodoFin), FMT_FECHA_LARGA));
  h.set(1, 5, txt("T.C. COMPRA (aplica al ACTIVO)")); h.set(2, 5, num(params.tcCompra, FMT_TC));
  h.set(1, 6, txt("T.C. VENTA (aplica al PASIVO)"));  h.set(2, 6, num(params.tcVenta, FMT_TC));
  h.set(1, 7, txt("Concepto del asiento"));       h.set(2, 7, txt(params.concepto || cfg.concepto));
  h.set(1, 8, txt("Número de asiento"));          h.set(2, 8, num(params.numeroAsiento, FMT_ENTERO));
  h.set(1, 9, txt("Cuenta de balanceo (resultado)")); h.set(2, 9, txt(cfg.cuentaBalanceo));
  h.set(1, 10, txt("Umbral materialidad (USD)"));  h.set(2, 10, num(cfg.materialidad, FMT_UMBRAL));
  h.set(1, 13, txt("Los importes salen del SyS del período; los tipos de cambio son los que se cargaron en la app."));
  return h.cerrar([3, 42, 26]);
}

// ---------------------------------------------------------------- Resumen y notas
function hojaResumen(asiento, revisadas, params, cfg) {
  const h = _hoja();
  const balanceo = asiento.find(l => l.esBalanceo);
  const publicadas = asiento.filter(l => !l.esBalanceo);
  const suma = redondear2(asiento.reduce((a, l) => a + l.ajusteUsd, 0));
  const [y, m] = String(params.periodoFin).split("-");

  h.set(1, 1, txt(`Resumen del asiento — ${m}/${y}`));
  h.set(1, 3, txt("Líneas que publican (cuentas)")); h.set(2, 3, num(publicadas.length, FMT_ENTERO));
  h.set(1, 4, txt(`Ajuste a Diferencia de Cambio USD (${cfg.cuentaBalanceo})`)); h.set(2, 4, num(balanceo ? balanceo.ajusteUsd : 0, FMT_MONEDA));
  h.set(1, 5, txt("Control de cierre (debe ser 0)")); h.set(2, 5, num(suma, FMT_MONEDA));
  h.set(1, 6, txt("Cuentas marcadas para revisar")); h.set(2, 6, num((revisadas || []).length, FMT_ENTERO));

  const notas = [
    "Cómo se arma (resumen):",
    `• Cuentas MONETARIAS del ACTIVO/PASIVO. No monetarias excluidas: prefijos ${(cfg.prefijosNoMonetarios || []).join(" / ")} ` +
      `(bienes de uso, depreciaciones y cargos diferidos)`,
    `  y las puntuales ${(cfg.noMonetariasExactas || []).join(" / ") || "—"}.`,
    "• ACTIVO usa T.C. COMPRA; PASIVO usa T.C. VENTA (los dos están en Parámetros).",
    "• USD teórico = pesos/TC (proveedores sin redondear el cociente; el resto redondeado a 2).",
    "• Ajuste = USD teórico − USD en libros. Publica si |ajuste| ≥ 0,01.",
    `• Contrapartida por el neto → ${cfg.cuentaBalanceo} ${cfg.denomBalanceo}. El asiento cierra en 0.`,
    "",
    "Salvaguarda (columna ¿Revisar? en Cálculo):",
    `• Marca una cuenta si el ajuste supera el ${Math.round((cfg.umbralRatio || 0) * 100)}% del saldo USD Y los ${cfg.umbralAbs} USD`,
    "  (es el perfil de una cuenta no monetaria mal clasificada). Ninguna entra al asiento sin confirmación.",
    "",
    "Los saldos en pesos y en dólares salen del SyS del período, sin el ajuste aplicado.",
  ];
  notas.forEach((t, i) => { if (t) h.set(1, 11 + i, txt(t)); });
  return h.cerrar([3, 100, 22]);
}

// ---------------------------------------------------------------- Cálculo
function hojaCalculo(asiento, revisadasExcluidas, cfg) {
  const h = _hoja();
  h.set(0, 0, txt("HOJA DE CÁLCULO — el USD teórico sale del saldo en pesos y el T.C. de cierre; el ajuste es la diferencia contra los libros."));
  COLS_CALCULO.forEach((t, c) => h.set(c, 1, txt(t)));

  const tipoDe = (l) => l.esBalanceo ? "Balanceo" : (l.esProveedor ? "Proveedor" : "General");
  let r = 2;
  const escribir = (l, publica, revisar) => {
    h.set(0, r, txt(l.codigo));
    h.set(1, r, txt(l.denominacion));
    h.set(2, r, txt(l.seccion));
    h.set(3, r, txt(tipoDe(l)));
    if (!l.esBalanceo) {
      h.set(4, r, num(l.saldoPesos, FMT_MONEDA));
      h.set(5, r, num(l.usdLibros, FMT_MONEDA));
      h.set(6, r, num(l.tcAplicado, FMT_TC));
      h.set(7, r, num(l.usdTeorico, FMT_MONEDA));
    }
    h.set(8, r, num(l.ajusteUsd, FMT_MONEDA));
    h.set(9, r, txt(publica));
    if (revisar) h.set(10, r, txt(revisar));
    r++;
  };

  asiento.filter(l => !l.esBalanceo).forEach(l => escribir(l, "SÍ", l.motivo ? "SÍ (confirmada)" : ""));
  (revisadasExcluidas || []).forEach(l => escribir(l, "NO", "SÍ (excluida)"));
  const balanceo = asiento.find(l => l.esBalanceo);
  if (balanceo) escribir(balanceo, "SÍ", "");

  r++;   // una fila en blanco antes del control, como el prototipo
  const suma = redondear2(asiento.reduce((a, l) => a + l.ajusteUsd, 0));
  h.set(7, r, txt("CONTROL — el asiento debe cerrar en 0:"));
  h.set(8, r, num(suma, FMT_MONEDA));
  h.set(9, r, txt(Math.abs(suma) < 0.005 ? "OK ✓" : "NO CIERRA"));

  return h.cerrar([12, 44, 14, 11, 18, 17, 13, 21, 15, 11, 16]);
}

// ---------------------------------------------------------------- Importador (vista)
function hojaImportadorVista(asiento, params) {
  const h = _hoja();
  h.set(0, 0, txt('ARCHIVO IMPORTADOR — formato Onvio "Asientos". Pesos = 0, USD firmado (+ Debe MEP / − Haber MEP).'));
  HEADERS_IMPORTADOR.forEach((t, c) => h.set(c, 1, txt(t)));

  const filas = filasImportador(asiento, params);
  filas.forEach((f, i) => {
    const r = i + 2;
    h.set(0, r, num(f.numeroAsiento, FMT_ENTERO));
    h.set(1, r, num(f.numeroPase, FMT_ENTERO));
    h.set(2, r, num(f.fechaSerial, FMT_FECHA_LARGA));
    h.set(3, r, txt(f.concepto));
    h.set(4, r, txt(f.codigo));
    h.set(5, r, num(f.importeLocal, FMT_MONEDA));
    h.set(6, r, num(f.importeExtranjero, FMT_MONEDA));
  });

  const r = filas.length + 2;
  const suma = redondear2(filas.reduce((a, f) => a + f.importeExtranjero, 0));
  h.set(5, r, txt("Σ ="));
  h.set(6, r, num(suma, FMT_MONEDA));
  h.set(7, r, txt(Math.abs(suma) < 0.005 ? "cierra en 0 ✓" : "NO CIERRA"));

  return h.cerrar([17, 15, 12, 22, 16, 22, 28, 16, 24, 22, 26, 28]);
}

// ---------------------------------------------------------------- el libro
function construirLibroCalculo({ asiento, revisadas, revisadasExcluidas, params, cfg, titulo }) {
  const X = _X();
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, hojaParametros(params, cfg, titulo), "Parámetros");
  X.utils.book_append_sheet(wb, hojaResumen(asiento, revisadas, params, cfg), "Resumen y notas");
  X.utils.book_append_sheet(wb, hojaCalculo(asiento, revisadasExcluidas, cfg), "Cálculo");
  X.utils.book_append_sheet(wb, hojaImportadorVista(asiento, params), "Importador");
  return wb;
}

function escribirLibroCalculo(datos) {
  const X = _X();
  return X.write(construirLibroCalculo(datos), { bookType: "xlsx", type: "array" });
}

if (typeof module !== "undefined") {
  const m = require("./motor_difcambio.js");
  const i = require("./importador_onvio.js");
  global.redondear2 = m.redondear2;
  global.serialExcel = i.serialExcel;
  global.filasImportador = i.filasImportador;
  global.HEADERS_IMPORTADOR = i.HEADERS_IMPORTADOR;
  module.exports = {
    COLS_CALCULO, construirLibroCalculo, escribirLibroCalculo,
    hojaParametros, hojaResumen, hojaCalculo, hojaImportadorVista,
  };
}
