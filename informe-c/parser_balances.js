// Lee el export de Onvio "Balance de SyS por Cód. de Cta.", que a diferencia del que
// usa el Informe A trae LAS DOS MONEDAS en la misma planilla.
//
// Estructura: unas filas de título, una fila de encabezados con "Debe/Haber/Saldo"
// por moneda, y después las cuentas agrupadas bajo cuatro capítulos (ACTIVO, PASIVO,
// PATRIMONIO NETO, RESULTADOS). Cierra con "Totales Generales:".

const CAPITULOS = ["ACTIVO", "PASIVO", "PATRIMONIO NETO", "RESULTADOS"];

// El capítulo también se deduce del primer dígito del código. Sirve de control
// cruzado contra los títulos de sección, y es lo que decide el capítulo de una cuenta
// nueva. Verificado: las 163 cuentas del export de junio 2026 cumplen la regla.
const CAPITULO_POR_DIGITO = { 1: "ACTIVO", 2: "PASIVO", 3: "PATRIMONIO NETO", 4: "RESULTADOS" };

// "111010001 - Caja"  ->  { codigo: "111010001", nombre: "Caja" }
const RE_CUENTA = /^\s*(\d{6,})\s*-\s*(.+?)\s*$/;

function normTexto(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/\s+/g, " ").trim().toUpperCase();
}

// Dónde cae realmente cada importe.
//
// No alcanza con buscar el texto del encabezado y leer esa columna: los títulos están
// combinados y el número no siempre cae en la primera columna del rango. En este export
// "Saldo ($)" está en la columna 26 pero el importe aparece en la 27, así que leer la 26
// devolvería vacío —un cero silencioso en todo el balance en pesos—. Y la fila de
// totales lo pone en la 26, o sea que tampoco se puede tomar una fila de muestra.
//
// Por eso se busca el encabezado, se mira hasta dónde llega su rango combinado, y dentro
// de ese rango se elige la columna donde de verdad hay números en las filas de cuenta.
function ubicarColumna(filas, merges, filaEnc, filasDeCuenta, textoBuscado) {
  const enc = filas[filaEnc] || [];
  let col = null;
  for (let c = 0; c < enc.length; c++) {
    if (normTexto(enc[c]) === normTexto(textoBuscado)) { col = c; break; }
  }
  if (col === null) return null;

  const m = (merges || []).find(x => x.s.r === filaEnc && x.s.c <= col && col <= x.e.c);
  const desde = m ? m.s.c : col;
  const hasta = m ? m.e.c : col;

  const votos = {};
  for (const f of filasDeCuenta) {
    for (let c = desde; c <= hasta; c++) {
      if (typeof f[c] === "number") votos[c] = (votos[c] || 0) + 1;
    }
  }
  const ganador = Object.entries(votos).sort((a, b) => b[1] - a[1])[0];
  return ganador ? Number(ganador[0]) : col;
}

function numero(v) {
  return typeof v === "number" ? v : 0;
}

// filas: sheet_to_json(header:1). merges: ws['!merges'].
function parseExportBalances(filas, merges) {
  // la fila de encabezados es la que tiene los títulos de importe
  let filaEnc = null;
  for (let r = 0; r < Math.min(filas.length, 40); r++) {
    const textos = (filas[r] || []).map(normTexto);
    if (textos.some(t => t.startsWith("SALDO (")) && textos.some(t => t.startsWith("DEBE ("))) {
      filaEnc = r;
      break;
    }
  }
  if (filaEnc === null) {
    throw new Error(
      "No encontré la fila de encabezados del export (la que dice \"Debe (u$s)\", " +
      "\"Saldo ($)\", etc.). ¿Es el reporte \"Balance de SyS por Cód. de Cta.\"?"
    );
  }

  const filasDeCuenta = [];
  for (let r = filaEnc + 1; r < filas.length; r++) {
    const a = (filas[r] || [])[0];
    if (a !== null && a !== undefined && RE_CUENTA.test(String(a))) filasDeCuenta.push(filas[r]);
  }
  if (!filasDeCuenta.length) throw new Error("El export no tiene ninguna línea de cuenta.");

  const cols = {
    debe_usd: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe (u$s)"),
    haber_usd: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber (u$s)"),
    saldo_usd: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo (u$s)"),
    debe_ars: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe ($)"),
    haber_ars: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber ($)"),
    saldo_ars: ubicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo ($)"),
  };
  const faltan = Object.entries(cols).filter(([, v]) => v === null).map(([k]) => k);
  if (faltan.length) {
    throw new Error(
      `No pude ubicar en el export las columnas: ${faltan.join(", ")}. ` +
      `El reporte tiene que traer las dos monedas.`
    );
  }

  // recorrido en orden: los títulos de sección marcan el capítulo de lo que sigue
  const cuentas = [];
  const discrepancias = [];
  let capituloActual = null;

  for (let r = filaEnc + 1; r < filas.length; r++) {
    const fila = filas[r] || [];
    const a = fila[0];
    if (a === null || a === undefined) continue;
    const texto = String(a).trim();
    if (!texto) continue;

    const comoCapitulo = normTexto(texto);
    if (CAPITULOS.includes(comoCapitulo)) { capituloActual = comoCapitulo; continue; }

    const m = RE_CUENTA.exec(texto);
    if (!m) continue;                       // "Totales Generales:", pies de página

    const codigo = m[1];
    const porDigito = CAPITULO_POR_DIGITO[codigo[0]] || null;
    if (capituloActual && porDigito && capituloActual !== porDigito) {
      discrepancias.push({ codigo, nombre: m[2], seccion: capituloActual, porDigito });
    }

    cuentas.push({
      codigo,
      nombre: m[2],
      capitulo: capituloActual || porDigito,
      debe_usd: numero(fila[cols.debe_usd]),
      haber_usd: numero(fila[cols.haber_usd]),
      saldo_usd: numero(fila[cols.saldo_usd]),
      debe_ars: numero(fila[cols.debe_ars]),
      haber_ars: numero(fila[cols.haber_ars]),
      saldo_ars: numero(fila[cols.saldo_ars]),
    });
  }

  return {
    cuentas,
    columnas: cols,
    filaEncabezados: filaEnc,
    // cuentas cuyo código no coincide con la sección en la que Onvio las puso: no se
    // corrige solo, se avisa, porque significa que algo cambió en el plan de cuentas
    discrepanciasCapitulo: discrepancias,
    totales: {
      saldo_usd: cuentas.reduce((s, c) => s + c.saldo_usd, 0),
      saldo_ars: cuentas.reduce((s, c) => s + c.saldo_ars, 0),
    },
  };
}

// El código de una cuenta que todavía no está en ningún maestro.
function capituloDeCodigo(codigo) {
  return CAPITULO_POR_DIGITO[String(codigo).trim()[0]] || null;
}

if (typeof module !== "undefined") {
  module.exports = {
    parseExportBalances, capituloDeCodigo, ubicarColumna,
    CAPITULOS, CAPITULO_POR_DIGITO,
  };
}
