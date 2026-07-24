// Port de parser_onvio.py. Lee el export de Onvio "Sumas y Saldos por Centro
// de Costos" a partir de las filas ya leídas por SheetJS (array de arrays).
//
// - Fila con "N - Nombre" y N de <=3 dígitos -> encabezado de centro de costo
//   (se ignoran "Sin Asignar" y "Totales").
// - Fila con "N - Nombre" y N de >=6 dígitos -> cuenta contable, se asocia al
//   centro de costo vigente.

function toFloat(v) {
  if (v === null || v === undefined || v === "") return 0.0;
  const n = typeof v === "number" ? v : parseFloat(v);
  return isNaN(n) ? 0.0 : n;
}

// Las columnas en dólares se ubican por su encabezado y nunca por posición, para
// que un cambio de orden en el export de Onvio no haga tomar las columnas en pesos.
function detectarColumnasUsd(rows) {
  const limite = Math.min(15, rows.length);
  for (let i = 0; i < limite; i++) {
    const fila = rows[i] || [];
    let debe = null, haber = null, saldo = null;
    for (let j = 0; j < fila.length; j++) {
      const v = fila[j];
      if (typeof v !== "string") continue;
      const t = v.trim().toLowerCase().replace(/\s/g, "");
      if (t === "debe(u$s)") debe = j;
      else if (t === "haber(u$s)") haber = j;
      else if (t === "saldo(u$s)") saldo = j;
    }
    if (debe !== null && haber !== null && saldo !== null) {
      return { debe, haber, saldo };
    }
  }
  throw new Error(
    "No encontré las columnas 'Debe (u$s)' / 'Haber (u$s)' / 'Saldo (u$s)' en el " +
    "archivo. Puede que el formato del export de Onvio haya cambiado. Revisá el " +
    "archivo antes de seguir (NO se tomaron valores en pesos por error)."
  );
}

function parseOnvioExport(rows) {
  const { debe: colDebe, haber: colHaber, saldo: colSaldo } = detectarColumnasUsd(rows);

  const filas = [];
  let ccCodigo = null;
  let ccNombre = null;

  for (let i = 0; i < rows.length; i++) {
    const fila = rows[i] || [];
    let celda = fila[0];
    if (typeof celda !== "string") continue;
    celda = celda.trim();
    if (!celda) continue;

    const pos = celda.indexOf(" - ");
    if (pos === -1) continue;

    const prefijo = celda.slice(0, pos).trim();
    const resto = celda.slice(pos + 3).trim();
    if (!/^\d+$/.test(prefijo)) continue;

    if (prefijo.length <= 3) {
      const r = resto.toLowerCase();
      if (r.includes("sin asignar") || r.includes("totales")) {
        ccCodigo = null;
        ccNombre = null;
      } else {
        ccCodigo = prefijo;
        ccNombre = resto;
      }
    } else if (prefijo.length >= 6 && ccCodigo !== null) {
      filas.push({
        cc_codigo: ccCodigo,
        cc_nombre_onvio: ccNombre,
        cuenta_codigo: prefijo,
        cuenta_label: resto,
        debe: toFloat(fila[colDebe]),
        haber: toFloat(fila[colHaber]),
        saldo: toFloat(fila[colSaldo]),
      });
    }
  }

  return filas;
}

if (typeof module !== "undefined") {
  module.exports = { parseOnvioExport, detectarColumnasUsd };
}
