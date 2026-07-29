// Lee el Balance de Sumas y Saldos BIMONETARIO (SyS_<periodo>_sin_ajuste.xls), que es
// el único insumo de datos del asiento de diferencia de cambio.
//
// Ojo, no es el mismo export que usa el Informe C: este trae el saldo en pesos y el
// saldo en dólares EN LIBROS (antes del ajuste), y sus columnas están al revés. Por eso
// los índices se toman de la configuración (`colSaldoPesos` / `colSaldoUsd`) y no se
// dan por sentados: la sección 11 de la especificación lo pide explícitamente.
//
// Layout de referencia (0-based), medido sobre el export de Crystal Reports:
//   A  = 0   "Cuenta - Denominación"  ->  "111010001 - Caja", y los rótulos de sección
//   P  = 15  "Saldo ($)"              ->  saldo en pesos
//   AB = 27  "Saldo (u$s)"            ->  saldo en dólares en libros
//
// El header dice AA pero el valor cae en AB: hay que usar 27, no 26.

const SECCIONES_SYS = ["ACTIVO", "PASIVO", "PATRIMONIO NETO", "RESULTADOS"];
const RE_CUENTA_SYS = /^(\d{9})\s*-\s*(.*)$/;

const COLS_SYS_DEFECTO = { cuenta: 0, saldoPesos: 15, saldoUsd: 27 };

function numeroSys(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  // por si el export trae los importes como texto con separadores
  const t = String(v).replace(/\s/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
  const n = parseFloat(t);
  return isFinite(n) ? n : 0;
}

function textoSys(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return v.result !== undefined ? textoSys(v.result) : "";
  return String(v);
}

// filas: matriz de la hoja (sheet_to_json con header:1). Devuelve CuentaSaldo[].
function parseSysBimonetario(filas, cols = COLS_SYS_DEFECTO) {
  const c = { ...COLS_SYS_DEFECTO, ...(cols || {}) };
  const cuentas = [];
  const seccionesVistas = [];
  let seccion = null;

  for (const fila of filas || []) {
    if (!fila) continue;
    const a = textoSys(fila[c.cuenta]).trim();
    if (!a) continue;

    const comoSeccion = a.toUpperCase().replace(/\s+/g, " ").trim();
    if (SECCIONES_SYS.indexOf(comoSeccion) >= 0) {
      seccion = comoSeccion;
      if (seccionesVistas.indexOf(seccion) < 0) seccionesVistas.push(seccion);
      continue;
    }

    const m = RE_CUENTA_SYS.exec(a);
    if (!m || !seccion) continue;

    cuentas.push({
      codigo: m[1],
      denominacion: m[2].trim(),
      seccion,
      saldoPesos: numeroSys(fila[c.saldoPesos]),
      saldoUsd: numeroSys(fila[c.saldoUsd]),
    });
  }

  return { cuentas, secciones: seccionesVistas, columnas: c };
}

// Comprobación de cordura del parseo. Si las columnas vinieran cambiadas —el caso que
// avisa la sección 11 de la spec— los números salen absurdos: en este cliente el saldo
// en pesos es del orden de mil veces el de dólares. Se avisa en vez de calcular mal.
function revisarParseoSys(cuentas) {
  const avisos = [];
  const conAmbos = cuentas.filter(x => x.saldoPesos !== 0 && x.saldoUsd !== 0);
  if (!cuentas.length) {
    avisos.push("No se reconoció ninguna cuenta. ¿Es el Balance de Sumas y Saldos bimonetario?");
    return avisos;
  }
  if (!conAmbos.length) {
    avisos.push("Ninguna cuenta tiene saldo en las dos monedas a la vez: puede que las columnas " +
                "de importe no sean las esperadas (pesos en P, dólares en AB).");
    return avisos;
  }
  const ratios = conAmbos.map(x => Math.abs(x.saldoPesos / x.saldoUsd)).sort((a, b) => a - b);
  const mediana = ratios[Math.floor(ratios.length / 2)];
  if (mediana < 10) {
    avisos.push(`Los saldos en pesos y en dólares están en el mismo orden de magnitud ` +
                `(relación típica ${mediana.toFixed(2)}). Puede que las columnas estén invertidas.`);
  }
  return avisos;
}

if (typeof module !== "undefined") {
  module.exports = {
    parseSysBimonetario, revisarParseoSys, numeroSys,
    SECCIONES_SYS, RE_CUENTA_SYS, COLS_SYS_DEFECTO,
  };
}
