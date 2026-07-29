// Motor del asiento "Ajuste por Conversión" (diferencia de cambio bimonetaria ARS/USD).
//
// Port del motor de referencia validado contra el asiento real Nº 2630 (06-2026), que
// reproduce sus 89 líneas materiales al centavo. La lógica NO se improvisa: sale de
// ESPEC_Ajuste_Diferencia_Cambio.md, que es la fuente de verdad.
//
// En una línea: para cada cuenta MONETARIA del ACTIVO/PASIVO,
//   ajuste_usd = (saldo_pesos / TC_cierre) − saldo_usd_en_libros
// con TC compra para el activo y TC venta para el pasivo, redondeando el cociente salvo
// en proveedores; se publican las de |ajuste| ≥ 0,01; y la contrapartida por el neto va
// a 423050000, que además absorbe el residuo de redondeo y hace cerrar el asiento en 0.

const CONFIG_DIFCAMBIO_DEFECTO = {
  // Cuentas NO monetarias: van a costo histórico, su diferencia peso/USD es real y no
  // debe ajustarse. No se auto-excluyen por materialidad, hay que excluirlas por criterio.
  prefijosNoMonetarios: ["1240", "1250"],   // Bienes de Uso + Dep. Acum., Cargos Diferidos
  noMonetariasExactas: [
    "114010016",   // Impuesto Crédito Diferido
    "114050005",   // Seguros a Devengar (pago anticipado)
    "211050000",   // Previsión AID (impuesto diferido pasivo)
  ],
  // Proveedores: el cociente NO se redondea (replica el subdiario, C = B/TC, E = C − D).
  prefijoProveedores: "21101",
  cuentaBalanceo: "423050000",
  denomBalanceo: "Diferencia de Cambio USD",
  materialidad: 0.005,
  concepto: "Ajuste por Conversión",
  // Cuentas monetarias ya confirmadas por la usuaria: no se vuelven a marcar.
  monetariasConfirmadas: [],
  // Salvaguarda (sección 12.3). Calibrado con datos reales: una no monetaria mal
  // incluida da ratio ~0,95 y una monetaria legítima ~0,05, así que 0,30 las separa sin
  // ambigüedad. Se exigen LAS DOS condiciones para no marcar cuentas de saldo chico
  // (Caja ajusta el 100% de su saldo pero son 24 USD: no se marca).
  umbralRatio: 0.30,
  umbralAbs: 500,
};

// Redondeo a 2 decimales con la semántica de REDONDEAR de Excel (medio hacia arriba en
// valor absoluto), corrigiendo antes el error de representación del punto flotante:
// 2.675*100 da 267.49999999999997 y redondearía mal.
function redondear2(x) {
  if (!isFinite(x) || x === 0) return 0;
  const escalado = Number((x * 100).toPrecision(15));
  return (escalado < 0 ? -1 : 1) * Math.round(Math.abs(escalado)) / 100;
}

function configDifCambio(parcial) {
  const c = { ...CONFIG_DIFCAMBIO_DEFECTO, ...(parcial || {}) };
  c.prefijosNoMonetarios = [...(c.prefijosNoMonetarios || [])];
  c.noMonetariasExactas = [...new Set(c.noMonetariasExactas || [])];
  c.monetariasConfirmadas = [...new Set(c.monetariasConfirmadas || [])];
  return c;
}

// Sección 5.1. Sólo el ACTIVO y el PASIVO entran; Patrimonio Neto y Resultados quedan
// enteros afuera (la cuenta de balanceo se agrega aparte, no sale de acá).
function esMonetaria(cuenta, cfg) {
  if (cuenta.seccion !== "ACTIVO" && cuenta.seccion !== "PASIVO") return false;
  const cod = String(cuenta.codigo);
  if (cfg.prefijosNoMonetarios.some(p => cod.slice(0, String(p).length) === String(p))) return false;
  if (cfg.noMonetariasExactas.indexOf(cod) >= 0) return false;
  return true;
}

// ACTIVO → T.C. COMPRA ; PASIVO → T.C. VENTA
function tcDe(cuenta, params) {
  return cuenta.seccion === "ACTIVO" ? params.tcCompra : params.tcVenta;
}

// Secciones 5.2 y 6.
function calcularLineas(cuentas, params, cfg) {
  const c = configDifCambio(cfg);
  const lineas = [];
  for (const cuenta of cuentas) {
    if (!esMonetaria(cuenta, c)) continue;
    const tc = tcDe(cuenta, params);
    if (!tc) throw new Error(`Falta el tipo de cambio para la sección ${cuenta.seccion}.`);

    const cociente = cuenta.saldoPesos / tc;
    const esProveedor = String(cuenta.codigo).indexOf(c.prefijoProveedores) === 0;
    const usdTeorico = esProveedor ? cociente : redondear2(cociente);
    const ajuste = usdTeorico - cuenta.saldoUsd;

    if (Math.abs(redondear2(ajuste)) >= c.materialidad) {
      lineas.push({
        codigo: cuenta.codigo,
        denominacion: cuenta.denominacion,
        seccion: cuenta.seccion,
        tcAplicado: tc,
        saldoPesos: cuenta.saldoPesos,
        usdTeorico,
        usdLibros: cuenta.saldoUsd,
        ajusteUsd: ajuste,
        esProveedor,
      });
    }
  }
  return lineas;
}

// Sección 12.3: una no monetaria mal incluida es a la vez desproporcionada y grande.
function motivoSospecha(linea, cfg) {
  const c = configDifCambio(cfg);
  if (c.monetariasConfirmadas.indexOf(String(linea.codigo)) >= 0) return null;
  const desproporcion = Math.abs(linea.usdLibros) > 1e-9 &&
    Math.abs(linea.ajusteUsd) > c.umbralRatio * Math.abs(linea.usdLibros);
  const material = Math.abs(linea.ajusteUsd) > c.umbralAbs;
  if (desproporcion && material) return "desproporcion";
  return null;
}

// Fase 1: calcula y separa, pero NO genera el importador.
function calcularConRevision(cuentas, params, cfg) {
  const c = configDifCambio(cfg);
  const lineas = calcularLineas(cuentas, params, c);
  const lineasOk = [], lineasARevisar = [];
  for (const l of lineas) {
    const motivo = motivoSospecha(l, c);
    if (motivo) lineasARevisar.push({ ...l, motivo });
    else lineasOk.push(l);
  }
  return { lineasOk, lineasARevisar };
}

// Fase 2: decisiones = { codigo: "incluir" | "excluir" }.
// Sin confirmación explícita NO entra: el default es excluir (regla de oro de la spec).
function aplicarDecisiones(lineasARevisar, decisiones) {
  const incluidas = [], aExcluir = [], aConfirmar = [];
  for (const l of lineasARevisar || []) {
    const d = (decisiones || {})[l.codigo] || "excluir";
    if (d === "incluir") { incluidas.push(l); aConfirmar.push(l.codigo); }
    else aExcluir.push(l.codigo);
  }
  return {
    incluidas,
    cambiosConfig: { noMonetariasExactasAgregar: aExcluir, monetariasConfirmadasAgregar: aConfirmar },
  };
}

function aplicarCambiosConfig(cfg, cambios) {
  const c = configDifCambio(cfg);
  const nueva = { ...c };
  nueva.noMonetariasExactas = [...new Set([...c.noMonetariasExactas, ...((cambios || {}).noMonetariasExactasAgregar || [])])].sort();
  nueva.monetariasConfirmadas = [...new Set([...c.monetariasConfirmadas, ...((cambios || {}).monetariasConfirmadasAgregar || [])])].sort();
  return nueva;
}

// Sección 7: cada ajuste se publica redondeado a 2, y la cuenta de resultado toma el
// neto cambiado de signo. Así Σ = 0 por construcción y esa cuenta absorbe el residuo.
function armarAsiento(lineas, cfg) {
  const c = configDifCambio(cfg);
  const publicadas = (lineas || []).map(l => ({ ...l, ajusteUsd: redondear2(l.ajusteUsd) }));
  const neto = redondear2(publicadas.reduce((a, l) => a + l.ajusteUsd, 0));
  publicadas.push({
    codigo: c.cuentaBalanceo,
    denominacion: c.denomBalanceo,
    seccion: "RESULTADOS",
    tcAplicado: 0,
    saldoPesos: 0,
    usdTeorico: 0,
    usdLibros: 0,
    ajusteUsd: redondear2(-neto),
    esBalanceo: true,
  });
  return publicadas;
}

// Red de contención independiente de todo lo anterior (secciones 11 y 12).
function verificarCierre(asiento) {
  const suma = redondear2((asiento || []).reduce((a, l) => a + l.ajusteUsd, 0));
  return { cierra: Math.abs(suma) < 0.005, suma };
}

if (typeof module !== "undefined") {
  module.exports = {
    CONFIG_DIFCAMBIO_DEFECTO, configDifCambio, redondear2,
    esMonetaria, tcDe, calcularLineas, motivoSospecha,
    calcularConRevision, aplicarDecisiones, aplicarCambiosConfig,
    armarAsiento, verificarCierre,
  };
}
