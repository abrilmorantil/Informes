// Núcleo de BALCOMPROBDOLARES portado a JS puro (corre en el navegador).
// Misma lógica que app.py, sin Flask ni disco: todo en memoria.

const CATEGORY_BY_FIRST_DIGIT = { "1": "ACTIVO", "2": "PASIVO", "3": "CAPITAL Y PATRIMONIO", "4": "RESULTADOS" };
const CATEGORIES = ["ACTIVO", "PASIVO", "CAPITAL Y PATRIMONIO", "RESULTADOS"];
const KNOWN_CATEGORIES = ["ACTIVO", "PASIVO", "CAPITAL Y PATRIMONIO", "PATRIMONIO NETO", "RESULTADOS"];
const EPSILON = 0.05;
const CODE_PATTERN = /^\s*([0-9]{5,})\s*-\s*(.+?)\s*$/;

function num(v) {
  if (v === null || v === undefined || v === "") return 0.0;
  const n = typeof v === "number" ? v : parseFloat(v);
  return isNaN(n) ? 0.0 : n;
}

// rows: array de arrays (como XLSX.utils.sheet_to_json(ws, {header:1}))
function parseSiseExport(rows) {
  let colDebe = null, colHaber = null, colSaldo = null, headerRow = null;
  const maxScan = Math.min(40, rows.length);
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i] || [];
    for (let j = 0; j < row.length; j++) {
      const v = row[j];
      if (typeof v === "string") {
        if (v.includes("Debe (u$s)")) { colDebe = j; headerRow = i; }
        else if (v.includes("Haber (u$s)")) { colHaber = j; }
        else if (v.includes("Saldo (u$s)")) { colSaldo = j; }
      }
    }
    if (colDebe !== null && colHaber !== null && colSaldo !== null) break;
  }
  if (colDebe === null) {
    throw new Error("No pude encontrar las columnas 'Debe (u$s)/Haber (u$s)/Saldo (u$s)' en el archivo. ¿Es un export de SISE en dólares?");
  }

  const cuentas = {};
  let control = null;
  const categoryTotals = {};
  let currentCategory = null;

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const label = row[0];
    const marker = row[7];

    if (typeof marker === "string" && marker.trim() === "Totales" && currentCategory) {
      categoryTotals[currentCategory] = { debe: num(row[colDebe]), haber: num(row[colHaber]) };
      continue;
    }

    if (typeof label !== "string" || !label.trim()) continue;

    if (label.trim().startsWith("Totales Generales")) {
      control = { debe: num(row[colDebe]), haber: num(row[colHaber]), saldo: num(row[colSaldo]) };
      continue;
    }

    if (KNOWN_CATEGORIES.includes(label.trim())) {
      currentCategory = label.trim();
      continue;
    }

    const m = CODE_PATTERN.exec(label);
    if (!m) continue;
    const code = m[1], desc = m[2];
    // `fila` es la fila del export de donde salio, contando desde 1 como la ve Excel. Con
    // eso el informe puede APUNTAR al export con una formula en vez de copiar el numero:
    // asi los importes los recalcula Excel y no hay que creerle a la app.
    cuentas[code] = { descripcion: desc, debe: num(row[colDebe]), haber: num(row[colHaber]),
                      saldo: num(row[colSaldo]), fila: i + 1 };
  }

  if (control === null) {
    throw new Error("No encontré la fila 'Totales Generales:' en el export.");
  }

  // Las columnas se devuelven porque el writer las necesita para armar las formulas que
  // apuntan a la hoja del export: son las de USD, ubicadas por su encabezado.
  return { cuentas, control, categoryTotals, colDebe, colHaber, colSaldo };
}

function knownCodes(mapping, cuentasSise) {
  const codes = new Set();
  for (const e of mapping) {
    // Una fila declarada "sin cuentas asignadas" no se adueña de su código. Si no, una
    // cuenta de Onvio con ese mismo número quedaría dada por conocida y `findUnmapped` no
    // avisaría: la plata desaparecería sin que nadie lo note. Es el caso de "421170000
    // Gastos Legales", que en Onvio es "Alojamiento Rel. Comunitarias Catamarca".
    if (!e.sin_cuentas) codes.add(e.code);
    (e.aliases || []).forEach(a => codes.add(a));
    (e.children || []).forEach(c => {
      codes.add(c.code);
      (c.aliases || []).forEach(a => codes.add(a));
    });
    if (e.type === "range" && cuentasSise) {
      for (const code of Object.keys(cuentasSise)) {
        if (code.startsWith(e.prefix)) codes.add(code);
      }
    }
  }
  return codes;
}

function lookupCuenta(cuentasSise, entryOrChild) {
  let c = cuentasSise[entryOrChild.code];
  if (c) return c;
  for (const alias of (entryOrChild.aliases || [])) {
    c = cuentasSise[alias];
    if (c) return c;
  }
  return null;
}

// Todas las cuentas del export que no van a parar a ningún renglón del informe, hayan tenido
// movimiento o no. Las que no movieron son las peligrosas: no rompen nada este mes, así que
// nadie las ve, y aparecen recién el mes que se mueven. Pasó con "Gastos en trámites" y
// "Canon", que estuvieron dormidas todo el ejercicio sin que nada las mostrara.
function findUnassigned(cuentasSise, mapping) {
  const conocidos = knownCodes(mapping, cuentasSise);
  const sueltas = [];
  for (const [code, c] of Object.entries(cuentasSise)) {
    if (conocidos.has(code)) continue;
    sueltas.push({
      code, description: c.descripcion, debe: c.debe, haber: c.haber,
      movio: !!(c.debe || c.haber),
      suggested_category: CATEGORY_BY_FIRST_DIGIT[code[0]] || "RESULTADOS",
    });
  }
  // Primero las que movieron, y dentro de cada grupo por importe.
  sueltas.sort((a, b) => (b.movio - a.movio) ||
    ((b.debe + b.haber) - (a.debe + a.haber)) || a.code.localeCompare(b.code));
  return sueltas;
}

// Las que además tuvieron movimiento: ésas sí trancan la descarga, porque su plata tiene que
// entrar en alguna parte para que el informe ate con Onvio. Sale de la misma función para que
// las dos listas no puedan contradecirse.
function findUnmapped(cuentasSise, mapping) {
  return findUnassigned(cuentasSise, mapping).filter(x => x.movio);
}

function findDuplicateCodes(mapping) {
  const occurrences = {};
  const register = (code, where) => {
    if (!occurrences[code]) occurrences[code] = [];
    occurrences[code].push(where);
  };
  for (const e of mapping) {
    register(e.code, `${e.code} - ${e.description} (${e.category})`);
    (e.aliases || []).forEach(a => register(a, `alias de ${e.code} - ${e.description}`));
    (e.children || []).forEach(c => {
      if (c.code === e.code) return; // madre==hija: no es doble conteo real
      register(c.code, `hijo de ${e.code} - ${e.description}`);
      (c.aliases || []).forEach(a => register(a, `alias de hijo ${c.code} (bajo ${e.code})`));
    });
  }
  const dups = {};
  for (const [code, wheres] of Object.entries(occurrences)) {
    if (wheres.length > 1) dups[code] = wheres;
  }
  return dups;
}

// Qué cuentas REALES de Onvio alimentan una fila del informe.
//
// Una fila puede leer más de una: una cuenta madre suma la suya y las de sus hijas, y el
// renglón de proveedores junta todas las que empiezan con su prefijo. Es la misma regla con
// la que `buildBalance` calcula el movimiento del mes — acá se la nombra aparte porque el
// SALDO ANTERIOR tiene que salir de las mismas cuentas y no lo estaba haciendo.
function codigosQueAlimentan(entry, prevBalances) {
  const out = new Set();
  if (entry.sin_cuentas) return out;          // declarada sin cuentas: no lee nada de Onvio
  if (entry.type === "range" && entry.prefix) {
    for (const k of Object.keys(prevBalances || {})) {
      if (String(k).startsWith(entry.prefix)) out.add(String(k));
    }
    return out;
  }
  out.add(String(entry.code));
  for (const h of (entry.children || [])) if (h && h.code) out.add(String(h.code));
  return out;
}

// El saldo anterior de una fila: la suma de los saldos guardados de todas sus cuentas.
//
// Antes era `prevBalances[entry.code]`, o sea SÓLO el de su propio código. Con eso, al juntar
// dos cuentas en una sola fila —"21409000 Retenciones SUSS a pagar" pasó a ser madre de
// "212020002 Retenciones SUSS a pagar"— el saldo de la hija (−30,97) dejaba de aparecer y el
// balance abría por esa diferencia. Medido en agosto 2026: la columna pasaba de −0,02 a 30,95.
function saldoAnteriorDeLinea(entry, prevBalances) {
  const previos = prevBalances || {};
  let total = 0, conocido = false;
  for (const cod of codigosQueAlimentan(entry, previos)) {
    if (!Object.prototype.hasOwnProperty.call(previos, cod)) continue;
    const v = previos[cod];
    if (typeof v !== "number" || !isFinite(v)) continue;
    total += v;
    conocido = true;
  }
  return { total: Math.round(total * 100) / 100, conocido };
}

function buildBalance(cuentasSise, mapping, prevBalances) {
  const lineas = [];
  let totalDebe = 0, totalHaber = 0;

  for (const entry of mapping) {
    // El saldo anterior de la fila es el de TODAS las cuentas reales que la alimentan, no
    // sólo el de su propio código. Ver `codigosQueAlimentan`.
    const code = entry.code;
    let debe = 0, haber = 0, detalle = [];

    if (entry.type === "parent") {
      const yaSumados = new Set();
      for (const fuente of [entry, ...(entry.children || [])]) {
        if (yaSumados.has(fuente.code)) continue;
        const c = lookupCuenta(cuentasSise, fuente);
        const dF = c ? c.debe : 0, hF = c ? c.haber : 0;
        debe += dF; haber += hF;
        if (fuente === entry) {
          if (dF || hF) {
            detalle.push({ code: fuente.code, description: fuente.description + " (directo)",
                           debe: dF, haber: hF, fila: c ? c.fila : null });
          }
        } else {
          detalle.push({ code: fuente.code, description: fuente.description,
                         debe: dF, haber: hF, fila: c ? c.fila : null });
        }
        yaSumados.add(fuente.code);
      }
    } else if (entry.type === "range") {
      // El rango de proveedores juntaba 118 cuentas en una fila sin dejar rastro de
      // cuales. Ahora cada una queda anotada en `detalle`, que es lo que alimenta la hoja
      // de trazabilidad del archivo.
      for (const [scode, c] of Object.entries(cuentasSise)) {
        if (scode.startsWith(entry.prefix)) {
          debe += c.debe; haber += c.haber;
          detalle.push({ code: scode, description: c.descripcion, debe: c.debe, haber: c.haber, fila: c.fila });
        }
      }
    } else if (entry.sin_cuentas) {
      // Declarada sin cuentas asignadas: no lee NADA de Onvio, queda siempre en cero.
      // Antes la marca era sólo un cartel: la fila igual se llevaba la cuenta que
      // coincidiera con su código y lo único que se salteaba era la trazabilidad. Con eso,
      // "421170000 Gastos Legales" se habría llevado el saldo de "Alojamiento Rel.
      // Comunitarias Catamarca" —los números coinciden de casualidad— y sin dejar rastro.
      debe = 0; haber = 0;
    } else {
      const c = lookupCuenta(cuentasSise, entry);
      debe = c ? c.debe : 0;
      haber = c ? c.haber : 0;
      // Una fila simple se llena con UNA cuenta, pero con las dos numeraciones esa cuenta
      // no tiene por que llamarse igual que la fila: hay que decirlo igual.
      detalle.push({ code: entry.code, description: entry.description, debe, haber,
                     fila: c ? c.fila : null });
    }

    const movimiento = debe - haber;
    const previo = saldoAnteriorDeLinea(entry, prevBalances);
    const saldoAnterior = previo.total;
    const saldoNuevo = saldoAnterior + movimiento;

    lineas.push({
      code, description: entry.description, category: entry.category, type: entry.type,
      // `code`/`description` son la identidad REAL de la cuenta, la que viene del export.
      // `cliente` es lo que se imprime en el informe: el plan de cuentas viejo, que es el
      // que el cliente lee. Los dos codigos conviven porque no hay regla que los relacione
      // —los codigos fueron reasignados, no reformateados— asi que la equivalencia se
      // declara a mano en mapping.json. Si no hay `cliente`, se imprime el code.
      cliente: entry.cliente || null,
      // `saldo_anterior_conocido` es false cuando NINGUNA de las cuentas que alimentan la
      // fila tiene saldo guardado: ahí la celda sale en amarillo para cargarla a mano.
      saldo_anterior_conocido: previo.conocido,
      ocultar_si_cero: !!entry.ocultar_si_cero,
      orden: entry.orden !== undefined ? entry.orden : 999999,
      saldo_anterior: saldoAnterior, debe, haber, movimiento, saldo_nuevo: saldoNuevo, detalle,
    });
    totalDebe += debe; totalHaber += haber;
  }

  const rangoCategoria = {};
  CATEGORIES.forEach((c, i) => { rangoCategoria[c] = i; });
  lineas.sort((a, b) => {
    const ra = rangoCategoria[a.category] ?? 99, rb = rangoCategoria[b.category] ?? 99;
    if (ra !== rb) return ra - rb;
    return a.orden - b.orden;
  });

  return { lineas, totalDebe, totalHaber };
}

function runValidation(lineas, totalDebe, totalHaber, control, unmapped, categoryTotals, duplicates) {
  const totalMovimiento = lineas.reduce((s, l) => s + l.movimiento, 0);
  const diffDebe = control.debe - totalDebe;
  const diffHaber = control.haber - totalHaber;

  const checks = [
    { name: "Movimiento total del mes = 0", detail: "Todo Debe tiene su Haber en algún lado del mes.",
      value: round2(totalMovimiento), ok: Math.abs(totalMovimiento) < EPSILON },
    { name: "Total armado vs. total de control de SISE (Debe)", detail: "Compara la suma de las cuentas mapeadas contra 'Totales Generales' del export.",
      value: round2(diffDebe), ok: Math.abs(diffDebe) < EPSILON },
    { name: "Total armado vs. total de control de SISE (Haber)", detail: "Idem, del lado del Haber.",
      value: round2(diffHaber), ok: Math.abs(diffHaber) < EPSILON },
  ];

  const categoryDiffs = [];
  if (categoryTotals) {
    let grupos;
    if ("PATRIMONIO NETO" in categoryTotals) {
      grupos = { "ACTIVO": ["ACTIVO"], "PASIVO": ["PASIVO"], "PATRIMONIO NETO": ["CAPITAL Y PATRIMONIO"], "RESULTADOS": ["RESULTADOS"] };
    } else {
      grupos = { "ACTIVO": ["ACTIVO"], "PASIVO": ["PASIVO", "CAPITAL Y PATRIMONIO"], "RESULTADOS": ["RESULTADOS"] };
    }
    for (const [siseCat, mapCats] of Object.entries(grupos)) {
      const siseTot = categoryTotals[siseCat];
      if (!siseTot) continue;
      const armadoDebe = lineas.filter(l => mapCats.includes(l.category)).reduce((s, l) => s + l.debe, 0);
      const armadoHaber = lineas.filter(l => mapCats.includes(l.category)).reduce((s, l) => s + l.haber, 0);
      const dDebe = siseTot.debe - armadoDebe, dHaber = siseTot.haber - armadoHaber;
      categoryDiffs.push({ categoria: siseCat, diff_debe: round2(dDebe), diff_haber: round2(dHaber),
        ok: Math.abs(dDebe) < EPSILON && Math.abs(dHaber) < EPSILON });
    }
  }

  const allOk = checks.every(c => c.ok) && unmapped.length === 0 && Object.keys(duplicates).length === 0;
  return { checks, allOk, categoryDiffs };
}

function round2(n) { return Math.round(n * 100) / 100; }

if (typeof module !== "undefined") {
  module.exports = {
    parseSiseExport, knownCodes, lookupCuenta, findUnmapped, findUnassigned, findDuplicateCodes,
    buildBalance, runValidation, CATEGORIES, codigosQueAlimentan, saldoAnteriorDeLinea,
  };
}
