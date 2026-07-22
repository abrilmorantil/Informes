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
    cuentas[code] = { descripcion: desc, debe: num(row[colDebe]), haber: num(row[colHaber]), saldo: num(row[colSaldo]) };
  }

  if (control === null) {
    throw new Error("No encontré la fila 'Totales Generales:' en el export.");
  }

  return { cuentas, control, categoryTotals };
}

function knownCodes(mapping, cuentasSise) {
  const codes = new Set();
  for (const e of mapping) {
    codes.add(e.code);
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

function findUnmapped(cuentasSise, mapping) {
  const conocidos = knownCodes(mapping, cuentasSise);
  const faltantes = [];
  for (const [code, c] of Object.entries(cuentasSise)) {
    if (conocidos.has(code)) continue;
    if (c.debe === 0 && c.haber === 0) continue;
    faltantes.push({
      code, description: c.descripcion, debe: c.debe, haber: c.haber,
      suggested_category: CATEGORY_BY_FIRST_DIGIT[code[0]] || "RESULTADOS",
    });
  }
  return faltantes;
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

function buildBalance(cuentasSise, mapping, prevBalances) {
  const lineas = [];
  let totalDebe = 0, totalHaber = 0;

  for (const entry of mapping) {
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
            detalle.push({ code: fuente.code, description: fuente.description + " (directo)", debe: dF, haber: hF });
          }
        } else {
          detalle.push({ code: fuente.code, description: fuente.description, debe: dF, haber: hF });
        }
        yaSumados.add(fuente.code);
      }
    } else if (entry.type === "range") {
      for (const [scode, c] of Object.entries(cuentasSise)) {
        if (scode.startsWith(entry.prefix)) { debe += c.debe; haber += c.haber; }
      }
    } else {
      const c = lookupCuenta(cuentasSise, entry);
      debe = c ? c.debe : 0;
      haber = c ? c.haber : 0;
    }

    const movimiento = debe - haber;
    const saldoAnterior = prevBalances[code] || 0;
    const saldoNuevo = saldoAnterior + movimiento;

    lineas.push({
      code, description: entry.description, category: entry.category, type: entry.type,
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
    parseSiseExport, knownCodes, lookupCuenta, findUnmapped, findDuplicateCodes,
    buildBalance, runValidation, CATEGORIES,
  };
}
