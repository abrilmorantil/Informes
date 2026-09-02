// "De dónde sale cada saldo" — la hoja de trazabilidad del balance.
//
// El equivalente de la que ya tiene el BALCOMPROBDOLARES, pero acá la cadena es más larga.
// Un importe recorre cuatro escalones antes de llegar al estado impreso:
//
//     Onvio  ->  Hoja1  ->  SALDOS  ->  Anexo II  o  Nota 4  ->  Balance
//
// Cada escalón se puede romper por su cuenta y ninguno avisa:
//   - Hoja1 -> SALDOS es un VLOOKUP contra el TEXTO "código - nombre". Si el texto no
//     coincide, la cuenta queda en cero y el balance igual cierra.
//   - SALDOS -> Anexo II lee el subtotal de la cuenta madre. Una madre que entra dos veces
//     cuenta el gasto doble; una que no entra lo hace desaparecer.
//   - SALDOS -> Nota 4 a veces es una referencia suelta y a veces un rango (los proveedores
//     entran por `SUM(G63:G117)`). Una cuenta insertada fuera de ese rango se carga y no
//     aparece en ningún estado: fue lo que pasó en agosto 2026 con dos proveedores nuevos y
//     1.915.260,00 que no llegaban al pasivo.
//
// Por eso la hoja lista una fila por cuenta con los cuatro escalones al lado, y marca en qué
// escalón se corta. No es un listado: es el control de que ninguno se corta.
//
// Nombres con prefijo `tz` por el ámbito global único que comparten los scripts del sitio.

const TZ_RE_CUENTA = /^\s*(\d{5,})\s*-\s*(.+?)\s*$/;

function tzTexto(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return typeof v.result === "string" ? v.result : "";
    return v.result === undefined || v.result === null ? "" : String(v.result);
  }
  return String(v);
}
function tzNumero(cell) {
  const v = cell && cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}
function tzLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
// Excel compara las claves del VLOOKUP sin distinguir mayúsculas. Comparar exacto daba
// falsos "no la encuentra" con cosas como "Materiales de campo" / "Materiales de Campo".
const tzClave = (s) => String(s).trim().toLowerCase().replace(/\s+/g, " ");

// Todas las filas de SALDOS que una fórmula nombra, contando los rangos.
// El rango importa: los proveedores entran al pasivo por `SUM(G63:G117)`, no uno por uno.
function tzFilasQueNombra(formula) {
  const out = new Set();
  const t = String(formula || "");
  const reRango = /(?:'?SALDOS'?!)?\$?([A-Z]{1,3})\$?(\d+)\s*:\s*(?:'?SALDOS'?!)?\$?([A-Z]{1,3})\$?(\d+)/gi;
  let m;
  const yaEnRango = [];
  while ((m = reRango.exec(t))) {
    const a = Math.min(+m[2], +m[4]), b = Math.max(+m[2], +m[4]);
    for (let r = a; r <= b; r++) out.add(r);
    yaEnRango.push([m.index, m.index + m[0].length]);
  }
  const reSuelta = /(?:'?SALDOS'?!)?\$?[A-Z]{1,3}\$?(\d+)/gi;
  while ((m = reSuelta.exec(t))) {
    if (yaEnRango.some(([a, b]) => m.index >= a && m.index < b)) continue;
    out.add(+m[1]);
  }
  return out;
}

// Dónde termina cada fila de SALDOS: en qué celda de qué hoja de estado se la lee.
// Se recorren todas las hojas menos SALDOS y Hoja1, que son el camino, no el destino.
function tzDestinos(wb) {
  const porFila = new Map();

  // Primero, los subtotales que viven DENTRO de SALDOS. Los proveedores no llegan a la Nota 4
  // uno por uno: entran a `SALDOS!I117 = SUM(G63:G117)` y la Nota 4 lee esa celda. Sin este
  // salto, las 55 filas del bloque figuraban como "no entra a ningun estado", que es
  // justamente lo contrario de lo que pasa.
  const subtotales = new Map();          // "I117" -> [filas que abarca]
  const sa = wb.getWorksheet("SALDOS");
  if (sa) {
    sa.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, ci) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && v.formula)) return;
        const f = String(v.formula);
        if (/VLOOKUP/i.test(f) || !/[:+]/.test(f)) return;
        const filas = [...tzFilasQueNombra(f)].filter(x => x !== r);
        if (filas.length < 2) return;               // no agrega a nadie: no es un subtotal
        subtotales.set(tzLetra(ci) + r, { fila: r, col: ci, filas, porRango: /:/.test(f) });
      });
    });
  }

  for (const ws of wb.worksheets) {
    if (ws.name === "SALDOS" || ws.name === "Hoja1") continue;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, ci) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && v.formula)) return;
        const f = String(v.formula);
        if (!/SALDOS!/i.test(f)) return;
        // el rótulo de la fila: la última celda con texto a la izquierda del importe
        let etiqueta = "";
        for (let k = 1; k < ci; k++) {
          const t = tzTexto(ws.getCell(r, k).value).trim();
          if (t) etiqueta = t;
        }
        for (const fila of tzFilasQueNombra(f)) {
          const anotar = (destFila, viaSubtotal) => {
            if (!porFila.has(destFila)) porFila.set(destFila, []);
            porFila.get(destFila).push({
              hoja: ws.name, celda: tzLetra(ci) + r, fila: r, etiqueta,
              porRango: /:/.test(f), viaSubtotal: viaSubtotal || null,
            });
          };
          anotar(fila, null);
          // si esa fila de SALDOS es un subtotal, el destino vale para todas las que suma
          const sub = subtotales.get(tzLetra(7) + fila) ||
                      [...subtotales.values()].find(x => x.fila === fila);
          if (sub) for (const hija of sub.filas) anotar(hija, tzLetra(sub.col) + sub.fila);
        }
      });
    });
  }
  return porFila;
}

// Una fila por cuenta de SALDOS, con los cuatro escalones.
//
// `madres` es lo que devuelve `madresResultados` (pesos) o la lista de líneas de RESULTADOS
// (dólares): sirve para decir, de una subcuenta, cuál es la madre que la lleva al Anexo II.
function tzFilas(wb, moneda, mapeo, madres) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  const h1 = wb.getWorksheet("Hoja1");
  if (!ws) return [];

  // Hoja1, por clave normalizada
  const enHoja1 = new Map();
  if (h1) {
    h1.eachRow({ includeEmpty: false }, (row, r) => {
      if (r < 2) return;
      const t = tzTexto(row.getCell(p.hoja1ColClave).value).trim();
      if (!t) return;
      const k = tzClave(t);
      if (!enHoja1.has(k)) enHoja1.set(k, { fila: r, texto: t, valor: tzNumero(row.getCell(p.hoja1ColValor)) });
    });
  }

  const destinos = tzDestinos(wb);

  // de qué madre es hija cada fila (para las subcuentas de RESULTADOS)
  const madreDe = new Map();
  for (const m of (madres || [])) {
    const b = m.bloque;
    if (!b) continue;
    const filas = b.tipo === "rango"
      ? Array.from({ length: b.hasta - b.desde + 1 }, (_, i) => b.desde + i)
      : (b.filas || []);
    for (const f of filas) if (f !== m.fila) madreDe.set(f, m);
  }

  const filas = [];
  for (const [codigo, info] of Object.entries(mapeo.cuentas || {})) {
    // El VLOOKUP no siempre esta en la misma columna: en el ACTIVO va en la del saldo (G) y
    // en RESULTADOS en la de al lado (F), con el subtotal de la madre en G. Mirar una sola
    // daba 390 filas "manuales" cuando los manuales son solo los ~55 proveedores.
    let g = ws.getCell(info.fila, p.saldosColValor);
    let formula = g && g.formula ? String(g.formula) : "";
    if (!/VLOOKUP/i.test(formula)) {
      const alt = ws.getCell(info.fila, p.saldosColValor - 1);
      if (alt && alt.formula && /VLOOKUP/i.test(String(alt.formula))) {
        g = alt; formula = String(alt.formula);
      }
    }

    // el escalón Hoja1 -> SALDOS: cuál es la clave que busca y si la encuentra
    let claveBuscada = null, enH1 = null;
    const mv = /VLOOKUP\(\s*\$?([A-Z]{1,3})\$?(\d+)/i.exec(formula);
    if (mv) {
      const col = mv[1].toUpperCase().split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
      claveBuscada = tzTexto(ws.getCell(+mv[2], col).value).trim();
      enH1 = enHoja1.get(tzClave(claveBuscada)) || null;
    } else {
      // filas manuales (el detalle de proveedores): el importe se escribe, no se busca
      enH1 = enHoja1.get(tzClave(info.clave)) || null;
    }

    // el escalón SALDOS -> estado
    const madre = madreDe.get(info.fila) || null;
    const propios = destinos.get(info.fila) || [];
    const deMadre = madre ? (destinos.get(madre.fila) || []) : [];
    const via = propios.length ? propios : deMadre;

    filas.push({
      codigo,
      nombre: info.nombre || "",
      // Hoja1
      hoja1Fila: enH1 ? enH1.fila : null,
      hoja1Importe: enH1 ? enH1.valor : null,
      claveBuscada,
      encuentraEnHoja1: !mv ? null : !!enH1,   // null = no busca, se escribe a mano
      // SALDOS
      saldosFila: info.fila,
      saldosImporte: tzNumero(g),
      esManual: !mv,
      // madre, si es subcuenta
      madreCodigo: madre ? madre.codigo : null,
      madreNombre: madre ? madre.nombre : null,
      madreFila: madre ? madre.fila : null,
      // destino
      destinos: via.map(d => ({ ...d, atravesMadre: !propios.length && !!madre })),
      sinDestino: via.length === 0,
    });
  }
  filas.sort((a, b) => a.saldosFila - b.saldosFila);
  return filas;
}

// Escribe la hoja en el libro. Se llama justo antes de descargar: NO se guarda en el maestro,
// porque es una foto de esta corrida y el mes que viene sería mentira.
function tzEscribirHoja(wb, moneda, mapeo, madres) {
  const NOMBRE = "De dónde sale cada saldo";
  // Si no se los pasan, se deducen del propio libro: así el que descarga no tiene que
  // acordarse de armarlos, y no pueden quedar desfasados del archivo que se está escribiendo.
  if (!mapeo) mapeo = derivarMapeoMaestro(wb, moneda, typeof clasificacion !== "undefined" ? clasificacion : null);
  if (!madres) {
    madres = madresResultados(wb, moneda);
    if (!madres.length && moneda === "dolares" && typeof pcdLineasResultados === "function") {
      madres = pcdLineasResultados(wb, mapeo);
    }
  }
  const vieja = wb.getWorksheet(NOMBRE);
  if (vieja) wb.removeWorksheet(vieja.id);
  const filas = tzFilas(wb, moneda, mapeo, madres);
  const ws = wb.addWorksheet(NOMBRE);
  const bold = { bold: true };
  const numFmt = "#,##0.00";

  ws.getCell("A1").value = "De dónde sale cada saldo del balance";
  ws.getCell("A1").font = { bold: true, size: 13 };
  ws.getCell("A2").value =
    "Una fila por cuenta. Se sigue el camino completo: lo que manda Onvio entra en Hoja1, " +
    "SALDOS lo busca por el texto \"código - nombre\", y de ahí va al Anexo II (los gastos) " +
    "o a la Nota 4 (activo y pasivo). La columna \"¿Dónde se corta?\" dice en qué escalón se " +
    "pierde una cuenta, si se pierde.";
  ws.getCell("A2").font = { italic: true, size: 9 };

  const headers = [
    "Código", "Nombre", "Fila Hoja1", "Importe en Hoja1", "Fila SALDOS", "Importe en SALDOS",
    "Cuenta madre", "Va a", "Renglón", "Celda", "¿Dónde se corta?",
  ];
  headers.forEach((h, j) => {
    const c = ws.getCell(4, 1 + j);
    c.value = h; c.font = bold; c.border = { bottom: { style: "thin" } };
  });

  let r = 5;
  for (const f of filas) {
    const d = f.destinos[0] || null;
    let corte = "";
    if (f.encuentraEnHoja1 === false) corte = "Hoja1 no tiene esa cuenta con ese texto";
    else if (f.sinDestino) corte = "no entra a ningún estado";
    ws.getCell(r, 1).value = f.codigo;
    ws.getCell(r, 2).value = f.nombre;
    const cf1 = ws.getCell(r, 3); cf1.value = f.hoja1Fila; cf1.numFmt = "0";
    const ch = ws.getCell(r, 4); ch.value = f.hoja1Importe; ch.numFmt = numFmt;
    const cf2 = ws.getCell(r, 5); cf2.value = f.saldosFila; cf2.numFmt = "0";
    const cs = ws.getCell(r, 6); cs.value = f.saldosImporte; cs.numFmt = numFmt;
    ws.getCell(r, 7).value = f.madreCodigo ? `${f.madreCodigo} - ${f.madreNombre}` : "";
    ws.getCell(r, 8).value = d ? d.hoja : "";
    ws.getCell(r, 9).value = d ? d.etiqueta : "";
    ws.getCell(r, 10).value = d ? d.celda + (d.porRango ? " (por rango)" : "") : "";
    ws.getCell(r, 11).value = corte;
    if (corte) {
      for (let c = 1; c <= headers.length; c++) {
        ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE0E0" } };
      }
    }
    r++;
  }

  ws.getCell(r + 1, 1).value = "TOTAL";
  ws.getCell(r + 1, 1).font = bold;
  for (const col of [4, 6]) {
    const c = ws.getCell(r + 1, col);
    c.value = { formula: `SUM(${tzLetra(col)}5:${tzLetra(col)}${r - 1})` };
    c.font = bold; c.numFmt = numFmt;
  }
  ws.getCell(r + 2, 1).value =
    "Los dos totales de importe tienen que dar lo mismo. Ojo: \"Importe en SALDOS\" son " +
    "fórmulas, así que en un archivo recién descargado sale en cero hasta que lo abrís en " +
    "Excel; recién ahí se puede comparar.";
  ws.getCell(r + 2, 1).font = { italic: true, size: 9 };

  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: r - 1, column: headers.length } };
  ws.views = [{ state: "frozen", ySplit: 4 }];
  ws.columns = [{ width: 13 }, { width: 44 }, { width: 11 }, { width: 18 }, { width: 12 },
                { width: 18 }, { width: 34 }, { width: 18 }, { width: 34 }, { width: 16 },
                { width: 36 }];
  return { filas: filas.length, cortes: filas.filter(x => x.encuentraEnHoja1 === false || x.sinDestino).length };
}

if (typeof module !== "undefined") {
  module.exports = { tzFilas, tzEscribirHoja, tzDestinos, tzFilasQueNombra, tzTexto };
}
