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

// Las CELDAS de SALDOS que una fórmula nombra: columna y fila.
//
// La columna hace falta, no es un detalle. La fila 61 tiene la cuenta "Cargos Diferidos" en
// G61 y, en la MISMA fila, el total del activo en `I61 = SUM(G3:G61)`. Emparejando sólo por
// fila, la Nota 4 que lee G61 parecía leer el total: y entonces las 59 cuentas del activo
// figuraban entrando dos veces al balance. Ninguna lo hacía.
function tzCeldasQueNombra(formula) {
  const out = [];
  const t = String(formula || "");
  const idx = (L) => L.toUpperCase().split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const reRango = /(?:'?SALDOS'?!)?\$?([A-Z]{1,3})\$?(\d+)\s*:\s*(?:'?SALDOS'?!)?\$?([A-Z]{1,3})\$?(\d+)/gi;
  let m;
  const yaEnRango = [];
  while ((m = reRango.exec(t))) {
    const c1 = idx(m[1]), c2 = idx(m[3]);
    const f1 = Math.min(+m[2], +m[4]), f2 = Math.max(+m[2], +m[4]);
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
      for (let r = f1; r <= f2; r++) out.push({ col: c, fila: r, porRango: true });
    }
    yaEnRango.push([m.index, m.index + m[0].length]);
  }
  const reSuelta = /(?:'?SALDOS'?!)?\$?([A-Z]{1,3})\$?(\d+)/gi;
  while ((m = reSuelta.exec(t))) {
    if (yaEnRango.some(([a, b]) => m.index >= a && m.index < b)) continue;
    out.push({ col: idx(m[1]), fila: +m[2], porRango: false });
  }
  return out;
}

// Compatibilidad: sólo las filas, sin la columna.
function tzFilasQueNombra(formula) {
  return new Set(tzCeldasQueNombra(formula).map(x => x.fila));
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
        const filas = [...new Set(tzCeldasQueNombra(f).map(x => x.fila))].filter(x => x !== r);
        if (filas.length < 2) return;               // no agrega a nadie: no es un subtotal
        subtotales.set(ci + "," + r, { fila: r, col: ci, filas, porRango: /:/.test(f) });
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
        const anotar = (destFila, viaSubtotal) => {
          if (!porFila.has(destFila)) porFila.set(destFila, []);
          const ya = porFila.get(destFila);
          const clave = ws.name + "!" + tzLetra(ci) + r;
          if (ya.some(d => d.hoja + "!" + d.celda === clave)) return;   // ya anotado
          ya.push({
            hoja: ws.name, celda: tzLetra(ci) + r, fila: r, col: ci, etiqueta,
            porRango: /:/.test(f), viaSubtotal: viaSubtotal || null,
          });
        };
        for (const cel of tzCeldasQueNombra(f)) {
          anotar(cel.fila, null);
          // Si esa CELDA es un subtotal, el destino vale para todas las filas que suma. Se
          // empareja por celda —columna y fila—, no por fila: en una misma fila conviven la
          // cuenta y un total que no tiene nada que ver con ella.
          const sub = subtotales.get(cel.col + "," + cel.fila);
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

  // Las cuatro columnas de gasto del Anexo II, para poder nombrar el TIPO DE GASTO en vez de
  // decir "columna H", que no le dice nada a nadie.
  //
  // Se leen de la propia hoja y no se depende de `a2Columnas`: si ese módulo no estuviera
  // cargado, la columna salía vacía y en silencio —pasó en el test— y una columna vacía no se
  // nota hasta que alguien la busca.
  const colsA2 = new Map();
  const a2ws = wb.getWorksheet("Anexo II");
  if (a2ws) {
    for (let c = 1; c <= a2ws.columnCount; c++) {
      const t = tzTexto(a2ws.getCell(8, c).value).trim();
      if (/ADMINISTRACION|COMERCIALIZACION|EXPLORACION|FINANCIEROS/i.test(t)) colsA2.set(c, t);
    }
  }

  ws.getCell("A1").value = "De dónde sale cada saldo del balance";
  ws.getCell("A1").font = { bold: true, size: 13 };
  ws.getCell("A2").value =
    "Una fila por cuenta: dónde sale impresa y, si es un gasto, en qué concepto del Anexo II " +
    "queda guardada y de qué tipo es. Las subcuentas llegan al Anexo II a través de su cuenta " +
    "madre. Filtrá por concepto para ver todo lo que arma cualquier renglón.";
  ws.getCell("A2").font = { italic: true, size: 9 };

  const headers = ["Cuenta", "Cuenta madre", "Sale impresa en", "Tipo de gasto",
                   "Importe", "Para revisar"];
  headers.forEach((h, j) => {
    const c = ws.getCell(4, 1 + j);
    c.value = h; c.font = bold; c.border = { bottom: { style: "thin" } };
  });

  let r = 5;
  for (const f of filas) {
    const d = f.destinos[0] || null;
    let revisar = "";
    if (f.encuentraEnHoja1 === false) revisar = "Onvio no manda esta cuenta con ese nombre";
    else if (f.sinDestino) revisar = "no sale impresa en ningún lado";
    ws.getCell(r, 1).value = `${f.codigo} - ${f.nombre}`;
    ws.getCell(r, 2).value = f.madreCodigo ? `${f.madreCodigo} - ${f.madreNombre}` : "";
    ws.getCell(r, 3).value = d ? d.etiqueta : "";
    ws.getCell(r, 4).value = (d && d.hoja === "Anexo II") ? (colsA2.get(d.col) || "") : "";
    const ci = ws.getCell(r, 5); ci.value = f.hoja1Importe; ci.numFmt = numFmt;
    ws.getCell(r, 6).value = revisar;
    if (revisar) {
      for (let c = 1; c <= headers.length; c++) {
        ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE0E0" } };
      }
    }
    r++;
  }

  const tot = ws.getCell(r + 1, 5);
  ws.getCell(r + 1, 1).value = "TOTAL";
  ws.getCell(r + 1, 1).font = bold;
  tot.value = { formula: `SUM(E5:E${r - 1})` };
  tot.font = bold; tot.numFmt = numFmt;
  ws.getCell(r + 2, 1).value =
    "El importe es el que manda Onvio. La suma de todos tiene que dar cero, como cualquier " +
    "balance.";
  ws.getCell(r + 2, 1).font = { italic: true, size: 9 };

  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: r - 1, column: headers.length } };
  ws.views = [{ state: "frozen", ySplit: 4 }];
  ws.columns = [{ width: 46 }, { width: 40 }, { width: 38 }, { width: 20 }, { width: 18 },
                { width: 40 }];
  return { filas: filas.length, cortes: filas.filter(x => x.encuentraEnHoja1 === false || x.sinDestino).length };
}

if (typeof module !== "undefined") {
  module.exports = { tzFilas, tzEscribirHoja, tzDestinos, tzFilasQueNombra, tzCeldasQueNombra, tzTexto };
}
