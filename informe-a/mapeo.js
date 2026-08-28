// Deriva el mapeo directamente del archivo base, en vez de depender de un
// mapeo_maestro.json preparado de antemano que se puede desincronizar.
//
// Lo importante: la correspondencia entre cada columna de "Dist.de gastos" y su
// centro de costo se saca de LAS FÓRMULAS (a qué columna de saldo apuntan), no del
// título de la columna. Los títulos no son confiables — en el archivo real la
// columna L dice "LOS MORTERITOS" pero suma PROYECTO LOS MORTERITOS.

const RUIDO = ["PROYECTO", "CENTRO", "REGIONAL"];

// Los nombres de centro de costo que NO coinciden letra por letra con el bloque del
// balance. El motor resuelve por texto exacto y no adivina: lo que no coincida tiene que
// estar declarado aca, o no se carga y la app avisa con el nombre textual.
//
// Antes esto se resolvia por parecido, y el parecido mandaba "Tanque Blanco" a TANQUE
// NEGRO y "Cerro Amarillo" a CERRO ABANICO. Si dan de alta un centro de costo nuevo, su
// plata entraba callada en la columna de otro proyecto sin que nada lo mostrara.
//
// La clave va en mayusculas y con los espacios ya normalizados.
const CC_NOMBRES_ONVIO = {
  // como lo escribe Onvio en el export
  "ESPERANZA": "CENTRO ESPERANZA",
  // como lo escriben las filas de la hoja "Gastos Acumulados"
  "PROYECTO SAMENTA": "SAMENTA",
  "PROYECTO LONCO- VACA PELENQUE": "LONCO VACA - PELENQUE",
  "CATAMARCA": "REGIONAL CATAMARCA",
  "LA SANTAS": "LAS SANTAS",          // dice "La Santas", le falta la ese
};

function normTexto(s) {
  if (s === null || s === undefined) return "";
  if (typeof s === "object") {
    if (s.richText) return normTexto(s.richText.map(t => t.text).join(""));
    if (s.result !== undefined) return normTexto(s.result);
    return "";
  }
  return String(s).replace(/\s+/g, " ").trim().toUpperCase();
}

function textoCelda(cell) {
  return normTexto(cell ? cell.value : null);
}

function numeroCelda(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

function formulaCelda(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v && typeof v === "object" && typeof v.formula === "string") return v.formula;
  return null;
}

function indiceACol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// --- centros de costo: tripletas DEBE/HABER/SALDO en la fila 5 ---
function detectarCcBlocks(wsSs) {
  const bloques = [];
  const max = wsSs.columnCount + 3;
  for (let c = 1; c <= max; c++) {
    if (textoCelda(wsSs.getCell(5, c)) !== "DEBE") continue;
    if (textoCelda(wsSs.getCell(5, c + 1)) !== "HABER") continue;
    if (textoCelda(wsSs.getCell(5, c + 2)) !== "SALDO") continue;
    // el nombre está en la fila 4, a veces descentrado por celdas combinadas
    let nombre = null;
    for (const off of [0, 1, -1, 2, -2, 3]) {
      const t = textoCelda(wsSs.getCell(4, c + off));
      if (t) { nombre = t; break; }
    }
    bloques.push({
      nombre_balance: nombre,
      col_debe: indiceACol(c),
      col_haber: indiceACol(c + 1),
      col_saldo: indiceACol(c + 2),
    });
    c += 2;
  }
  return bloques;
}

// --- rango de cuentas y fila TOTALES ---
function ubicarRangoCuentas(wsSs) {
  let primera = null, filaTot = null;
  for (let r = 1; r <= wsSs.rowCount; r++) {
    if (numeroCelda(wsSs.getCell(r, 1)) !== null && primera === null) primera = r;
    if (textoCelda(wsSs.getCell(r, 2)) === "TOTALES") filaTot = r;
  }
  return { primera, filaTot };
}

function detectarCuentas(wsSs) {
  const { primera, filaTot } = ubicarRangoCuentas(wsSs);
  const cuentas = {};
  for (let r = primera; r < filaTot; r++) {
    const codigo = numeroCelda(wsSs.getCell(r, 1));
    const label = wsSs.getCell(r, 2).value;
    if (codigo === null || !label) continue;
    cuentas[String(Math.round(codigo))] = {
      ss_row: r,
      label: String(typeof label === "object" ? normTexto(label) : label).trim(),
      categoria: null,
      ccs_en_dist: [],
    };
  }
  return cuentas;
}

// --- columnas de Dist.de gastos -> centro de costo, según las fórmulas ---
const REF_SS = /'Sumas y Saldos'!\$?([A-Z]{1,3})\$?(\d+)/g;

function detectarDistColumnas(wsDist, ccBlocks, filaFinCategorias) {
  const colSaldoACc = {};
  for (const b of ccBlocks) colSaldoACc[b.col_saldo] = b.nombre_balance;

  const distColToCc = {};
  const sinVotos = [];   // columnas de centro de costo que todavía no tienen ninguna cuenta
  for (let dc = 6; dc <= wsDist.columnCount; dc++) {
    const col = indiceACol(dc);
    const titulo = `${textoCelda(wsDist.getCell(6, dc))} ${textoCelda(wsDist.getCell(7, dc))}`.trim();
    // las columnas de total por mes van después de los centros de costo
    if (titulo.startsWith("TOTAL")) break;

    // ¿a qué columna de saldo apunta la mayoría de las fórmulas de esta columna?
    const votos = {};
    for (let r = 8; r <= filaFinCategorias; r++) {
      const f = formulaCelda(wsDist.getCell(r, dc));
      if (!f) continue;
      let m;
      REF_SS.lastIndex = 0;
      while ((m = REF_SS.exec(f)) !== null) {
        const cc = colSaldoACc[m[1]];
        if (cc) votos[cc] = (votos[cc] || 0) + 1;
      }
    }
    const ganador = Object.entries(votos).sort((a, b) => b[1] - a[1])[0];
    if (ganador) {
      distColToCc[col] = {
        texto_header: titulo,
        nombre_balance: ganador[0],
        refs: ganador[1],
      };
    } else if (titulo) {
      sinVotos.push({ col, titulo });
    }
  }

  // Una columna de un centro de costo que todavía no tiene ninguna cuenta no tiene
  // fórmulas, así que no puede votar. Sin esto quedaba invisible para siempre: no se
  // la reconocía hasta tener una cuenta, y no podía recibir una cuenta hasta que se la
  // reconociera. Para esas —y sólo esas— vale el encabezado.
  const yaTomados = new Set(Object.values(distColToCc).map(i => i.nombre_balance));
  for (const { col, titulo } of sinVotos) {
    const objetivo = normTexto(titulo);
    const bloque = ccBlocks.find(b => !yaTomados.has(b.nombre_balance) && normTexto(b.nombre_balance) === objetivo);
    if (!bloque) continue;
    yaTomados.add(bloque.nombre_balance);
    distColToCc[col] = { texto_header: titulo, nombre_balance: bloque.nombre_balance, refs: 0 };
  }
  return distColToCc;
}

// --- categorías: una fila de Dist.de gastos por categoría ---
function detectarCategorias(wsDist, wsSs, distColToCc, cuentas) {
  const filaTotalGastos = (() => {
    for (let r = 8; r <= wsDist.rowCount; r++) {
      if (textoCelda(wsDist.getCell(r, 3)).includes("TOTAL GASTOS")) return r;
    }
    return wsDist.rowCount;
  })();

  const filaARow = {};
  for (const [codigo, info] of Object.entries(cuentas)) filaARow[info.ss_row] = codigo;

  const categorias = [];
  // hasta la fila de totales SIN incluirla: TOTAL GASTOS no es una categoría
  for (let r = 8; r < filaTotalGastos; r++) {
    const desc = textoCelda(wsDist.getCell(r, 3));
    if (!desc) continue;

    const ssRows = new Set();
    for (const col of Object.keys(distColToCc)) {
      const dc = col.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
      const f = formulaCelda(wsDist.getCell(r, dc));
      if (!f) continue;
      let m;
      REF_SS.lastIndex = 0;
      while ((m = REF_SS.exec(f)) !== null) {
        const fila = parseInt(m[2], 10);
        ssRows.add(fila);
        const codigo = filaARow[fila];
        if (codigo) {
          const cc = distColToCc[col].nombre_balance;
          if (!cuentas[codigo].ccs_en_dist.includes(cc)) cuentas[codigo].ccs_en_dist.push(cc);
          if (!cuentas[codigo].categoria) cuentas[codigo].categoria = desc;
        }
      }
    }
    categorias.push({ desc, dist_row: r, ss_rows: [...ssRows].sort((a, b) => a - b) });
  }
  return { categorias, filaTotalGastos };
}

// --- punto de entrada ---
function derivarMapeo(wb) {
  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const wsSys = wb.getWorksheet("SyS");
  if (!wsSs || !wsDist || !wsSys) {
    throw new Error(
      "El archivo no tiene las hojas esperadas ('Sumas y Saldos', 'Dist.de gastos', 'SyS'). " +
      "¿Es el archivo del Balance de Comprobación?"
    );
  }

  const ccBlocks = detectarCcBlocks(wsSs);
  if (!ccBlocks.length) {
    throw new Error("No encontré los centros de costo en 'Sumas y Saldos' (fila 5: DEBE/HABER/SALDO).");
  }
  const cuentas = detectarCuentas(wsSs);
  const { filaTot } = ubicarRangoCuentas(wsSs);

  // fila hasta donde llegan las categorías (para no leer la fila de totales)
  let filaFin = wsDist.rowCount;
  for (let r = 8; r <= wsDist.rowCount; r++) {
    if (textoCelda(wsDist.getCell(r, 3)).includes("TOTAL GASTOS")) { filaFin = r; break; }
  }

  const distColToCc = detectarDistColumnas(wsDist, ccBlocks, filaFin);
  const { categorias } = detectarCategorias(wsDist, wsSs, distColToCc, cuentas);

  const conColumna = (n) => Object.values(distColToCc).some(i => i.nombre_balance === n);
  let sinColumna = ccBlocks.map(b => b.nombre_balance).filter(n => !conColumna(n));

  // "Sumas y Saldos" repite algunos proyectos con dos nombres —SOL y PROYECTO SOL,
  // LOS MORTERITOS y PROYECTO LOS MORTERITOS— y solo uno de los dos tiene columna en
  // Dist.de gastos. Sin esto el importe podia caer en el bloque sin columna y quedar
  // fuera de la distribucion. Si el nombre huerfano esta contenido en el de un bloque
  // que si tiene columna, los dos apuntan a esa columna.
  const alias = {};
  for (const huerfano of sinColumna.slice()) {
    const gemelo = ccBlocks
      .map(b => b.nombre_balance)
      .filter(n => n !== huerfano && conColumna(n))
      .find(n => n.endsWith(" " + huerfano) || n.startsWith(huerfano + " "));
    if (gemelo) alias[huerfano] = gemelo;
  }
  sinColumna = sinColumna.filter(n => !alias[n]);

  return {
    cc_blocks: ccBlocks,
    dist_col_to_cc: distColToCc,
    cuentas,
    categorias,
    fila_totales: filaTot,
    ccs_sin_columna_dist: sinColumna,
    cc_alias: alias,
    cc_nombres_onvio: Object.assign({}, CC_NOMBRES_ONVIO),
  };
}

if (typeof module !== "undefined") {
  module.exports = { derivarMapeo, indiceACol, normTexto, numeroCelda, formulaCelda, CC_NOMBRES_ONVIO };
}
