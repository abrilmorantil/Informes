// Port de motor_balance.py — carga del Balance de Comprobación USD (SCA).
//
// A diferencia del original, que preguntaba la categoría de las cuentas nuevas por
// consola, acá el proceso se parte en dos etapas para que la pantalla pueda pedirlas
// con un formulario:
//   1. detectarPendientes(...)  -> qué cuentas nuevas hay que clasificar
//   2. procesar(...)            -> hace la carga con esas categorías ya elegidas

const PALABRAS_RUIDO = ["PROYECTO", "CENTRO", "REGIONAL"];

function norm(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/\s+/g, " ").trim().toUpperCase();
}

function limpiar(txt) {
  return norm(txt).split(" ").filter(p => !PALABRAS_RUIDO.includes(p)).join(" ");
}

function colAIndice(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function indiceACol(n) {
  let s = "";
  while (n > 0) {
    const resto = (n - 1) % 26;
    s = String.fromCharCode(65 + resto) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function esFormula(cell) {
  const v = cell.value;
  return !!(v && typeof v === "object" && (v.formula !== undefined || v.sharedFormula !== undefined));
}

function valorNumerico(cell) {
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

function resolverCcBlock(mapeo, nombreOnvio) {
  const nombre = norm(nombreOnvio);
  const nombres = mapeo.cc_blocks.map(b => b.nombre_balance);
  if (nombres.includes(nombre)) {
    return mapeo.cc_blocks.find(b => b.nombre_balance === nombre);
  }
  const objetivo = limpiar(nombre);
  const limpios = {};
  for (const n of nombres) limpios[limpiar(n)] = n;
  if (limpios[objetivo] !== undefined) {
    return mapeo.cc_blocks.find(b => b.nombre_balance === limpios[objetivo]);
  }
  const match = getCloseMatches(objetivo, Object.keys(limpios), 1, 0.6);
  if (match.length) {
    return mapeo.cc_blocks.find(b => b.nombre_balance === limpios[match[0]]);
  }
  return null;
}

function ccNombreADistCol(mapeo, nombreBalance) {
  for (const [col, info] of Object.entries(mapeo.dist_col_to_cc)) {
    if (info && info.nombre_balance === nombreBalance) return col;
  }
  return null;
}

function filaDistDeCategoria(mapeo, categoria) {
  const c = mapeo.categorias.find(x => x.desc === categoria);
  return c ? c.dist_row : null;
}

// Ubica el rango de cuentas: desde la primera fila con código numérico en A hasta
// la fila anterior a TOTALES.
function ubicarRango(wsSs) {
  let primera = null, filaTot = null;
  for (let r = 1; r <= wsSs.rowCount; r++) {
    const a = valorNumerico(wsSs.getCell(r, 1));
    const b = wsSs.getCell(r, 2).value;
    if (a !== null && primera === null) primera = r;
    if (typeof b === "string" && b.trim().toUpperCase() === "TOTALES") filaTot = r;
  }
  return { primera, filaTot };
}

function agregarRefDistDeGastos(wsDist, distRow, distCol, colSaldoSs, ssRow) {
  const celda = wsDist.getCell(`${distCol}${distRow}`);
  const ref = `'Sumas y Saldos'!${colSaldoSs}${ssRow}`;
  const v = celda.value;

  if (v && typeof v === "object" && v.sharedFormula) {
    throw new Error(
      `La celda Dist.de gastos!${distCol}${distRow} usa una fórmula compartida y no ` +
      `se le puede agregar la referencia de la cuenta nueva de forma segura. ` +
      `NO se generó ningún archivo.`
    );
  }

  // ExcelJS guarda el texto de la fórmula SIN el '=' inicial.
  if (v && typeof v === "object" && typeof v.formula === "string") {
    celda.value = { formula: `${v.formula}+${ref}` };
  } else {
    celda.value = { formula: `+${ref}` };
  }
}

// Cada corrida carga SOLO el mes en curso, así que antes de escribir hay que borrar
// los saldos del mes anterior. Si no, las cuentas sin movimiento este mes conservan
// su valor viejo y el total no cierra.
function limpiarSaldosAnteriores(wsSs, mapeo, log) {
  const { primera, filaTot } = ubicarRango(wsSs);
  if (primera === null || filaTot === null) {
    log("  ⚠ No pude ubicar el rango de datos; no limpié saldos anteriores.");
    return 0;
  }
  const fin = filaTot - 1;
  let borradas = 0;
  for (const b of mapeo.cc_blocks) {
    for (const col of [b.col_debe, b.col_haber, b.col_saldo]) {
      if (!col) continue;
      const ci = colAIndice(col);
      for (let r = primera; r <= fin; r++) {
        const cell = wsSs.getCell(r, ci);
        if (esFormula(cell)) continue;
        if (cell.value !== null && cell.value !== undefined && cell.value !== 0) borradas++;
        cell.value = 0;
      }
    }
  }
  log(`  Limpiados saldos anteriores: ${borradas} celdas puestas en 0 (rango ${primera}:${fin}).`);
  return borradas;
}

// La fila TOTALES del archivo original viene con valores fijos pegados a mano, no
// fórmulas, así que nunca recalcula. Se reescribe con =SUM() reales.
function reconstruirFilaTotales(wsSs, mapeo, log) {
  const { primera, filaTot } = ubicarRango(wsSs);
  if (primera === null || filaTot === null) {
    log("  ⚠ No pude ubicar la fila TOTALES o el rango de datos; no la reconstruí.");
    return 0;
  }
  const finRango = filaTot - 1;
  let puestas = 0;
  for (const b of mapeo.cc_blocks) {
    for (const col of [b.col_debe, b.col_haber, b.col_saldo]) {
      if (!col) continue;
      wsSs.getCell(`${col}${filaTot}`).value = { formula: `SUM(${col}${primera}:${col}${finRango})` };
      puestas++;
    }
  }
  log(`  Fila TOTALES (${filaTot}) reconstruida con ${puestas} fórmulas SUM (rango ${primera}:${finRango}).`);
  return puestas;
}

// Idem en vertical: la columna "TOTALES SALDOS" (BQ) suma, por cada cuenta, sus
// saldos en todos los centros de costo, y también venía con valores fijos.
function reconstruirColumnaTotalesSaldos(wsSs, mapeo, log) {
  let colTot = null;
  for (let c = 1; c <= wsSs.columnCount + 1; c++) {
    const v4 = wsSs.getCell(4, c).value;
    const v5 = wsSs.getCell(5, c).value;
    if (v4 && v5 && String(v4).toUpperCase().includes("TOTAL") && String(v5).toUpperCase().includes("SALDO")) {
      colTot = indiceACol(c);
      break;
    }
  }
  if (colTot === null) {
    const ultimoCc = Math.max(...mapeo.cc_blocks.map(b => colAIndice(b.col_saldo)));
    colTot = indiceACol(ultimoCc + 1);
    log(`  (columna TOTALES SALDOS no rotulada; usando ${colTot} por posición)`);
  }

  const { primera, filaTot } = ubicarRango(wsSs);
  if (primera === null || filaTot === null) {
    log("  ⚠ No pude ubicar el rango; no reconstruí la columna TOTALES SALDOS.");
    return 0;
  }

  const colsSaldo = mapeo.cc_blocks.map(b => b.col_saldo).filter(Boolean);
  let puestas = 0;
  for (let r = primera; r < filaTot; r++) {
    if (valorNumerico(wsSs.getCell(r, 1)) !== null) {
      wsSs.getCell(`${colTot}${r}`).value = { formula: colsSaldo.map(c => `${c}${r}`).join("+") };
      puestas++;
    }
  }
  wsSs.getCell(`${colTot}${filaTot}`).value = { formula: `SUM(${colTot}${primera}:${colTot}${filaTot - 1})` };
  log(`  Columna TOTALES SALDOS (${colTot}) reconstruida con ${puestas} fórmulas.`);
  return puestas;
}

// Etapa 1: qué cuentas del export no se pueden cargar sin que alguien elija su
// categoría, y qué centros de costo no se pudieron identificar.
function detectarPendientes(lineas, mapeo) {
  const pendientes = [];
  const sinCc = [];
  const vistas = new Set();

  for (const linea of lineas) {
    const ccBlock = resolverCcBlock(mapeo, linea.cc_nombre_onvio);
    if (ccBlock === null) {
      sinCc.push(linea.cc_nombre_onvio);
      continue;
    }
    const codigo = linea.cuenta_codigo;
    if (vistas.has(codigo)) continue;
    const cuenta = mapeo.cuentas[codigo];
    if (cuenta === undefined || cuenta.categoria === null || cuenta.categoria === undefined) {
      vistas.add(codigo);
      pendientes.push({
        codigo,
        label: linea.cuenta_label,
        cc_nombre: ccBlock.nombre_balance,
        saldo: linea.saldo,
        tipo: cuenta === undefined ? "nueva" : "sin_categoria",
      });
    }
  }

  return { pendientes, sinCc: [...new Set(sinCc)].sort() };
}

function categoriasDisponibles(mapeo) {
  return [...new Set(mapeo.categorias.map(c => c.desc))].sort();
}

// Etapa 2. `categoriasElegidas` es un objeto {codigo_de_cuenta: nombre_de_categoria}
// con lo que se haya respondido en el formulario para las cuentas pendientes.
// Devuelve el mapeo actualizado; el que se recibe no se modifica.
function procesar({ wb, lineas, mapeo, categoriasElegidas = {}, periodo = null, log = () => {} }) {
  mapeo = JSON.parse(JSON.stringify(mapeo));

  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const wsSys = wb.getWorksheet("SyS");
  if (!wsSs || !wsDist || !wsSys) {
    throw new Error(
      "El archivo maestro no tiene las hojas esperadas ('Sumas y Saldos', " +
      "'Dist.de gastos', 'SyS'). ¿Es el .xlsm del Balance de Comprobación?"
    );
  }

  // Excel tiene que recalcular al abrir: las fórmulas que se reescriben acá quedan
  // sin resultado cacheado.
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  log(`${lineas.length} líneas de cuenta leídas del balance.`);

  let nuevas = 0, clasificadas = 0, conocidas = 0;
  const sinCc = [];

  limpiarSaldosAnteriores(wsSs, mapeo, log);

  for (const linea of lineas) {
    const ccBlock = resolverCcBlock(mapeo, linea.cc_nombre_onvio);
    if (ccBlock === null) {
      sinCc.push(linea.cc_nombre_onvio);
      continue;
    }

    const codigo = linea.cuenta_codigo;
    let cuenta = mapeo.cuentas[codigo];

    if (cuenta === undefined) {
      const categoria = categoriasElegidas[codigo];
      if (!categoria) {
        throw new Error(`Falta elegir la categoría de la cuenta nueva ${codigo} - ${linea.cuenta_label}.`);
      }
      const cat = mapeo.categorias.find(c => c.desc === categoria);
      if (!cat) {
        throw new Error(
          `"${categoria}" no es una categoría existente. Crear categorías nuevas en ` +
          `Dist.de gastos no está soportado todavía en este motor; elegí una de la lista.`
        );
      }

      const insertAt = cat.ss_rows.length ? Math.max(...cat.ss_rows) + 1 : 399;

      const mod = insertRowSumasYSaldos(wsSs, wb, insertAt);
      log(`  Insertada fila ${insertAt} en Sumas y Saldos (${mod} referencias de fórmula reacomodadas).`);

      // corregir en memoria las filas del mapeo que quedaron corridas
      for (const info of Object.values(mapeo.cuentas)) {
        if (info.ss_row >= insertAt) info.ss_row += 1;
      }
      for (const c2 of mapeo.categorias) {
        c2.ss_rows = c2.ss_rows.map(r => (r >= insertAt ? r + 1 : r));
      }

      wsSs.getCell(insertAt, 1).value = parseInt(codigo, 10);
      wsSs.getCell(insertAt, 2).value = linea.cuenta_label;

      cuenta = { ss_row: insertAt, label: linea.cuenta_label, categoria, ccs_en_dist: [] };
      mapeo.cuentas[codigo] = cuenta;
      cat.ss_rows.push(insertAt);
      nuevas++;
    } else if (cuenta.categoria === null || cuenta.categoria === undefined) {
      const categoria = categoriasElegidas[codigo];
      if (!categoria) {
        throw new Error(`Falta elegir la categoría de la cuenta ${codigo} - ${linea.cuenta_label}.`);
      }
      const cat = mapeo.categorias.find(c => c.desc === categoria);
      if (!cat) throw new Error(`"${categoria}" no es una categoría existente.`);
      cuenta.categoria = categoria;
      if (!cuenta.ccs_en_dist) cuenta.ccs_en_dist = [];
      if (!cat.ss_rows.includes(cuenta.ss_row)) cat.ss_rows.push(cuenta.ss_row);
      clasificadas++;
    } else {
      conocidas++;
    }

    // Traba de seguridad: la fila que el mapeo dice para esta cuenta, ¿realmente
    // tiene ese código en el archivo que se está procesando? Si no coincide, el
    // archivo y el mapeo están desincronizados y se frena antes de pisar datos.
    const rCheck = cuenta.ss_row;
    const codigoEnArchivo = valorNumerico(wsSs.getCell(rCheck, 1));
    if (codigoEnArchivo === null || parseInt(codigoEnArchivo, 10) !== parseInt(codigo, 10)) {
      throw new Error(
        `DESINCRONIZACIÓN DETECTADA: el mapeo dice que la cuenta ${codigo} está en la ` +
        `fila ${rCheck} de "Sumas y Saldos", pero esa fila tiene la cuenta ` +
        `${codigoEnArchivo} ("${wsSs.getCell(rCheck, 2).value}").\n` +
        `Esto pasa si procesaste el balance ORIGINAL después de que ya se había ` +
        `insertado una fila nueva en una corrida anterior.\n` +
        `Solución: usá como entrada el archivo que generó la corrida anterior (no el ` +
        `original), o volvé a empezar con una copia limpia del mapeo.\n` +
        `NO SE GUARDÓ NADA — tu archivo de entrada está intacto.`
      );
    }

    // asegurar que esta cuenta esté referenciada en Dist.de gastos para este CC
    if (!cuenta.ccs_en_dist) cuenta.ccs_en_dist = [];
    if (!cuenta.ccs_en_dist.includes(ccBlock.nombre_balance)) {
      const distRow = filaDistDeCategoria(mapeo, cuenta.categoria);
      const distCol = ccNombreADistCol(mapeo, ccBlock.nombre_balance);
      if (distRow && distCol) {
        agregarRefDistDeGastos(wsDist, distRow, distCol, ccBlock.col_saldo, cuenta.ss_row);
        cuenta.ccs_en_dist.push(ccBlock.nombre_balance);
      }
    }

    const r = cuenta.ss_row;
    wsSs.getCell(`${ccBlock.col_debe}${r}`).value = linea.debe;
    wsSs.getCell(`${ccBlock.col_haber}${r}`).value = linea.haber;
    wsSs.getCell(`${ccBlock.col_saldo}${r}`).value = linea.saldo;

    // espejo en SyS (igual que hacía la macro), buscando/creando por clave
    const clave = `${linea.cc_codigo}_${codigo}`;
    let filaSys = null;
    for (let rr = 1; rr <= wsSys.rowCount; rr++) {
      if (wsSys.getCell(rr, 2).value === clave) { filaSys = rr; break; }
    }
    if (filaSys === null) {
      filaSys = wsSys.rowCount + 1;
      wsSys.getCell(filaSys, 1).value = `${codigo} - ${linea.cuenta_label}`;
      wsSys.getCell(filaSys, 2).value = clave;
    }
    wsSys.getCell(filaSys, 11).value = linea.debe;
    wsSys.getCell(filaSys, 13).value = linea.haber;
    wsSys.getCell(filaSys, 16).value = linea.saldo;
  }

  if (sinCc.length) {
    log(`\n⚠ ${sinCc.length} centro(s) de costo del balance no se pudo(pudieron) identificar ` +
        `en Sumas y Saldos: ${[...new Set(sinCc)].sort().join(", ")}`);
  }

  reconstruirFilaTotales(wsSs, mapeo, log);
  reconstruirColumnaTotalesSaldos(wsSs, mapeo, log);

  if (periodo) {
    wsSs.getCell("A2").value = periodo;
    log(`  Rótulo de período escrito en A2: ${periodo}`);
  }

  log(`\nResumen: ${conocidas} cuentas ya conocidas, ${clasificadas} recién clasificadas, ` +
      `${nuevas} cuentas nuevas insertadas.`);

  return { mapeo, resumen: { conocidas, clasificadas, nuevas, sinCc: [...new Set(sinCc)].sort() } };
}

if (typeof module !== "undefined") {
  const { getCloseMatches } = require("./similitud.js");
  const { insertRowSumasYSaldos } = require("./formula_utils.js");
  global.getCloseMatches = getCloseMatches;
  global.insertRowSumasYSaldos = insertRowSumasYSaldos;
  module.exports = {
    procesar, detectarPendientes, categoriasDisponibles, resolverCcBlock,
    limpiarSaldosAnteriores, reconstruirFilaTotales, reconstruirColumnaTotalesSaldos,
    norm, limpiar, colAIndice, indiceACol,
  };
}
