// Motor de carga del Balance de Comprobación USD (SCA).
//
// El archivo maestro calcula todo solo: "Sumas y Saldos" trae los importes de la
// hoja "SyS" con fórmulas VLOOKUP, y de ahí salen "Dist.de gastos", los totales y
// la columna de totales por cuenta. Por eso el motor escribe ÚNICAMENTE en "SyS":
// cualquier número puesto a mano en "Sumas y Saldos" pisaría esas fórmulas y
// rompería el archivo.
//
// Lo único que se toca fuera de "SyS" es cuando aparece una cuenta que todavía no
// existe: ahí hay que crearle la fila (con sus fórmulas) y referenciarla en
// "Dist.de gastos".

const PALABRAS_RUIDO = ["PROYECTO", "CENTRO", "REGIONAL"];

function norm(s) {
  if (s === null || s === undefined) return "";
  if (typeof s === "object") {
    if (s.richText) return norm(s.richText.map(t => t.text).join(""));
    if (s.result !== undefined) return norm(s.result);
    return "";
  }
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
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function valorNumerico(cell) {
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

function textoPlano(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined && typeof v.result !== "object") return String(v.result);
    return null;
  }
  return String(v);
}

function resolverCcBlock(mapeo, nombreOnvio) {
  const nombre = norm(nombreOnvio);
  const nombres = mapeo.cc_blocks.map(b => b.nombre_balance);
  if (nombres.includes(nombre)) return mapeo.cc_blocks.find(b => b.nombre_balance === nombre);

  const objetivo = limpiar(nombre);
  const limpios = {};
  for (const n of nombres) limpios[limpiar(n)] = n;
  if (limpios[objetivo] !== undefined) {
    return mapeo.cc_blocks.find(b => b.nombre_balance === limpios[objetivo]);
  }
  const match = getCloseMatches(objetivo, Object.keys(limpios), 1, 0.6);
  if (match.length) return mapeo.cc_blocks.find(b => b.nombre_balance === limpios[match[0]]);
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

function categoriasDisponibles(mapeo) {
  return [...new Set(mapeo.categorias.map(c => c.desc))].sort();
}

// ---------------------------------------------------------------- hoja SyS

// Todas las filas de SyS que son dato (tienen la clave CC_cuenta en la columna B).
// Las filas de encabezado de centro de costo no tienen clave.
function filasDeDatosSys(wsSys) {
  const filas = new Map();
  for (let r = 1; r <= wsSys.rowCount; r++) {
    const clave = textoPlano(wsSys.getCell(r, 2));
    if (clave && clave.includes("_")) filas.set(clave.trim(), r);
  }
  return filas;
}

// Cada corrida es de UN mes. Si no se borran los importes del mes anterior, las
// cuentas sin movimiento este mes conservan el valor viejo (el VLOOKUP lo sigue
// encontrando) y el total no cierra.
function limpiarSys(wsSys, log = () => {}) {
  const filas = filasDeDatosSys(wsSys);
  let borradas = 0;
  for (const r of filas.values()) {
    for (const col of [11, 13, 16]) {   // K = Debe, M = Haber, P = Saldo
      const cell = wsSys.getCell(r, col);
      if (cell.value !== null && cell.value !== undefined && cell.value !== 0) borradas++;
      cell.value = 0;
    }
  }
  log(`  Limpiada la hoja SyS: ${borradas} importes del mes anterior puestos en 0 (${filas.size} filas).`);
  return borradas;
}

function escribirLineaSys(wsSys, filasSys, linea, codigo) {
  const clave = `${linea.cc_codigo}_${codigo}`;
  let fila = filasSys.get(clave);
  if (fila === undefined) {
    fila = wsSys.rowCount + 1;
    wsSys.getCell(fila, 1).value = `${codigo} - ${linea.cuenta_label}`;
    wsSys.getCell(fila, 2).value = clave;
    filasSys.set(clave, fila);
  }
  wsSys.getCell(fila, 11).value = linea.debe;
  wsSys.getCell(fila, 13).value = linea.haber;
  wsSys.getCell(fila, 16).value = linea.saldo;
  return fila;
}

// ------------------------------------------------- cuentas nuevas en el balance

// La fila que inserta ExcelJS viene vacía. Hay que copiarle las fórmulas de una
// fila vecina (VLOOKUP contra SyS, los saldos =+C-D y la columna de totales),
// porque si no la cuenta nueva queda en cero para siempre.
function copiarFormulasDeFila(wsSs, filaDestino, filaOrigen, log = () => {}) {
  const origen = wsSs.getRow(filaOrigen);
  let copiadas = 0;

  origen.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (colNumber <= 2) return;               // A = código, B = descripción
    const destino = wsSs.getCell(filaDestino, colNumber);
    destino.style = cell.style;

    const v = cell.value;
    if (v && typeof v === "object" && typeof v.formula === "string") {
      // Las fórmulas de estas filas solo se refieren a su propia fila, así que
      // alcanza con cambiar el número de fila donde aparezca.
      const nueva = v.formula.replace(
        new RegExp(`\\b([A-Z]{1,3})${filaOrigen}\\b`, "g"),
        (_, col) => `${col}${filaDestino}`
      );
      destino.value = { formula: nueva };
      copiadas++;
    } else if (typeof v === "number") {
      destino.value = 0;
    }
  });

  log(`  Fila ${filaDestino}: copiadas ${copiadas} fórmulas desde la fila ${filaOrigen}.`);
  return copiadas;
}

function agregarRefDistDeGastos(wsDist, distRow, distCol, colSaldoSs, ssRow) {
  const celda = wsDist.getCell(`${distCol}${distRow}`);
  const ref = `'Sumas y Saldos'!${colSaldoSs}${ssRow}`;
  const v = celda.value;

  if (v && typeof v === "object" && v.sharedFormula) {
    throw new Error(
      `La celda Dist.de gastos!${distCol}${distRow} usa una fórmula compartida y no se le puede ` +
      `agregar la referencia de la cuenta nueva de forma segura. NO se generó ningún archivo.`
    );
  }
  // ExcelJS guarda el texto de la fórmula sin el '=' inicial.
  if (v && typeof v === "object" && typeof v.formula === "string") {
    celda.value = { formula: `${v.formula}+${ref}` };
  } else {
    celda.value = { formula: `+${ref}` };
  }
}

// ---------------------------------------------------------------- etapa 1

// Este balance es de cuentas de RESULTADO: las 382 cuentas del archivo empiezan con 4.
// El export de Onvio es un balance de sumas y saldos completo, así que además trae
// cuentas patrimoniales — Caja, IVA Crédito Fiscal, anticipos de fondos, cuentas de
// proveedores — que no son gasto y no van ni a la distribución ni a la hoja SyS.
//
// Se apartan en un solo lugar, del que dependen las dos etapas: así no se pregunta por
// ellas como si fueran cuentas nuevas a clasificar (que era la única salida que daba la
// app, y metía el importe dentro de una categoría de gasto), no se les escribe fila en
// SyS, y no entran en el total que tiene que dar el balance.
function esCuentaDeResultado(codigo) {
  return /^4/.test(String(codigo).trim());
}

function separarCuentasDeResultado(lineas) {
  const deResultado = [], fueraDelBalance = [];
  for (const linea of lineas) {
    (esCuentaDeResultado(linea.cuenta_codigo) ? deResultado : fueraDelBalance).push(linea);
  }
  return { deResultado, fueraDelBalance };
}

// Junta las descartadas por cuenta, para poder mostrarlas y que quede claro qué quedó
// afuera en vez de que desaparezca sin aviso.
function resumirFueraDelBalance(lineas) {
  const porCuenta = new Map();
  for (const l of lineas) {
    const codigo = String(l.cuenta_codigo);
    if (!porCuenta.has(codigo)) {
      porCuenta.set(codigo, { codigo, label: l.cuenta_label, lineas: 0, saldo: 0 });
    }
    const c = porCuenta.get(codigo);
    c.lineas++;
    c.saldo += l.saldo || 0;
  }
  return [...porCuenta.values()].sort((a, b) => a.codigo.localeCompare(b.codigo));
}

function detectarPendientes(lineas, mapeo) {
  const pendientes = [];
  const sinCc = [];
  const vistas = new Set();

  const { deResultado, fueraDelBalance } = separarCuentasDeResultado(lineas);

  for (const linea of deResultado) {
    const ccBlock = resolverCcBlock(mapeo, linea.cc_nombre_onvio);
    if (ccBlock === null) { sinCc.push(linea.cc_nombre_onvio); continue; }

    const codigo = linea.cuenta_codigo;
    if (vistas.has(codigo)) continue;
    const cuenta = mapeo.cuentas[codigo];
    if (cuenta === undefined || !cuenta.categoria) {
      vistas.add(codigo);
      pendientes.push({
        codigo, label: linea.cuenta_label, cc_nombre: ccBlock.nombre_balance,
        saldo: linea.saldo, tipo: cuenta === undefined ? "nueva" : "sin_categoria",
      });
    }
  }
  return {
    pendientes,
    sinCc: [...new Set(sinCc)].sort(),
    fueraDelBalance: resumirFueraDelBalance(fueraDelBalance),
  };
}

// Centros de costo que traen movimiento este mes pero no tienen columna en
// Dist.de gastos: su gasto no va a aparecer en la distribución.
function detectarCcSinColumna(lineas, mapeo) {
  const afectados = {};
  for (const linea of lineas) {
    const b = resolverCcBlock(mapeo, linea.cc_nombre_onvio);
    if (!b) continue;
    if (ccNombreADistCol(mapeo, b.nombre_balance) === null) {
      if (!afectados[b.nombre_balance]) afectados[b.nombre_balance] = { lineas: 0, saldo: 0 };
      afectados[b.nombre_balance].lineas++;
      afectados[b.nombre_balance].saldo += linea.saldo;
    }
  }
  return Object.entries(afectados).map(([nombre, d]) => ({ nombre, ...d }));
}

// ---------------------------------------------------------------- etapa 2

function procesar({ wb, lineas, mapeo, categoriasElegidas = {}, periodo, log = () => {} }) {
  mapeo = JSON.parse(JSON.stringify(mapeo));

  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const wsSys = wb.getWorksheet("SyS");
  if (!wsSs || !wsDist || !wsSys) {
    throw new Error(
      "El archivo no tiene las hojas esperadas ('Sumas y Saldos', 'Dist.de gastos', 'SyS')."
    );
  }
  if (!periodo) throw new Error("Falta indicar el período que se está cargando.");
  const { mes } = parsearPeriodo(periodo);

  // Solo puede haber un mes "vivo" (el que sigue al movimiento del mes). Si quedó
  // otro sin cerrar, cargar ahora le escribiría los importes de este mes encima y
  // el TOTAL AÑO quedaría mal. Se corta antes de tocar nada.
  const otrosVivos = mesesVivos(wsDist).filter(m => m.mes !== mes);
  if (otrosVivos.length) {
    throw new Error(
      `El archivo tiene sin cerrar: ${otrosVivos.map(m => m.nombre).join(", ")}. ` +
      `Hay que cerrar ese mes antes de cargar ${nombreMes(mes)}, porque si no los ` +
      `importes nuevos se escriben encima de ese total. NO se generó ningún archivo.`
    );
  }

  // Excel recalcula al abrir: las fórmulas nuevas se escriben sin resultado.
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  log(`${lineas.length} líneas de cuenta leídas del export.`);

  // Solo las cuentas de resultado llegan al balance; el resto se aparta acá.
  const { deResultado, fueraDelBalance } = separarCuentasDeResultado(lineas);
  const fuera = resumirFueraDelBalance(fueraDelBalance);
  lineas = deResultado;
  if (fuera.length) {
    log(`  ${fueraDelBalance.length} línea(s) que no son de cuentas de resultado quedan afuera: ` +
        fuera.map(c => `${c.codigo} ${c.label} (${c.saldo.toFixed(2)})`).join(", "));
  }

  limpiarSys(wsSys, log);
  const filasSys = filasDeDatosSys(wsSys);

  let nuevas = 0, clasificadas = 0, conocidas = 0;
  const sinCc = [];

  for (const linea of lineas) {
    const ccBlock = resolverCcBlock(mapeo, linea.cc_nombre_onvio);
    if (ccBlock === null) { sinCc.push(linea.cc_nombre_onvio); continue; }

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
          `Dist.de gastos no está soportado; elegí una de la lista.`
        );
      }

      const insertAt = cat.ss_rows.length ? Math.max(...cat.ss_rows) + 1 : mapeo.fila_totales;
      const mod = insertRowSumasYSaldos(wsSs, wb, insertAt);
      log(`  Insertada fila ${insertAt} en Sumas y Saldos (${mod} referencias de fórmula reacomodadas).`);

      // el mapeo en memoria queda corrido una fila
      for (const info of Object.values(mapeo.cuentas)) {
        if (info.ss_row >= insertAt) info.ss_row += 1;
      }
      for (const c2 of mapeo.categorias) {
        c2.ss_rows = c2.ss_rows.map(r => (r >= insertAt ? r + 1 : r));
      }
      if (mapeo.fila_totales >= insertAt) mapeo.fila_totales += 1;

      wsSs.getCell(insertAt, 1).value = parseInt(codigo, 10);
      wsSs.getCell(insertAt, 2).value = linea.cuenta_label;
      // la fila de abajo es la que estaba antes en esta posición: sirve de molde
      copiarFormulasDeFila(wsSs, insertAt, insertAt + 1, log);

      cuenta = { ss_row: insertAt, label: linea.cuenta_label, categoria, ccs_en_dist: [] };
      mapeo.cuentas[codigo] = cuenta;
      cat.ss_rows.push(insertAt);
      nuevas++;
    } else if (!cuenta.categoria) {
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

    // Traba de seguridad: si la fila que dice el mapeo no tiene esa cuenta, el
    // archivo y el mapeo están desincronizados. Se corta antes de escribir.
    const rCheck = cuenta.ss_row;
    const codigoEnArchivo = valorNumerico(wsSs.getCell(rCheck, 1));
    if (codigoEnArchivo === null || parseInt(codigoEnArchivo, 10) !== parseInt(codigo, 10)) {
      throw new Error(
        `DESINCRONIZACIÓN DETECTADA: el mapeo dice que la cuenta ${codigo} está en la fila ` +
        `${rCheck} de "Sumas y Saldos", pero esa fila tiene la cuenta ${codigoEnArchivo}.\n` +
        `Puede que el archivo base guardado no sea el que corresponde a este mapeo.\n` +
        `NO SE GUARDÓ NADA.`
      );
    }

    // que la cuenta esté referenciada en Dist.de gastos para este centro de costo
    if (!cuenta.ccs_en_dist) cuenta.ccs_en_dist = [];
    if (!cuenta.ccs_en_dist.includes(ccBlock.nombre_balance)) {
      const distRow = filaDistDeCategoria(mapeo, cuenta.categoria);
      const distCol = ccNombreADistCol(mapeo, ccBlock.nombre_balance);
      if (distRow && distCol) {
        agregarRefDistDeGastos(wsDist, distRow, distCol, ccBlock.col_saldo, cuenta.ss_row);
        cuenta.ccs_en_dist.push(ccBlock.nombre_balance);
      }
    }

    escribirLineaSys(wsSys, filasSys, linea, codigo);
  }

  if (sinCc.length) {
    log(`\n⚠ ${sinCc.length} línea(s) con centro de costo no identificado: ` +
        `${[...new Set(sinCc)].sort().join(", ")}`);
  }

  activarColumnaMes(wsDist, mes, log);

  const ccSinColumna = detectarCcSinColumna(lineas, mapeo);
  if (ccSinColumna.length) {
    log(`\n⚠ Centros de costo con movimiento que no tienen columna en Dist.de gastos: ` +
        ccSinColumna.map(c => `${c.nombre} (${c.saldo.toFixed(2)})`).join(", "));
  }

  log(`\nResumen: ${conocidas} cuentas ya conocidas, ${clasificadas} recién clasificadas, ` +
      `${nuevas} cuentas nuevas insertadas.`);

  return {
    mapeo,
    resumen: {
      conocidas, clasificadas, nuevas,
      lineas: lineas.length,
      totalSaldo: lineas.reduce((s, l) => s + l.saldo, 0),
      sinCc: [...new Set(sinCc)].sort(),
      ccSinColumna,
      fueraDelBalance: fuera,
    },
  };
}

// Cierra el mes: reemplaza la fórmula de la columna del mes por el número que
// calculó Excel. Se hace sobre el archivo que la usuaria revisó y volvió a subir.
function aprobarMes({ wb, periodo, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  if (!wsDist) throw new Error("El archivo no tiene la hoja 'Dist.de gastos'.");
  const { mes } = parsearPeriodo(periodo);
  const congeladas = congelarColumnaMes(wsDist, mes, log);
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  return { congeladas };
}

if (typeof module !== "undefined") {
  const { getCloseMatches } = require("./similitud.js");
  const { insertRowSumasYSaldos } = require("./formula_utils.js");
  const meses = require("./meses.js");
  global.getCloseMatches = getCloseMatches;
  global.insertRowSumasYSaldos = insertRowSumasYSaldos;
  global.parsearPeriodo = meses.parsearPeriodo;
  global.activarColumnaMes = meses.activarColumnaMes;
  global.congelarColumnaMes = meses.congelarColumnaMes;
  global.mesesVivos = meses.mesesVivos;
  global.nombreMes = meses.nombreMes;
  module.exports = {
    procesar, aprobarMes, detectarPendientes, detectarCcSinColumna, categoriasDisponibles,
    esCuentaDeResultado, separarCuentasDeResultado, resumirFueraDelBalance,
    resolverCcBlock, limpiarSys, filasDeDatosSys, copiarFormulasDeFila,
    norm, limpiar, colAIndice, indiceACol,
  };
}
