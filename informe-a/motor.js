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

function agregarRefDistDeGastos(wsDist, distRow, distCol, colSaldoSs, ssRow, signo = "+") {
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
    // Si la referencia ya está, no se agrega de nuevo: sumaría el importe dos veces. El
    // texto de la fórmula manda sobre lo que diga el mapeo, que puede venir desactualizado.
    if (formulaTieneRef(v.formula, colSaldoSs, ssRow)) return false;
    celda.value = { formula: `${v.formula}${signo}${ref}` };
  } else {
    celda.value = { formula: `${signo}${ref}` };
  }
  return true;
}

function formulaTieneRef(formula, col, fila) {
  const re = new RegExp(`'Sumas y Saldos'!\\$?${col}\\$?${fila}(?!\\d)`);
  return re.test(formula);
}

// Saca de una celda de Dist.de gastos la referencia a una fila de 'Sumas y Saldos'.
// Es la inversa de agregarRefDistDeGastos, y hace falta para MOVER una cuenta de
// categoría: primero hay que sacarla de donde está hoy.
//
// Sólo sabe sacar referencias SUMADAS (`+'Sumas y Saldos'!E226`). Si la referencia
// aparece restada o dentro de una función, no la toca y corta con un error: la
// columna CR tiene fórmulas con signos y con referencias a otras celdas de la hoja,
// y recortarlas a ciegas daría un número mal sin que se note.
function quitarRefDistDeGastos(wsDist, distRow, distCol, colSs, ssRow) {
  const celda = wsDist.getCell(`${distCol}${distRow}`);
  const v = celda.value;
  if (v && typeof v === "object" && v.sharedFormula) {
    throw new Error(
      `La celda Dist.de gastos!${distCol}${distRow} usa una fórmula compartida. ` +
      `El archivo tiene que abrirse con abrirWorkbook(). NO se modificó nada.`
    );
  }
  if (!(v && typeof v === "object" && typeof v.formula === "string")) return 0;

  const original = v.formula;
  const re = new RegExp(`'Sumas y Saldos'!\\$?${colSs}\\$?${ssRow}(?!\\d)`);
  let out = original;
  const signos = [];
  for (;;) {
    const m = re.exec(out);
    if (!m) break;
    const i = m.index;
    const fin = i + m[0].length;
    let antes = out.slice(0, i).replace(/\s+$/, "");
    // El término puede venir sumado (`+ref`), restado (`-ref`) o con las dos cosas
    // (`+-ref`, que el archivo real usa en la columna CR). El signo se DEVUELVE para
    // que quien mueva la cuenta la vuelva a poner igual: si se reagregara sumando algo
    // que estaba restado, el total cambiaría sin que nadie lo note.
    let signo = "+";
    if (antes === "") {
      // la referencia arranca la fórmula
    } else if (antes.endsWith("-")) {
      signo = "-";
      antes = antes.slice(0, -1).replace(/\s+$/, "");
      if (antes.endsWith("+")) antes = antes.slice(0, -1);   // el caso `+-ref`
    } else if (antes.endsWith("+")) {
      antes = antes.slice(0, -1);
    } else {
      throw new Error(
        `En Dist.de gastos!${distCol}${distRow} la cuenta de la fila ${ssRow} está dentro de ` +
        `una función o multiplicando ("…${original.slice(Math.max(0, i - 14), fin + 4)}…"). ` +
        `No la toco: hay que corregir esa fórmula a mano. NO se modificó nada.`
      );
    }
    out = antes + out.slice(fin);
    signos.push(signo);
  }
  if (!signos.length) return { quitadas: 0, signos: [] };

  out = out.replace(/^\s*\+(?![-])/, "").replace(/[+\-]\s*$/, "").trim();
  // Si no quedó nada, la celda pasa a valer 0: era la única cuenta que sumaba.
  celda.value = out ? { formula: out } : 0;
  return { quitadas: signos.length, signos };
}

// Todas las celdas de Dist.de gastos que hoy suman esta fila de 'Sumas y Saldos'.
// Es la fuente de verdad de a qué categorías va una cuenta: el mapeo guarda una sola
// etiqueta, pero las fórmulas pueden tenerla repartida entre varias filas.
// Se busca CUALQUIER referencia a la fila de la cuenta, sea cual sea la columna de
// 'Sumas y Saldos' que lea. La versión anterior sólo miraba la columna SALDO de cada
// centro de costo, y por eso no veía las 38 referencias del archivo que leen el DEBE o
// el HABER (las columnas R y S, que las dos leen SHEYLA). Esas quedaban sin mover: la
// cuenta terminaba sumando en la categoría nueva Y en la vieja a la vez.
//
// La columna de 'Sumas y Saldos' se conserva tal cual al mover: cambiar de categoría
// cambia DÓNDE se suma el importe, no QUÉ celda se lee.
function refsDeCuentaEnDist(wsDist, mapeo, ssRow, filasCategoria) {
  const colCr = columnaCrDeDist(wsDist);
  const cols = Object.keys(mapeo.dist_col_to_cc);
  if (colCr) cols.push(colCr);
  const re = new RegExp(`'Sumas y Saldos'!\\$?([A-Z]{1,3})\\$?${ssRow}(?!\\d)`, "g");
  const encontradas = [];
  for (const distRow of filasCategoria) {
    for (const distCol of cols) {
      const f = wsDist.getCell(`${distCol}${distRow}`).value;
      if (!(f && typeof f === "object" && typeof f.formula === "string")) continue;
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(f.formula)) !== null) {
        encontradas.push({
          distRow, distCol, colSs: m[1],
          tipo: distCol === colCr ? "haber" : "saldo",
        });
      }
    }
  }
  return encontradas;
}

// La columna "CR" de Dist.de gastos junta los HABER de todos los centros de costo, y la
// columna "MOVIMIENTO MES DR" la suma: `SUM(F:V)+E`. Las columnas de centro traen el SALDO
// (debe − haber), así que saldo + haber = debe, y por eso DR da el débito del mes.
//
// El motor sólo agregaba referencias a la columna del centro. Si una categoría no tenía ya
// una fórmula en CR —47 de las 95 no la tienen— su haber no aparecía en ningún lado: la
// columna quedaba en cero y DR salía corto por el mismo importe.
function columnaCrDeDist(wsDist) {
  for (let r = 1; r <= 12; r++) {
    for (let c = 1; c <= Math.min(wsDist.columnCount, 30); c++) {
      const v = wsDist.getCell(r, c).value;
      const t = (v && typeof v === "object")
        ? (v.richText ? v.richText.map(x => x.text).join("") : String(v.result ?? ""))
        : String(v == null ? "" : v);
      if (t.trim().toUpperCase() === "CR") return wsDist.getColumn(c).letter;
    }
  }
  return null;
}

// ------------------------------------------------- crear una categoría nueva
//
// Hasta ahora crear categorías era un paso manual en Excel. Hace falta cuando una
// cuenta no encaja en ninguna de las que ya existen.
//
// Dónde se inserta: la fila TOTAL GASTOS suma rangos por columna, y esos rangos NO
// terminan todos en la misma fila (en el archivo real 16 columnas llegan hasta la 101,
// 15 hasta la 100 y una hasta la 102 — una inconsistencia que ya venía). Se inserta
// dentro del rango MÁS CORTO, así la categoría nueva queda comprendida por todos y
// ninguno hay que tocar. Insertar después del más corto la dejaría fuera del total en
// esas 15 columnas, en silencio.
function filaTotalGastosDeDist(wsDist) {
  for (let r = 8; r <= wsDist.rowCount; r++) {
    const t = textoPlano(wsDist.getCell(r, 3));
    if (t && t.toUpperCase().includes("TOTAL GASTOS")) return r;
  }
  return null;
}

function crearCategoriaEnDist({ wb, mapeo, nombre, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const limpio = String(nombre == null ? "" : nombre).trim();
  if (!limpio) throw new Error("La categoría necesita un nombre.");
  if (mapeo.categorias.some(c => c.desc === limpio)) {
    throw new Error(`La categoría "${limpio}" ya existe en Dist.de gastos.`);
  }
  const filaTot = filaTotalGastosDeDist(wsDist);
  if (!filaTot) throw new Error("No encontré la fila 'TOTAL GASTOS' en Dist.de gastos.");

  // el final del rango más corto de la fila de totales
  let finMinimo = null;
  for (let c = 4; c <= wsDist.columnCount; c++) {
    const v = wsDist.getCell(filaTot, c).value;
    if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
    for (const m of v.formula.matchAll(/[A-Z]{1,3}\d+:[A-Z]{1,3}(\d+)/g)) {
      const fin = parseInt(m[1], 10);
      if (finMinimo === null || fin < finMinimo) finMinimo = fin;
    }
  }
  if (finMinimo === null) throw new Error("La fila TOTAL GASTOS no tiene rangos que pueda leer.");
  const insertAt = finMinimo;

  if (typeof insertRowEn !== "function") {
    throw new Error("Falta formula_hojas.js: sin él no se pueden insertar filas en Dist.de gastos.");
  }
  const mod = insertRowEn(wb, "Dist.de gastos", insertAt);
  log(`  Insertada la fila ${insertAt} en Dist.de gastos (${mod} referencias reacomodadas).`);

  // el mapeo queda corrido una fila
  for (const c of mapeo.categorias) if (c.dist_row >= insertAt) c.dist_row += 1;

  // la fila modelo es la que ahora quedó debajo: la que ocupaba este lugar
  const modelo = insertAt + 1;
  const fila = wsDist.getRow(insertAt);
  wsDist.getRow(modelo).eachCell({ includeEmpty: false }, (cell, c) => {
    fila.getCell(c).style = cell.style;
  });

  wsDist.getCell(insertAt, 3).value = limpio;
  // MOVIMIENTO MES: la misma fórmula que las demás categorías, con su propia fila
  const dModelo = wsDist.getCell(modelo, 4).value;
  if (dModelo && typeof dModelo === "object" && typeof dModelo.formula === "string") {
    wsDist.getCell(insertAt, 4).value = {
      formula: dModelo.formula.replace(new RegExp(`\\b([A-Z]{1,3})${modelo}\\b`, "g"),
                                       (_, col) => `${col}${insertAt}`),
    };
  }
  // Columnas de mes: los meses YA CERRADOS van en 0 —la categoría no existía, así que
  // no tuvo movimiento— y el TOTAL AÑO copia su fórmula. Los meses todavía sin abrir
  // quedan vacíos, igual que en el resto de las categorías.
  //
  // Las columnas de CENTRO DE COSTO y la CR se saltean a propósito: ahí es donde viven
  // las referencias a las cuentas, y copiarlas de la fila modelo le daría a la categoría
  // nueva una copia de las cuentas de su vecina — o sea, sumar esos importes dos veces.
  // La categoría nace vacía y se llena moviéndole cuentas.
  const colCr = columnaCrDeDist(wsDist);
  const esDeCuentas = (c) => {
    const letra = indiceACol(c);
    return letra === colCr || Object.prototype.hasOwnProperty.call(mapeo.dist_col_to_cc, letra);
  };
  for (let c = 5; c <= wsDist.columnCount; c++) {
    if (esDeCuentas(c)) continue;
    const v = wsDist.getCell(modelo, c).value;
    if (v === null || v === undefined) continue;
    if (typeof v === "object" && typeof v.formula === "string") {
      wsDist.getCell(insertAt, c).value = {
        formula: v.formula.replace(new RegExp(`\\b([A-Z]{1,3})${modelo}\\b`, "g"),
                                   (_, col) => `${col}${insertAt}`),
      };
    } else if (typeof v === "number") {
      wsDist.getCell(insertAt, c).value = 0;
    }
  }

  const nueva = { desc: limpio, dist_row: insertAt, ss_rows: [] };
  mapeo.categorias.push(nueva);
  mapeo.categorias.sort((a, b) => a.dist_row - b.dist_row);
  log(`  Categoría "${limpio}" creada en la fila ${insertAt} de Dist.de gastos.`);
  return nueva;
}

// Índice de todo el tablero en UNA sola pasada: fila de 'Sumas y Saldos' -> filas de
// Dist.de gastos que la suman. El panel necesita el estado de las 282 cuentas juntas, y
// preguntarlo cuenta por cuenta con refsDeCuentaEnDist serían casi medio millón de
// lecturas de celda; así son unas mil setecientas.
// Se devuelven SEPARADOS el saldo y el haber, porque son dos problemas distintos:
// que el SALDO de una cuenta esté repartido en varias categorías (146 cuentas), y que
// el saldo esté en una sola pero su HABER —que vive en la columna CR— esté enganchado
// en otra (9 cuentas). Mezclarlos daba un número que no se entendía.
function mapaDeDistribucion(wsDist, mapeo) {
  const colsCc = Object.keys(mapeo.dist_col_to_cc);
  const colCr = columnaCrDeDist(wsDist);
  const saldo = new Map();
  const haber = new Map();
  const re = /'Sumas y Saldos'!\$?([A-Z]{1,3})\$?(\d+)/g;
  const anotar = (mapa, fila, distRow) => {
    if (!mapa.has(fila)) mapa.set(fila, new Set());
    mapa.get(fila).add(distRow);
  };
  for (const c of mapeo.categorias) {
    for (const col of colCr ? colsCc.concat([colCr]) : colsCc) {
      const v = wsDist.getCell(`${col}${c.dist_row}`).value;
      if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(v.formula)) !== null) {
        anotar(col === colCr ? haber : saldo, parseInt(m[2], 10), c.dist_row);
      }
    }
  }
  return { saldo, haber };
}

// Crea la fila de una cuenta que todavía no está en el balance, dentro del bloque de
// su categoría, y la registra en el mapeo. La usan los dos caminos que dan de alta una
// cuenta: la corrida mensual (cuando el export trae una cuenta nueva) y el panel de
// categorización (cuando se agrega una a mano). Es la misma función a propósito: son
// la misma operación y no pueden divergir.
function insertarCuentaEnBalance({ wb, wsSs, mapeo, codigo, label, categoria, log = () => {} }) {
  wsSs = wsSs || wb.getWorksheet("Sumas y Saldos");
  if (mapeo.cuentas[codigo]) {
    throw new Error(`La cuenta ${codigo} ya está en el balance, en la fila ${mapeo.cuentas[codigo].ss_row}.`);
  }
  const limpio = String(label == null ? "" : label).trim();
  if (!limpio) {
    // Sin nombre en la columna B el motor no la registra: quedaría invisible y la
    // corrida siguiente le insertaría OTRA fila. Es el caso de la fila 131.
    throw new Error("La cuenta necesita un nombre: sin nombre el motor no la reconoce y la duplicaría.");
  }
  const cat = resolverCategoriaDestino(mapeo, categoria);
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
  wsSs.getCell(insertAt, 2).value = limpio;
  // la fila de abajo es la que estaba antes en esta posición: sirve de molde
  copiarFormulasDeFila(wsSs, insertAt, insertAt + 1, log);

  const cuenta = { ss_row: insertAt, label: limpio, categoria: cat.desc, ccs_en_dist: [] };
  mapeo.cuentas[codigo] = cuenta;
  cat.ss_rows.push(insertAt);
  cat.ss_rows.sort((a, b) => a - b);
  return cuenta;
}

// ------------------------------------------------- edición de la categorización
//
// Estas tres operaciones son las que usa el panel "Configurar categorización". Todas
// trabajan sobre el workbook y sobre el mapeo A LA VEZ, a propósito: la etiqueta del
// mapeo sola no mueve ningún importe (sólo decide dónde se engancha una referencia
// nueva), así que cambiarla sin tocar las fórmulas daría la sensación de haber
// arreglado algo que sigue igual.

function _filasDeCategorias(mapeo) {
  return mapeo.categorias.map(c => c.dist_row);
}

// Elegir la categoría por su NOMBRE es ambiguo: en el archivo real hay 95 filas de
// categoría con sólo 91 nombres distintos — ALQUILERES VARIOS está en las filas 19 y 55,
// ACUERDO DE INVERSIONES en la 75 y la 80, ROBO en la 81 y la 101, GASTOS DE
// ESPECTROMETRÍA en la 85 y la 89. Buscando por nombre se agarraba siempre la primera,
// en silencio. Por eso lo que se pasa es la FILA, y el nombre queda como respaldo para
// no romper nada que todavía mande texto.
function resolverCategoriaDestino(mapeo, valor) {
  if (valor === null || valor === undefined || valor === "") return null;
  const n = Number(valor);
  if (Number.isInteger(n) && n > 0) {
    const porFila = mapeo.categorias.find(x => x.dist_row === n);
    if (porFila) return porFila;
  }
  const iguales = mapeo.categorias.filter(x => x.desc === valor);
  if (iguales.length > 1) {
    throw new Error(
      `"${valor}" figura ${iguales.length} veces en Dist.de gastos (filas ` +
      `${iguales.map(x => x.dist_row).join(", ")}), así que no se sabe a cuál va. ` +
      `Elegila desde "Configurar categorización", que distingue la fila.`
    );
  }
  return iguales[0] || null;
}

function _sincronizarSsRows(mapeo, ssRow, distRowDestino) {
  for (const c of mapeo.categorias) {
    const i = c.ss_rows.indexOf(ssRow);
    if (c.dist_row === distRowDestino) {
      if (i === -1) c.ss_rows.push(ssRow);
    } else if (i !== -1) {
      c.ss_rows.splice(i, 1);
    }
  }
  const d = mapeo.categorias.find(c => c.dist_row === distRowDestino);
  if (d) d.ss_rows.sort((a, b) => a - b);
}

// Mueve TODAS las referencias de una cuenta a una sola categoría. Como cada cuenta
// suma una vez por columna (verificado en el archivo real: 0 pares cuenta/columna
// repetidos), mover las referencias de fila no cambia ningún total: el importe sigue
// sumando en la misma columna, sólo que en el renglón de la categoría correcta.
function moverCuentaDeCategoria({ wb, mapeo, codigo, categoriaDestino, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const cuenta = mapeo.cuentas[codigo];
  if (!cuenta) throw new Error(`La cuenta ${codigo} no está en el mapeo.`);
  const destino = resolverCategoriaDestino(mapeo, categoriaDestino);
  if (!destino) {
    throw new Error(
      `"${categoriaDestino}" no es una categoría existente. Crear categorías nuevas ` +
      `sigue siendo un paso manual en el Excel.`
    );
  }

  const refs = refsDeCuentaEnDist(wsDist, mapeo, cuenta.ss_row, _filasDeCategorias(mapeo));
  const fuera = refs.filter(r => r.distRow !== destino.dist_row);

  // Primero se sacan TODAS y recién después se agregan: así una referencia que ya
  // estaba en el destino no se duplica ni se pierde. El signo con el que estaba cada
  // una se conserva: en la columna CR hay cuatro que están RESTADAS, y reagregarlas
  // sumando les cambiaría el signo al importe sin que se note.
  for (const r of fuera) {
    const q = quitarRefDistDeGastos(wsDist, r.distRow, r.distCol, r.colSs, cuenta.ss_row);
    r.signo = (q.signos && q.signos[0]) || "+";
  }
  for (const r of fuera) {
    agregarRefDistDeGastos(wsDist, destino.dist_row, r.distCol, r.colSs, cuenta.ss_row, r.signo);
  }

  const ccDeCol = {};
  for (const [col, info] of Object.entries(mapeo.dist_col_to_cc)) ccDeCol[col] = info.nombre_balance;
  cuenta.categoria = destino.desc;
  cuenta.excluida = false;
  cuenta.ccs_en_dist = [...new Set(refs.filter(r => r.tipo === "saldo").map(r => ccDeCol[r.distCol]))]
    .filter(Boolean);
  _sincronizarSsRows(mapeo, cuenta.ss_row, destino.dist_row);

  log(`${codigo} ${cuenta.label}: ${fuera.length} referencia(s) movidas a "${destino.desc}" (fila ${destino.dist_row}).`);
  return { movidas: fuera.length, yaEstaban: refs.length - fuera.length, categoria: destino.desc };
}

// Saca la cuenta de la distribución sin borrarle la fila. La fila queda en
// 'Sumas y Saldos' con su saldo, pero deja de sumar en ninguna categoría.
// NO se la borra del mapeo: si se la sacara, la corrida siguiente la vería como
// cuenta nueva y le insertaría OTRA fila, duplicando el importe. Queda marcada
// como excluida, que además es lo que evita que se vuelva a preguntar por ella.
function quitarCuentaDeDistribucion({ wb, mapeo, codigo, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const cuenta = mapeo.cuentas[codigo];
  if (!cuenta) throw new Error(`La cuenta ${codigo} no está en el mapeo.`);

  const refs = refsDeCuentaEnDist(wsDist, mapeo, cuenta.ss_row, _filasDeCategorias(mapeo));
  for (const r of refs) quitarRefDistDeGastos(wsDist, r.distRow, r.distCol, r.colSs, cuenta.ss_row);
  // (el signo no importa acá: la cuenta deja de sumar en ningún lado)

  cuenta.categoria = null;
  cuenta.excluida = true;
  cuenta.ccs_en_dist = [];
  for (const c of mapeo.categorias) {
    const i = c.ss_rows.indexOf(cuenta.ss_row);
    if (i !== -1) c.ss_rows.splice(i, 1);
  }
  log(`${codigo} ${cuenta.label}: ${refs.length} referencia(s) quitadas. Ya no se distribuye.`);
  return { quitadas: refs.length };
}

// Cambia el nombre que se ve en 'Sumas y Saldos'. Es seguro: el VLOOKUP de cada fila
// busca por el CÓDIGO de la columna A (`VLOOKUP("1_"&A131,…)`), así que la columna B
// es sólo texto y cambiarla no mueve ningún importe. Sí importa que NO quede vacía:
// el motor registra una cuenta sólo si tiene código Y nombre, así que una fila sin
// nombre queda invisible y la corrida siguiente la inserta de nuevo, duplicada.
function renombrarCuenta({ wb, mapeo, codigo, nombre, log = () => {} }) {
  const cuenta = mapeo.cuentas[codigo];
  if (!cuenta) throw new Error(`La cuenta ${codigo} no está en el mapeo.`);
  const limpio = String(nombre == null ? "" : nombre).trim();
  if (!limpio) throw new Error("El nombre no puede quedar vacío.");
  wb.getWorksheet("Sumas y Saldos").getCell(`B${cuenta.ss_row}`).value = limpio;
  const antes = cuenta.label;
  cuenta.label = limpio;
  log(`${codigo}: "${antes}" pasa a llamarse "${limpio}".`);
  return { antes, ahora: limpio };
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
    // Una cuenta excluida a propósito desde el panel no se vuelve a preguntar: la
    // decisión de dejarla afuera ya está tomada y guardada.
    if (cuenta && cuenta.excluida) continue;
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

function procesar({ wb, lineas, mapeo, categoriasElegidas = {}, excluidas = [], periodo, log = () => {} }) {
  mapeo = JSON.parse(JSON.stringify(mapeo));

  const wsSs = wb.getWorksheet("Sumas y Saldos");
  const wsDist = wb.getWorksheet("Dist.de gastos");
  const wsSys = wb.getWorksheet("SyS");
  if (!wsSs || !wsDist || !wsSys) {
    throw new Error(
      "El archivo no tiene las hojas esperadas ('Sumas y Saldos', 'Dist.de gastos', 'SyS')."
    );
  }
  const colCr = columnaCrDeDist(wsDist);
  const haberesAgregados = [];
  if (!colCr) {
    log("\n⚠ No encontré la columna 'CR' en Dist.de gastos: los haberes no se van a cargar ahí.");
  }
  if (!periodo) throw new Error("Falta indicar el período que se está cargando.");
  const { anio, mes } = parsearPeriodo(periodo);

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

  // El encabezado de "Sumas y Saldos" (A2, formato mmm-yy) lleva el mes del informe y nadie
  // lo actualizaba: el maestro venía diciendo "abr-26" con el archivo ya en julio. Es el
  // título de la hoja principal, así que no puede quedar en el mes de cuando se armó.
  fecharSumasYSaldos(wsSs, anio, mes, log);

  log(`${lineas.length} líneas de cuenta leídas del export.`);

  // Solo las cuentas de resultado llegan al balance; el resto se aparta acá.
  const { deResultado, fueraDelBalance } = separarCuentasDeResultado(lineas);
  const fuera = resumirFueraDelBalance(fueraDelBalance);
  lineas = deResultado;
  if (fuera.length) {
    log(`  ${fueraDelBalance.length} línea(s) que no son de cuentas de resultado quedan afuera: ` +
        fuera.map(c => `${c.codigo} ${c.label} (${c.saldo.toFixed(2)})`).join(", "));
  }

  // Cuentas que se decidió NO incluir: o bien se contestó "no incluirla" en la pregunta
  // de cuentas nuevas, o bien se las sacó desde el panel de categorización. No se les
  // escribe línea en SyS ni se les crea fila, así que su importe queda fuera del balance
  // a propósito — y por eso tampoco tiene que entrar en el total, que es la cifra de
  // control. Quedan anotadas en el mapeo para no volver a preguntar por ellas.
  for (const codigo of excluidas) {
    const c = mapeo.cuentas[codigo];
    if (c) {
      c.excluida = true;
      c.categoria = null;
    } else {
      const l = lineas.find(x => x.cuenta_codigo === codigo);
      mapeo.cuentas[codigo] = {
        ss_row: null, label: (l && l.cuenta_label) || "", categoria: null,
        excluida: true, ccs_en_dist: [],
      };
    }
  }
  const esExcluida = (l) => !!(mapeo.cuentas[l.cuenta_codigo] || {}).excluida;
  const dejadasAfuera = resumirFueraDelBalance(lineas.filter(esExcluida));
  lineas = lineas.filter(l => !esExcluida(l));
  if (dejadasAfuera.length) {
    log(`  ${dejadasAfuera.length} cuenta(s) dejadas afuera a propósito: ` +
        dejadasAfuera.map(c => `${c.codigo} ${c.label} (${c.saldo.toFixed(2)})`).join(", "));
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
      cuenta = insertarCuentaEnBalance({
        wb, wsSs, mapeo, codigo, label: linea.cuenta_label, categoria, log,
      });
      nuevas++;
    } else if (!cuenta.categoria) {
      const categoria = categoriasElegidas[codigo];
      if (!categoria) {
        throw new Error(`Falta elegir la categoría de la cuenta ${codigo} - ${linea.cuenta_label}.`);
      }
      const cat = resolverCategoriaDestino(mapeo, categoria);
      if (!cat) throw new Error(`"${categoria}" no es una categoría existente.`);
      cuenta.categoria = cat.desc;
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

    // y que su HABER esté referenciado en la columna CR, que es la que los junta
    if (colCr && Number(linea.haber)) {
      const distRow = filaDistDeCategoria(mapeo, cuenta.categoria);
      if (distRow) {
        if (agregarRefDistDeGastos(wsDist, distRow, colCr, ccBlock.col_haber, cuenta.ss_row)) {
          haberesAgregados.push(
            `${codigo} (${cuenta.categoria}) — haber de ${ccBlock.nombre_balance}: ` +
            `${Number(linea.haber).toFixed(2)}`);
        }
      }
    }

    escribirLineaSys(wsSys, filasSys, linea, codigo);
  }

  if (haberesAgregados.length) {
    log(`\nSe engancharon ${haberesAgregados.length} haber(es) a la columna CR de ` +
        `Dist.de gastos, que antes no los traía:`);
    haberesAgregados.forEach(h => log(`  ${h}`));
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
      dejadasAfuera,
    },
  };
}

// Cierra el mes: reemplaza la fórmula de la columna del mes por el número que
// calculó Excel. Se hace sobre el archivo que la usuaria revisó y volvió a subir.
const MESES_ACUM = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
                    "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// Lo que gastó cada centro de costo en el mes: es lo mismo que la fila TOTAL GASTOS de su
// columna en Dist.de gastos, pero calculado acá a partir de las líneas del export en vez de
// leerlo del archivo. Leerlo sería depender del último resultado que Excel dejó guardado, y
// si el archivo no se abrió después de la corrida anterior ese número es de otro mes.
function totalesPorProyecto(lineas, mapeo) {
  const total = {};
  for (const l of separarCuentasDeResultado(lineas).deResultado) {
    const bloque = resolverCcBlock(mapeo, l.cc_nombre_onvio);
    if (!bloque) continue;
    total[bloque.nombre_balance] = (total[bloque.nombre_balance] || 0) + (Number(l.saldo) || 0);
  }
  return total;
}

// Pone el mes del informe en el encabezado de "Sumas y Saldos". La celda es una FECHA real
// con formato `mmm-yy`, no un texto, así que se escribe como fecha: si se le pusiera "jul-26"
// Excel lo tomaría como texto y perdería el formato. Se usa el día 1 porque lo que se muestra
// es el mes, y sólo se toca si la celda ya era una fecha — si alguien puso otra cosa ahí, se
// avisa en vez de pisarla.
function fecharSumasYSaldos(ws, anio, mes, log = () => {}) {
  const celda = ws.getCell("A2");
  const previo = celda.value;
  if (previo instanceof Date && !isNaN(previo.getTime())) {
    if (previo.getUTCFullYear() === anio && previo.getUTCMonth() + 1 === mes) return null;
    celda.value = new Date(Date.UTC(anio, mes - 1, 1));
    log(`\nEl encabezado de 'Sumas y Saldos' pasó de ${previo.getUTCMonth() + 1}/` +
        `${previo.getUTCFullYear()} a ${mes}/${anio}.`);
    return { antes: previo, despues: celda.value };
  }
  log(`\n⚠ El encabezado de 'Sumas y Saldos' (A2) no es una fecha, así que lo dejé como ` +
      `está. Hay que ponerle el mes a mano.`);
  return null;
}

// La fila de TOTALES de "Gastos Acumulados" trae el total del acumulado del año como un
// número fijo, no como fórmula: sus cuatro vecinas (años anteriores, mes, total y acumulado a
// la fecha) sí son =SUM(). Como el motor reescribe la columna del acumulado fila por fila, ese
// número queda con el total de la corrida anterior y la hoja muestra importes correctos que no
// suman. Es el mismo caso que la fila TOTALES y la columna BQ de "Sumas y Saldos".
//
// El rango NO se hardcodea: se copia del de una vecina de la MISMA fila, así sigue solo a
// cualquier inserción de filas, igual que lo hace Excel con las otras cuatro.
function repararTotalGastosAcumulados(ws, colAcum) {
  const RANGO = /SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?\1\$?(\d+)\s*\)/i;
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      if (c === colAcum) continue;
      const f = ws.getCell(r, c).formula;
      const m = f && String(f).match(RANGO);
      if (!m) continue;

      const desde = Number(m[2]), hasta = Number(m[3]);
      const letra = ws.getColumn(colAcum).letter;
      const celda = ws.getCell(r, colAcum);
      const previo = typeof celda.value === "number" ? celda.value : null;

      // el resultado se calcula acá además de escribir la fórmula, así el total ya es el bueno
      // sin depender de que el archivo pase por Excel. Si alguna celda del rango es una fórmula
      // sin resultado guardado no se inventa nada: se deja que lo calcule Excel al abrir.
      let suma = 0, calculable = true;
      for (let fila = desde; fila <= hasta; fila++) {
        const v = ws.getCell(fila, colAcum).value;
        if (v === null || v === undefined) continue;
        if (typeof v === "number") { suma += v; continue; }
        if (typeof v === "object" && typeof v.result === "number") { suma += v.result; continue; }
        if (typeof v === "object" && v.formula !== undefined) { calculable = false; break; }
      }
      const formula = `SUM(${letra}${desde}:${letra}${hasta})`;
      celda.value = calculable
        ? { formula, result: Math.round(suma * 100) / 100 }
        : { formula };

      return { fila: r, formula, previo, nuevo: calculable ? Math.round(suma * 100) / 100 : null };
    }
  }
  return null;
}

// Pasa el mes que se cierra a la columna del acumulado del año y corre los rótulos al mes
// siguiente: cerrando junio, "ENERO - MAYO" pasa a "ENERO - JUNIO" y la columna del mes
// pasa a decir "JULIO".
//
// La columna "ACUMULADO AÑOS ANTERIORES" NO se toca nunca: es de la usuaria.
function avanzarGastosAcumulados({ wb, mapeo, lineas, periodo, log = () => {} }) {
  const ws = wb.getWorksheet("Gastos Acumulados");
  if (!ws) return { filas: 0 };
  const { anio, mes } = parsearPeriodo(periodo);

  if (mes === 12) {
    log("\n⚠ Se está cerrando diciembre: el pase del acumulado del año a 'ACUMULADO AÑOS " +
        "ANTERIORES' es un cambio de ejercicio y se hace a mano. La hoja 'Gastos Acumulados' " +
        "queda como está.");
    return { filas: 0 };
  }

  // La hoja dice de qué mes es su columna "mes en curso". Si no es el que se está cerrando,
  // el acumulado no está donde debería y sumarle el mes lo cuenta dos veces —o saltea uno—
  // sin que nada lo delate: los rótulos quedan iguales igual. Pasó de verdad al sembrar la
  // hoja con un informe terminado sin registrar además ese mes como cerrado.
  const rotuloMes = String(ws.getCell(11, 4).value || "").trim().toUpperCase();
  if (rotuloMes && rotuloMes !== MESES_ACUM[mes - 1]) {
    throw new Error(
      `"Gastos Acumulados" tiene el mes en curso en ${rotuloMes} y se está cerrando ` +
      `${MESES_ACUM[mes - 1]}. Si sumara el mes ahí, el acumulado del año quedaría mal ` +
      `(contaría un mes dos veces o saltearía uno) y nada lo mostraría. NO se cerró el mes: ` +
      `primero hay que dejar esa hoja en ${MESES_ACUM[mes - 1]}.`
    );
  }

  const totales = totalesPorProyecto(lineas, mapeo);
  let filas = 0;
  const sinProyecto = [], conFila = new Set();
  for (let r = 1; r <= ws.rowCount; r++) {
    const nombre = String(ws.getCell(r, 1).value || "").trim();
    if (!nombre || !ws.getCell(r, 4).formula) continue;      // fila de proyecto: A + fórmula en D
    const bloque = resolverCcBlock(mapeo, nombre);
    if (!bloque) { sinProyecto.push(nombre); continue; }
    conFila.add(bloque.nombre_balance);
    const delMes = totales[bloque.nombre_balance] || 0;
    const acum = ws.getCell(r, 3);
    const previo = typeof acum.value === "number" ? acum.value : 0;
    acum.value = Math.round((previo + delMes) * 100) / 100;
    filas++;
  }

  // Un centro de costo que gastó este mes y no tiene fila en la hoja es plata que se pierde sin
  // dejar rastro: no entra al acumulado y el total tampoco la muestra. Ese es el aviso que
  // importa. Las filas de `sinProyecto` son proyectos dados de baja: mientras no traigan saldo
  // no hay nada que hacer con ellas, así que se informan sin alarma.
  const sinFila = Object.entries(totales)
    .filter(([nombre, saldo]) => Math.abs(saldo) >= 0.005 && !conFila.has(nombre))
    .map(([nombre, saldo]) => ({ nombre, saldo: Math.round(saldo * 100) / 100 }))
    .sort((a, b) => Math.abs(b.saldo) - Math.abs(a.saldo));

  // los rótulos pasan al mes siguiente
  ws.getCell(3, 1).value = `${MESES_ACUM[mes]} ${anio}`;
  ws.getCell(11, 3).value = `ENERO - ${MESES_ACUM[mes - 1]}`;
  ws.getCell(11, 4).value = MESES_ACUM[mes];

  const total = repararTotalGastosAcumulados(ws, 3);

  log(`\nGastos Acumulados: ${MESES_ACUM[mes - 1]} pasó al acumulado del año en ${filas} ` +
      `proyecto(s). Ahora dice "ENERO - ${MESES_ACUM[mes - 1]}" y el mes en curso es ` +
      `${MESES_ACUM[mes]}. La columna de años anteriores no se tocó.`);
  if (total) {
    log(`  El total del acumulado (fila ${total.fila}) pasó a ser =${total.formula}` +
        (total.previo !== null ? `, que antes era el número fijo ${total.previo.toFixed(2)}` : "") +
        (total.nuevo !== null ? ` y ahora da ${total.nuevo.toFixed(2)}` : "") + ".");
  } else {
    log("  ⚠ No encontré la fila de totales de la hoja: revisá a mano que el total del " +
        "acumulado sume la columna entera.");
  }
  if (sinFila.length) {
    log(`  ⚠ Gastaron este mes y NO tienen fila en la hoja "Gastos Acumulados" del archivo ` +
        `original, así que su gasto quedó afuera del acumulado: ` +
        sinFila.map(x => `${x.nombre} (${x.saldo.toFixed(2)})`).join(", ") +
        `. Hay que agregarles la fila a mano.`);
  }
  if (sinProyecto.length) {
    log(`  Sin centro de costo que les corresponda, no vinieron con saldo: ` +
        `${sinProyecto.join(", ")}. Quedaron como estaban.`);
  }
  return { filas, sinProyecto, sinFila, total };
}

// Deshace el cierre de un mes: su columna vuelve a seguir al movimiento (=D<fila>) en vez de
// tener el importe fijo, así se lo puede volver a cargar. Es el inverso de congelarColumnaMes.
//
// Lo que NO deshace, a propósito: las cuentas que se hayan dado de alta en ese mes quedan en
// el mapeo y en el archivo (borrarlas correría todas las filas de vuelta), y el acumulado de
// "Gastos Acumulados" no se descuenta — si el mes que se reabre ya se había acumulado, hay
// que volver a sembrarlo con el informe terminado.
function reabrirMes({ wb, periodo, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  if (!wsDist) throw new Error("El archivo no tiene la hoja 'Dist.de gastos'.");
  const { mes } = parsearPeriodo(periodo);

  const vivosAntes = mesesVivos(wsDist).filter(m => m.mes !== mes);
  if (vivosAntes.length) {
    throw new Error(
      `No puedo reabrir ${nombreMes(mes)}: ya hay otro mes abierto ` +
      `(${vivosAntes.map(v => v.nombre).join(", ")}). Cerralo o descartalo primero. ` +
      `NO se guardó nada.`
    );
  }

  const filas = activarColumnaMes(wsDist, mes, log);
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  log(`  ${nombreMes(mes)} quedó abierto de nuevo: se puede volver a cargar su export.`);
  return { mes, filas };
}

function aprobarMes({ wb, periodo, mapeo, lineas, log = () => {} }) {
  const wsDist = wb.getWorksheet("Dist.de gastos");
  if (!wsDist) throw new Error("El archivo no tiene la hoja 'Dist.de gastos'.");
  const { mes } = parsearPeriodo(periodo);
  const congeladas = congelarColumnaMes(wsDist, mes, log);
  // el mes que se cierra pasa al acumulado del año de "Gastos Acumulados"
  let acumulados = { filas: 0 };
  if (mapeo && lineas) {
    acumulados = avanzarGastosAcumulados({ wb, mapeo, lineas, periodo, log });
  } else {
    log("\n⚠ No se actualizó 'Gastos Acumulados': hacen falta las líneas del export y el mapeo.");
  }
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  return { congeladas, acumulados };
}

if (typeof module !== "undefined") {
  // crearCategoriaEnDist inserta filas en 'Dist.de gastos' y para eso usa insertRowEn,
  // que vive en informe-c/formula_hojas.js (en el navegador lo carga el index.html).
  if (typeof insertRowEn === "undefined") {
    try { global.insertRowEn = require("../informe-c/formula_hojas.js").insertRowEn; } catch (e) {}
  }
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
    columnaCrDeDist, agregarRefDistDeGastos, formulaTieneRef, filaDistDeCategoria,
    quitarRefDistDeGastos, refsDeCuentaEnDist, insertarCuentaEnBalance, mapaDeDistribucion,
    crearCategoriaEnDist, filaTotalGastosDeDist,
    resolverCategoriaDestino,
    moverCuentaDeCategoria, quitarCuentaDeDistribucion, renombrarCuenta,
    avanzarGastosAcumulados, totalesPorProyecto, reabrirMes,
    repararTotalGastosAcumulados, fecharSumasYSaldos,
  };
}
