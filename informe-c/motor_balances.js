// Motor del Informe C — Balance Pesos y Dólares.
//
// Un solo motor parametrizado por moneda (decisión de la usuaria). Por ahora la app
// solo corre PESOS: el archivo de dólares usa códigos de cuenta viejos y sus
// equivalencias todavía no están validadas (ver equivalencias_dolares_propuestas.json).
//
// Flujo dentro del maestro: Hoja1 es la zona de pegado del export (clave
// "código - nombre" en A, saldo en E para pesos); SALDOS busca cada cuenta contra
// Hoja1 con VLOOKUP; las hojas de estados (Activo y Pasivo, Anexo II, Balance...)
// se cuelgan de SALDOS.

const PARAMS = {
  pesos: { hoja1ColClave: 1, hoja1ColValor: 5, saldosColsCuenta: [3, 4, 5], saldosColValor: 7, campoSaldo: "saldo_ars" },
  dolares: { hoja1ColClave: 1, hoja1ColValor: 4, saldosColsCuenta: [3], saldosColValor: 3, campoSaldo: "saldo_usd" },
};

const RE_CUENTA_TXT = /^\s*(\d{6,})\s*-\s*(.+?)\s*$/;

function textoPlano(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined) return textoPlano(v.result);
    if (v.formula !== undefined) return "";
    return "";
  }
  return String(v);
}

// ------------------------------------------------- mapeo del maestro, en el momento
//
// Igual que el Informe A: la correspondencia se deriva del archivo al abrirlo, no de
// un JSON preparado de antemano que se puede desincronizar. El JSON extraído
// (mapeo_balances_pesos_dolares.json) queda como referencia y para las listas de la
// interfaz, pero la verdad es siempre el archivo.
function derivarMapeoMaestro(wb, moneda) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El maestro no tiene la hoja 'SALDOS'.");

  // El capítulo sale del primer dígito del código (1=ACTIVO ... 4=RESULTADOS), no de
  // títulos de sección: SALDOS no los tiene de forma consistente (en el archivo real
  // de pesos solo aparece un "ACTIVO" suelto en C2). La regla del dígito está
  // verificada contra las 163 cuentas del export de junio 2026 sin excepciones.
  const CAP_POR_DIGITO = { 1: "ACTIVO", 2: "PASIVO", 3: "PATRIMONIO NETO", 4: "RESULTADOS" };
  const cuentas = {};       // codigo -> {fila, col, clave, capitulo}
  const duplicadas = [];

  for (let r = 1; r <= ws.rowCount; r++) {
    for (const c of p.saldosColsCuenta) {
      const t = textoPlano(ws.getCell(r, c).value).trim();
      if (!t) continue;
      const m = RE_CUENTA_TXT.exec(t);
      if (!m) continue;
      const codigo = m[1];
      if (cuentas[codigo]) {
        duplicadas.push({ codigo, fila: r, filaPrevia: cuentas[codigo].fila });
      } else {
        cuentas[codigo] = { fila: r, col: c, clave: t, nombre: m[2], capitulo: CAP_POR_DIGITO[codigo[0]] || null };
      }
      break;              // la primera columna con texto de cuenta define la fila
    }
  }

  return { moneda, cuentas, duplicadas };
}

// --------------------------------------------------------------------- Hoja1
//
// Cada corrida REEMPLAZA la zona de pegado, no acumula: se limpia todo lo viejo y se
// escribe el export fresco con el mismo layout (título de capítulo, y debajo cada
// cuenta con su clave "código - nombre" y su saldo).
function reescribirHoja1(wb, cuentasExport, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("Hoja1");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Hoja1' (la zona de pegado del export).");

  const filasViejas = ws.rowCount;
  for (let r = 1; r <= filasViejas; r++) {
    ws.getCell(r, p.hoja1ColClave).value = null;
    ws.getCell(r, p.hoja1ColValor).value = null;
  }

  let r = 1;
  let capituloPrevio = null;
  for (const c of cuentasExport) {
    if (c.capitulo !== capituloPrevio) {
      ws.getCell(r, p.hoja1ColClave).value = c.capitulo;
      capituloPrevio = c.capitulo;
      r++;
    }
    ws.getCell(r, p.hoja1ColClave).value = `${c.codigo} - ${c.nombre}`;
    ws.getCell(r, p.hoja1ColValor).value = c[p.campoSaldo];
    r++;
  }
  const ultimaFila = r - 1;
  log(`  Hoja1 reescrita: ${cuentasExport.length} cuentas (filas 1 a ${ultimaFila}, antes ${filasViejas}).`);
  return ultimaFila;
}

// Los VLOOKUP de SALDOS apuntan a rangos fijos de Hoja1 que quedaron de corridas
// viejas: en el maestro real conviven $E$377 (sano) con $E$47/$E$50/$E$53, que cubren
// ~50 filas. Toda cuenta que caiga después de esa fila desaparece EN SILENCIO (el
// IFERROR la deja vacía). Como la app acaba de reescribir Hoja1 y sabe exactamente
// hasta qué fila hay datos, normaliza todos los rangos cortos a la extensión real.
function normalizarRangosVlookup(wb, moneda, ultimaFilaHoja1, log = () => {}) {
  const p = PARAMS[moneda];
  const colValor = String.fromCharCode(64 + p.hoja1ColValor);   // E en pesos, D en dólares
  const RE = new RegExp(`(Hoja1!\\$A\\$(\\d+):\\$${colValor}\\$)(\\d+)`, "gi");
  const ws = wb.getWorksheet("SALDOS");
  let corregidas = 0;

  ws.eachRow(row => row.eachCell(cell => {
    const v = cell.value;
    if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
    let cambio = false;
    const nueva = v.formula.replace(RE, (m, pre, desde, fin) => {
      if (parseInt(fin, 10) >= ultimaFilaHoja1) return m;
      cambio = true;
      return `${pre}${ultimaFilaHoja1}`;
    });
    if (cambio) { cell.value = { formula: nueva }; corregidas++; }
  }));

  if (corregidas) log(`  ${corregidas} VLOOKUP de SALDOS tenían el rango corto y se extendieron hasta la fila ${ultimaFilaHoja1}.`);
  return corregidas;
}

// ------------------------------------------------------------- cuentas nuevas

function detectarNuevas(cuentasExport, mapeoMaestro) {
  return cuentasExport.filter(c => !mapeoMaestro.cuentas[c.codigo]);
}

// Copia el patrón de fórmula de la fila vecina, cambiando solo el número de fila.
// El $ de las referencias absolutas ($E$377) queda entre la letra y el número, así
// que el reemplazo \b<COL><FILA>\b no las toca — es exactamente lo que se busca.
function copiarPatronFila(ws, filaOrigen, filaDestino, cols) {
  let copiadas = 0;
  for (const c of cols) {
    const origen = ws.getCell(filaOrigen, c);
    const destino = ws.getCell(filaDestino, c);
    destino.style = origen.style;
    const v = origen.value;
    if (v && typeof v === "object" && typeof v.formula === "string") {
      const nueva = v.formula.replace(
        new RegExp(`\\b([A-Z]{1,3})${filaOrigen}\\b`, "g"),
        (_, col) => `${col}${filaDestino}`
      );
      destino.value = { formula: nueva };
      copiadas++;
    }
  }
  return copiadas;
}

// La fórmula de saldo de una fila de SALDOS, si la tiene (VLOOKUP en G para el
// activo, en F para resultados) — o null si el importe está escrito a mano.
function colFormulaDeFila(ws, fila, p) {
  for (const c of [p.saldosColValor, p.saldosColValor - 1]) {
    const v = ws.getCell(fila, c).value;
    if (v && typeof v === "object" && typeof v.formula === "string") return c;
  }
  return null;
}

// La columna donde la fila guarda su importe a mano, si es una fila manual.
function colValorManual(ws, fila, p) {
  for (const c of [p.saldosColValor, p.saldosColValor - 1]) {
    const v = ws.getCell(fila, c).value;
    if (typeof v === "number") return c;
  }
  return null;
}

// Inserta la cuenta nueva en SALDOS junto a las cuentas de código más parecido: la
// vecina es la del prefijo común más largo (y de ellas, la de fila más alta). Eso
// mete cada proveedor nuevo (2110xxxxx) dentro del bloque de proveedores y cada
// gasto cerca de su familia — importante porque los subtotales son rangos
// (SUM(G63:G114), SUM(F219:F226)...) que se expanden solos cuando la fila nueva cae
// adentro, cosa que el shift de fórmulas ya garantiza.
//
// El patrón de la fila se replica tal cual sea el de la vecina: si tiene VLOOKUP se
// copia la fórmula ajustada; si es un importe a mano (todo el detalle de proveedores
// de pesos es así), se escribe el saldo fresco del export como número.
function insertarCuentaEnSaldos(wb, mapeoMaestro, cuenta, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");

  const delCapitulo = Object.values(mapeoMaestro.cuentas)
    .filter(x => x.capitulo === cuenta.capitulo);
  if (!delCapitulo.length) {
    throw new Error(
      `No encontré ninguna cuenta del capítulo ${cuenta.capitulo} en SALDOS para usar ` +
      `de modelo al insertar ${cuenta.codigo}. NO se generó ningún archivo.`
    );
  }
  const prefijoComun = (a, b) => {
    let i = 0;
    while (i < a.length && i < b.length && a[i] === b[i]) i++;
    return i;
  };
  const vecina = delCapitulo.reduce((a, b) => {
    const pa = prefijoComun(a.clave, cuenta.codigo), pb = prefijoComun(b.clave, cuenta.codigo);
    if (pa !== pb) return pa > pb ? a : b;
    return a.fila >= b.fila ? a : b;
  });
  const filaNueva = vecina.fila + 1;

  const mod = insertRowEn(wb, "SALDOS", filaNueva);
  // el mapeo en memoria queda corrido una fila
  for (const info of Object.values(mapeoMaestro.cuentas)) {
    if (info.fila >= filaNueva) info.fila += 1;
  }

  const clave = `${cuenta.codigo} - ${cuenta.nombre}`;
  ws.getCell(filaNueva, vecina.col).value = clave;
  // la hoja usa una segunda columna con el mismo texto cuando la vecina la tiene
  const colDesc = vecina.col + 1;
  if (p.saldosColsCuenta.includes(colDesc) &&
      textoPlano(ws.getCell(vecina.fila, colDesc).value).trim()) {
    ws.getCell(filaNueva, colDesc).value = clave;
  }

  const colF = colFormulaDeFila(ws, vecina.fila, p);
  let colValor;
  if (colF !== null) {
    copiarPatronFila(ws, vecina.fila, filaNueva, [p.saldosColValor, p.saldosColValor - 1]);
    colValor = colF;
    log(`  ${clave} insertada en SALDOS fila ${filaNueva}, con la fórmula de la vecina "${vecina.clave}" (${mod} referencias reacomodadas).`);
  } else {
    colValor = colValorManual(ws, vecina.fila, p) || p.saldosColValor;
    ws.getCell(filaNueva, colValor).value = cuenta[p.campoSaldo];
    ws.getCell(filaNueva, colValor).style = ws.getCell(vecina.fila, colValor).style;
    log(`  ${clave} insertada en SALDOS fila ${filaNueva}, con el importe directo como su bloque (${mod} referencias reacomodadas).`);
  }

  mapeoMaestro.cuentas[cuenta.codigo] = {
    fila: filaNueva, col: vecina.col, clave, nombre: cuenta.nombre, capitulo: cuenta.capitulo,
  };
  return { filaNueva, colValor };
}

// Las filas de SALDOS sin fórmula (el detalle de proveedores en pesos) no se
// actualizan solas desde Hoja1: hay que escribirles el saldo fresco del export en
// cada corrida, o el balance mostraría los importes del mes pasado.
function actualizarSaldosManuales(wb, mapeoMaestro, cuentasExport, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  let actualizadas = 0;
  for (const c of cuentasExport) {
    const info = mapeoMaestro.cuentas[c.codigo];
    if (!info) continue;
    if (colFormulaDeFila(ws, info.fila, p) !== null) continue;   // esa se actualiza sola
    const col = colValorManual(ws, info.fila, p) || p.saldosColValor;
    const nuevo = c[p.campoSaldo];
    const viejo = ws.getCell(info.fila, col).value;
    if (viejo !== nuevo) { ws.getCell(info.fila, col).value = nuevo; actualizadas++; }
  }
  if (actualizadas) log(`  ${actualizadas} filas de SALDOS con importe a mano actualizadas con el saldo del export.`);
  return actualizadas;
}

// RESULTADOS: la cuenta nueva se suma al concepto elegido del Anexo II, en la columna
// elegida (F=Administración, G=Comercialización, H=Exploración, I=Financieros).
function agregarRefAnexo2(wb, moneda, filaConcepto, colLetra, filaSaldos, colSaldos, log = () => {}) {
  const ws = wb.getWorksheet("Anexo II");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Anexo II'.");
  const celda = ws.getCell(`${colLetra}${filaConcepto}`);
  const ref = `SALDOS!${String.fromCharCode(64 + colSaldos)}${filaSaldos}`;
  const v = celda.value;
  if (v && typeof v === "object" && typeof v.formula === "string") {
    celda.value = { formula: `${v.formula}+${ref}` };
  } else if (typeof v === "number" && v !== 0) {
    celda.value = { formula: `${v}+${ref}` };
  } else {
    celda.value = { formula: `+${ref}` };
  }
  log(`  Anexo II ${colLetra}${filaConcepto}: sumada la referencia a ${ref}.`);
}

// ACTIVO/PASIVO: se inserta una línea de detalle en la Nota 4 (hoja 'Activo y
// Pasivo') debajo de la última línea del rubro elegido, copiando su patrón: la
// etiqueta va en la columna de texto y la fórmula ±SALDOS!G<fila> en la de importes.
// El signo se hereda de la línea modelo (el pasivo va con el signo invertido).
function agregarLineaNota4(wb, moneda, lineaModelo, cuenta, filaSaldos, colSaldos, log = () => {}) {
  const ws = wb.getWorksheet("Activo y Pasivo");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Activo y Pasivo'.");

  const filaModelo = lineaModelo.fila;
  const filaNueva = filaModelo + 1;
  const mod = insertRowEn(wb, "Activo y Pasivo", filaNueva);

  const colF = lineaModelo.colFormula;
  const modelo = ws.getCell(filaModelo, colF).value;
  const signo = (modelo && typeof modelo === "object" && /^\s*-/.test(modelo.formula || "")) ? "-" : "+";
  const ref = `SALDOS!${String.fromCharCode(64 + colSaldos)}${filaSaldos}`;

  ws.getCell(filaNueva, lineaModelo.colTexto).value = `- ${cuenta.nombre}`;
  ws.getCell(filaNueva, lineaModelo.colTexto).style = ws.getCell(filaModelo, lineaModelo.colTexto).style;
  ws.getCell(filaNueva, colF).value = { formula: `${signo}${ref}` };
  ws.getCell(filaNueva, colF).style = ws.getCell(filaModelo, colF).style;

  log(`  Nota 4: línea "- ${cuenta.nombre}" insertada en la fila ${filaNueva} (${signo}${ref}, ${mod} referencias reacomodadas).`);
  return filaNueva;
}

// Las líneas de detalle de un rubro de la Nota 4, leídas del archivo: texto en la
// columna B/C y fórmula que referencia SALDOS en la misma fila.
function lineasDeNota4(wb) {
  const ws = wb.getWorksheet("Activo y Pasivo");
  if (!ws) return [];
  const lineas = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 2; c <= 12; c++) {
      const v = ws.getCell(r, c).value;
      if (v && typeof v === "object" && typeof v.formula === "string" && /SALDOS!/i.test(v.formula)) {
        // el texto de la línea está a la izquierda de la fórmula
        let colTexto = null, texto = "";
        for (let ct = c - 1; ct >= 1; ct--) {
          const t = textoPlano(ws.getCell(r, ct).value).trim();
          if (t) { colTexto = ct; texto = t; break; }
        }
        lineas.push({ fila: r, colFormula: c, colTexto: colTexto || c - 1, texto });
        break;
      }
    }
  }
  return lineas;
}

// Los conceptos del Anexo II, leídos del archivo: texto en B y un total
// SUM(F<r>:I<r>) en E. Las cuatro columnas de destino son F/G/H/I.
function conceptosAnexo2(wb) {
  const ws = wb.getWorksheet("Anexo II");
  if (!ws) return [];
  const conceptos = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const texto = textoPlano(ws.getCell(r, 2).value).trim();
    if (!texto) continue;
    const e = ws.getCell(r, 5).value;
    if (e && typeof e === "object" && typeof e.formula === "string" &&
        new RegExp(`SUM\\(F${r}:I${r}\\)`, "i").test(e.formula)) {
      conceptos.push({ fila: r, concepto: texto });
    }
  }
  return conceptos;
}

const ANEXO2_COLUMNAS = [
  { col: "F", nombre: "Administración" },
  { col: "G", nombre: "Comercialización" },
  { col: "H", nombre: "Exploración" },
  { col: "I", nombre: "Financieros" },
];

// ---------------------------------------------------------------- la corrida

function procesarBalance({ wb, cuentasExport, moneda, destinosElegidos = {}, log = () => {} }) {
  const mapeo = derivarMapeoMaestro(wb, moneda);
  if (mapeo.duplicadas.length) {
    log(`  ⚠ SALDOS tiene cuentas repetidas: ` +
        mapeo.duplicadas.map(d => `${d.codigo} (filas ${d.filaPrevia} y ${d.fila})`).join(", ") +
        `. Se usa la primera; conviene limpiar el archivo a mano.`);
  }

  const nuevas = detectarNuevas(cuentasExport, mapeo);
  // Los proveedores nuevos (regla confirmada por la usuaria) no preguntan destino:
  // en pesos cada uno inserta su fila en el detalle del pasivo, y el subtotal que
  // es un rango los absorbe al expandirse. El resto sí necesita destino elegido.
  const esProveedor = (n) => /^2110/.test(n.codigo);
  const sinDestino = nuevas.filter(n => !esProveedor(n) && !destinosElegidos[n.codigo]);
  if (sinDestino.length) {
    const e = new Error("Hay cuentas nuevas sin destino elegido.");
    e.nuevasSinDestino = sinDestino;
    throw e;
  }

  // primero las inserciones (mueven filas de SALDOS), después la reescritura de Hoja1
  for (const n of nuevas) {
    const destino = destinosElegidos[n.codigo];
    const { filaNueva, colValor } = insertarCuentaEnSaldos(wb, mapeo, n, moneda, log);
    if (esProveedor(n)) continue;
    if (n.capitulo === "RESULTADOS") {
      agregarRefAnexo2(wb, moneda, destino.filaConcepto, destino.columna, filaNueva, colValor, log);
    } else {
      agregarLineaNota4(wb, moneda, destino.lineaModelo, n, filaNueva, colValor, log);
    }
  }

  const ultimaFila = reescribirHoja1(wb, cuentasExport, moneda, log);
  normalizarRangosVlookup(wb, moneda, ultimaFila, log);
  actualizarSaldosManuales(wb, mapeo, cuentasExport, moneda, log);

  // control: qué cuentas del export no quedaron enganchadas a ninguna fila de SALDOS
  const p = PARAMS[moneda];
  const noEnganchadas = cuentasExport.filter(c => !mapeo.cuentas[c.codigo] && c[p.campoSaldo] !== 0);

  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  const total = cuentasExport.reduce((s, c) => s + c[p.campoSaldo], 0);
  log(`\nResumen ${moneda}: ${cuentasExport.length} cuentas del export, ${nuevas.length} nuevas insertadas.`);

  return {
    resumen: {
      moneda,
      cuentas: cuentasExport.length,
      nuevas: nuevas.length,
      total,
      duplicadas: mapeo.duplicadas,
      noEnganchadas: noEnganchadas.map(c => `${c.codigo} - ${c.nombre}`),
    },
  };
}

if (typeof module !== "undefined") {
  module.exports = {
    PARAMS, derivarMapeoMaestro, reescribirHoja1, normalizarRangosVlookup,
    detectarNuevas, insertarCuentaEnSaldos, agregarRefAnexo2, agregarLineaNota4,
    lineasDeNota4, procesarBalance, copiarPatronFila, actualizarSaldosManuales,
    conceptosAnexo2, ANEXO2_COLUMNAS,
  };
}
