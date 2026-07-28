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

// Las filas de SALDOS que una fórmula de saldo agrega, sin contar la propia. Vacío
// significa "esta fórmula habla solo de su fila": una cuenta, no una cuenta madre.
function filasQueAgrega(formula, filaPropia) {
  const filas = new Set();
  // fuera las referencias a otras hojas: no son filas de SALDOS
  let f = formula.replace(/(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)![^\s,;)+\-*/]*/g, " ");
  // primero los rangos (y se los saca, para que no los relea el paso siguiente)
  f = f.replace(/\b[A-Z]{1,3}(\d+)\s*:\s*[A-Z]{1,3}(\d+)/g, (_, a, b) => {
    const desde = Math.min(+a, +b), hasta = Math.max(+a, +b);
    for (let r = desde; r <= hasta; r++) filas.add(r);
    return " ";
  });
  f.replace(/\b[A-Z]{1,3}(\d+)\b/g, (_, n) => { filas.add(+n); return " "; });
  filas.delete(filaPropia);
  return [...filas];
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

  // Qué columna es LA cuenta de la fila no lo decide el texto sino la fórmula: en
  // RESULTADOS la columna D lleva el título de la familia (la "cuenta madre", con el
  // subtotal en G) y la cuenta de verdad vive en E — a veces las dos en la misma
  // fila. Tomar la primera columna con texto agarraba a la madre, tapaba a la hija y
  // encima la hija aparecía después como "cuenta nueva" repetida. Por eso:
  //   1. si la fila tiene VLOOKUP contra Hoja1, la clave es la columna que usa de
  //      argumento (E en resultados, C en el activo);
  //   2. si tiene otra fórmula de saldo (SUM / +F...) es una fila de subtotal de la
  //      madre: no es una cuenta, se saltea;
  //   3. si no tiene ninguna fórmula es una fila manual (el detalle de proveedores):
  //      vale la primera columna con texto de cuenta, como siempre.
  const RE_VLOOKUP_CLAVE = /VLOOKUP\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*[,;]\s*Hoja1!/i;
  const colAIdx = (s) => s.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);

  for (let r = 1; r <= ws.rowCount; r++) {
    let colClave = null;
    let agregaOtrasFilas = false;
    let tieneFormulaSaldo = false;
    for (const c of [p.saldosColValor, p.saldosColValor - 1]) {
      const v = ws.getCell(r, c).value;
      if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
      tieneFormulaSaldo = true;
      const m = RE_VLOOKUP_CLAVE.exec(v.formula);
      if (m && parseInt(m[2], 10) === r) { colClave = colAIdx(m[1]); break; }
      if (filasQueAgrega(v.formula, r).length) agregaOtrasFilas = true;
    }

    let candidatas;
    if (colClave !== null) candidatas = [colClave];
    // Una fila es cuenta madre solo si su saldo AGREGA otras filas (SUM(F10:F20),
    // +F63+F64). Si la fórmula habla únicamente de su propia fila es una cuenta de
    // verdad con fórmula a medida: el caso de "423060000 - Diferencia de cambio $"
    // (F282 = +Hoja1!E206+Hoja1!E207, G282 = +F282), que si se saltea aparece
    // después como cuenta nueva y se inserta duplicada.
    else if (agregaOtrasFilas) continue;
    else candidatas = p.saldosColsCuenta;         // fila manual o fórmula a medida

    for (const c of candidatas) {
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
      break;
    }
  }

  return { moneda, cuentas, duplicadas };
}

// --------------------------------------------------------------------- Hoja1
//
// Hoja1 NO se reordena: se actualizan los importes en la fila donde ya está cada
// cuenta, y las que no estaban se insertan al final de su capítulo.
//
// La primera versión reescribía la hoja entera de cero, y eso rompía en silencio dos
// cosas del maestro real:
//   - Fórmulas que la leen POR POSICIÓN. `SALDOS!F282 = +Hoja1!E206+Hoja1!E207` es la
//     "Diferencia de cambio $", que alimenta `Resultados!C17` (la propia celda H282
//     aclara "No va en Anexo II"); y `SALDOS!E426 = +Hoja1!A267` toma de ahí el NOMBRE
//     de una cuenta. Al reordenar quedaban apuntando a cualquier cosa.
//   - Controles que viven DENTRO de la zona de pegado: `E52 = SUM(E2:E51)` (total del
//     activo, que lee `SALDOS!I62`) y `E318 = SUM(E192:E316)` (total de resultados, que
//     lee `SALDOS!G508` y alimenta la comprobación de diferencia de cambio). Borrar la
//     columna los eliminaba.
//
// Actualizando en el lugar no hace falta reacomodar nada: las filas no se mueven, y las
// pocas inserciones pasan por insertRowEn, que ya corrige todas las fórmulas del archivo
// —incluidos esos subtotales y los rangos de los VLOOKUP—, igual que haría Excel.
//
// El texto de la clave NO se toca: SALDOS busca contra Hoja1 con su propio texto, así que
// si Onvio le cambia el nombre a una cuenta y lo pisáramos acá, el VLOOKUP dejaría de
// encontrarla. Se empareja por CÓDIGO, que es lo estable.
function actualizarHoja1(wb, cuentasExport, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("Hoja1");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Hoja1' (la zona de pegado del export).");

  const indice = () => {
    const m = new Map();
    for (let r = 1; r <= ws.rowCount; r++) {
      const v = ws.getCell(r, p.hoja1ColClave).value;
      if (v && typeof v === "object" && typeof v.formula === "string") continue;  // celda de control
      const t = String(v === null || v === undefined ? "" : v).trim();
      const mm = RE_CUENTA_TXT.exec(t);
      if (!mm) continue;
      if (!m.has(mm[1])) m.set(mm[1], []);
      m.get(mm[1]).push(r);
    }
    return m;
  };

  let filasPorCodigo = indice();
  const repetidas = [...filasPorCodigo.entries()].filter(([, f]) => f.length > 1);

  // cada corrida REEMPLAZA: lo que no venga en el export queda en cero, no arrastra
  for (const filas of filasPorCodigo.values()) {
    for (const r of filas) ws.getCell(r, p.hoja1ColValor).value = 0;
  }

  let actualizadas = 0;
  const pendientes = [];
  for (const c of cuentasExport) {
    const filas = filasPorCodigo.get(c.codigo);
    if (!filas) { pendientes.push(c); continue; }
    ws.getCell(filas[0], p.hoja1ColValor).value = c[p.campoSaldo];
    actualizadas++;
  }

  for (const c of pendientes) {
    // al final del bloque de su capítulo, que en Hoja1 se reconoce por el primer dígito
    let ultima = null;
    for (const [codigo, filas] of filasPorCodigo) {
      if (codigo[0] !== c.codigo[0]) continue;
      const f = Math.max(...filas);
      if (ultima === null || f > ultima) ultima = f;
    }
    if (ultima === null) {
      ultima = 0;
      for (const filas of filasPorCodigo.values()) ultima = Math.max(ultima, ...filas);
    }
    // se inserta EN la fila de la última cuenta del capítulo, no después: así queda
    // estrictamente dentro de los rangos que terminan ahí (los controles internos
    // E52=SUM(E2:E51) y E318=SUM(E192:E316)) y se expanden solos, como en Excel.
    // Insertando después, la cuenta caía justo afuera del control.
    const filaNueva = ultima;
    insertRowEn(wb, "Hoja1", filaNueva);
    ws.getCell(filaNueva, p.hoja1ColClave).value = `${c.codigo} - ${c.nombre}`;
    ws.getCell(filaNueva, p.hoja1ColValor).value = c[p.campoSaldo];
    ws.getCell(filaNueva, p.hoja1ColClave).style = ws.getCell(ultima + 1, p.hoja1ColClave).style;
    ws.getCell(filaNueva, p.hoja1ColValor).style = ws.getCell(ultima + 1, p.hoja1ColValor).style;
    filasPorCodigo = indice();
  }

  let ultimaFila = 0;
  for (const filas of filasPorCodigo.values()) ultimaFila = Math.max(ultimaFila, ...filas);

  log(`  Hoja1 actualizada: ${actualizadas} cuentas en su fila` +
      (pendientes.length ? `, ${pendientes.length} agregada(s) al final de su capítulo` : "") +
      ` (los importes que no vinieron en el export quedaron en cero).`);
  for (const [codigo, filas] of repetidas) {
    log(`  ⚠ Hoja1 tiene ${codigo} repetida en las filas ${filas.join(", ")}: se carga en la primera y las otras quedan en cero.`);
  }
  return ultimaFila;
}

// Se normaliza el rango entero y la columna que se pide:
//   - el ARRANQUE a la fila 2: hay VLOOKUPs que empiezan en $A$86 (donde arrancaban los
//     gastos en el layout viejo) y si la sección queda más arriba tras reescribir Hoja1,
//     esas cuentas se perderían. Las claves son únicas, agrandar nunca cambia el resultado.
//   - el FIN a la última fila con datos.
//   - la COLUMNA pedida a la del saldo. En el maestro real hay cuatro fórmulas rotas
//     (`Hoja1!$A$86:$C$353` pidiendo la columna 5, o la 3, que en Hoja1 está vacía): son
//     un #REF!/vacío que el IFERROR convierte en cero, así que esas cuentas nunca
//     levantan su saldo y nadie se entera. Hoy dan cero igual porque no vienen en el
//     export, pero el día que aparezcan tienen que funcionar.
function normalizarRangosVlookup(wb, moneda, ultimaFilaHoja1, log = () => {}) {
  const p = PARAMS[moneda];
  const colValor = String.fromCharCode(64 + p.hoja1ColValor);   // E en pesos, D en dólares
  const idxValor = p.hoja1ColValor;                             // el rango arranca en A
  const RE = /Hoja1!\$A\$(\d+):\$([A-Z]{1,3})\$(\d+)\s*[,;]\s*(\d+)/gi;
  const ws = wb.getWorksheet("SALDOS");
  let extendidas = 0;
  const reparadas = [];

  ws.eachRow(row => row.eachCell(cell => {
    const v = cell.value;
    if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
    let cambio = false, rota = false;
    const nueva = v.formula.replace(RE, (m, desde, col, fin, idx) => {
      const d = Math.min(parseInt(desde, 10), 2);
      const f = Math.max(parseInt(fin, 10), ultimaFilaHoja1);
      if (col.toUpperCase() !== colValor || parseInt(idx, 10) !== idxValor) rota = true;
      if (d === parseInt(desde, 10) && f === parseInt(fin, 10) &&
          col.toUpperCase() === colValor && parseInt(idx, 10) === idxValor) return m;
      cambio = true;
      return `Hoja1!$A$${d}:$${colValor}$${f},${idxValor}`;
    });
    if (cambio) {
      cell.value = { formula: nueva };
      if (rota) reparadas.push(cell.address); else extendidas++;
    }
  }));

  if (extendidas) log(`  ${extendidas} VLOOKUP de SALDOS tenían el rango corto y se extendieron hasta la fila ${ultimaFilaHoja1}.`);
  if (reparadas.length) {
    log(`  ⚠ ${reparadas.length} VLOOKUP del archivo apuntaban a una columna equivocada y devolvían siempre cero ` +
        `(${reparadas.join(", ")}): corregidos para que lean el saldo.`);
  }
  return extendidas + reparadas.length;
}

// ------------------------------------------------------------- cuentas nuevas

// Una cuenta solo es "nueva" para ESTE balance si además tiene saldo en ESTA moneda.
// El export trae las dos juntas y muchas cuentas se mueven solo en dólares (50 de 163
// en junio 2026): crearles fila en el balance de pesos agrega una línea que siempre
// va a valer cero. Si el mes que viene tienen saldo en pesos, se preguntan ahí.
// Igual entran todas a Hoja1: hay fórmulas que las leen de ahí aunque no tengan fila
// propia en SALDOS (ej. las dos de diferencia de cambio, que SALDOS!F282 lee de ahí).
function detectarNuevas(cuentasExport, mapeoMaestro, moneda) {
  const campo = PARAMS[moneda].campoSaldo;
  return cuentasExport.filter(c => !mapeoMaestro.cuentas[c.codigo] && c[campo] !== 0);
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

// Las cuentas madre de RESULTADOS: los estados NO leen las subcuentas — el Anexo II
// referencia el subtotal G de la madre (+SALDOS!G324, etc.), así que una subcuenta
// nueva solo tiene que caer adentro del bloque de su madre y el subtotal la absorbe.
// Nada más: agregarle además una referencia en el Anexo II la contaría DOS veces.
//
// Una madre es una fila con texto de cuenta cuyo G no es un VLOOKUP sino la suma de
// su familia: SUM(F<a>:F<b>) o una lista +F<x>+F<y> (a veces vacía, +F de su propia
// fila). El texto del título vive en D; da igual, se toma de donde esté.
function madresResultados(wb, moneda) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) return [];
  const madres = [];
  const RE_RANGO = /^\s*\+?SUM\(F(\d+):F(\d+)\)\s*$/i;
  const RE_LISTA = /^\s*(?:\+F\d+)+\s*$/i;

  for (let r = 1; r <= ws.rowCount; r++) {
    const v = ws.getCell(r, p.saldosColValor).value;
    if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
    const f = v.formula;
    if (/VLOOKUP/i.test(f)) continue;
    let bloque = null;
    const mR = RE_RANGO.exec(f);
    if (mR) bloque = { tipo: "rango", desde: parseInt(mR[1], 10), hasta: parseInt(mR[2], 10) };
    else if (RE_LISTA.test(f)) bloque = { tipo: "lista", filas: [...f.matchAll(/\+F(\d+)/gi)].map(x => parseInt(x[1], 10)) };
    if (!bloque) continue;
    // una madre AGREGA otras filas; si la fórmula habla solo de la suya es una cuenta
    // común con fórmula a medida (ej. "Diferencia de cambio $", G282 = +F282)
    if (!filasQueAgrega(f, r).length) continue;

    let texto = null;
    for (const c of p.saldosColsCuenta) {
      const t = textoPlano(ws.getCell(r, c).value).trim();
      if (RE_CUENTA_TXT.test(t)) { texto = t; break; }
    }
    if (!texto) continue;
    const m = RE_CUENTA_TXT.exec(texto);
    madres.push({ fila: r, codigo: m[1], nombre: m[2], bloque });
  }
  return madres;
}

// Inserta una subcuenta nueva de RESULTADOS dentro del bloque de su cuenta madre.
// Si el bloque es un rango SUM(Fa:Fb), la fila entra antes de la última para que el
// rango se expanda solo con el shift; si es una lista +F..+F.., entra al final y se
// le agrega su término al subtotal de la madre.
function insertarHijaEnMadre(wb, mapeoMaestro, cuenta, madreFila, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  const madre = madresResultados(wb, moneda).find(x => x.fila === madreFila);
  if (!madre) throw new Error(`No encontré la cuenta madre en la fila ${madreFila} de SALDOS. NO se generó ningún archivo.`);

  // la fila modelo es una subcuenta con VLOOKUP: del propio bloque si tiene, si no
  // la subcuenta de RESULTADOS más cercana (bloques vacíos)
  const filasBloque = madre.bloque.tipo === "rango"
    ? Array.from({ length: madre.bloque.hasta - madre.bloque.desde + 1 }, (_, i) => madre.bloque.desde + i)
    : madre.bloque.filas;
  const esHija = (r) => {
    const v = ws.getCell(r, p.saldosColValor - 1).value;
    return v && typeof v === "object" && /VLOOKUP/i.test(v.formula || "");
  };
  let filaModelo = filasBloque.filter(r => r !== madre.fila && esHija(r)).pop() || null;
  if (filaModelo === null) {
    const hijas = Object.values(mapeoMaestro.cuentas)
      .filter(x => x.capitulo === "RESULTADOS" && esHija(x.fila))
      .sort((a, b) => Math.abs(a.fila - madre.fila) - Math.abs(b.fila - madre.fila));
    if (!hijas.length) throw new Error(`No hay ninguna subcuenta con fórmula en RESULTADOS para usar de modelo. NO se generó ningún archivo.`);
    filaModelo = hijas[0].fila;
  }

  const filaNueva = madre.bloque.tipo === "rango"
    ? Math.max(madre.bloque.desde, madre.bloque.hasta)      // antes de la última: el rango se expande
    : Math.max(...madre.bloque.filas, madre.fila) + 1;      // al final de la lista

  const mod = insertRowEn(wb, "SALDOS", filaNueva);
  for (const info of Object.values(mapeoMaestro.cuentas)) {
    if (info.fila >= filaNueva) info.fila += 1;
  }
  const filaModeloCorrida = filaModelo >= filaNueva ? filaModelo + 1 : filaModelo;

  // clave en la columna que usa el VLOOKUP del modelo, fórmula en F
  const vModelo = ws.getCell(filaModeloCorrida, p.saldosColValor - 1).value;
  const mV = /VLOOKUP\(\s*\$?([A-Z]{1,3})\$?\d+/i.exec(vModelo.formula);
  const colClave = mV[1].split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  const clave = `${cuenta.codigo} - ${cuenta.nombre}`;
  ws.getCell(filaNueva, colClave).value = clave;
  ws.getCell(filaNueva, colClave).style = ws.getCell(filaModeloCorrida, colClave).style;
  copiarPatronFila(ws, filaModeloCorrida, filaNueva, [p.saldosColValor - 1]);

  // en los bloques de lista, el subtotal de la madre suma el término nuevo a mano
  if (madre.bloque.tipo === "lista") {
    const madreFilaCorrida = madre.fila >= filaNueva ? madre.fila + 1 : madre.fila;
    const g = ws.getCell(madreFilaCorrida, p.saldosColValor).value;
    ws.getCell(madreFilaCorrida, p.saldosColValor).value = { formula: `${g.formula}+F${filaNueva}` };
  }

  mapeoMaestro.cuentas[cuenta.codigo] = {
    fila: filaNueva, col: colClave, clave, nombre: cuenta.nombre, capitulo: "RESULTADOS",
  };
  log(`  ${clave} insertada como subcuenta de "${madre.nombre}" (fila ${filaNueva}, ${mod} referencias reacomodadas).`);
  return filaNueva;
}

// ---------------------------------------------------------------- la corrida

function procesarBalance({ wb, cuentasExport, moneda, destinosElegidos = {}, log = () => {} }) {
  const mapeo = derivarMapeoMaestro(wb, moneda);
  if (mapeo.duplicadas.length) {
    log(`  ⚠ SALDOS tiene cuentas repetidas: ` +
        mapeo.duplicadas.map(d => `${d.codigo} (filas ${d.filaPrevia} y ${d.fila})`).join(", ") +
        `. Se usa la primera; conviene limpiar el archivo a mano.`);
  }

  const nuevas = detectarNuevas(cuentasExport, mapeo, moneda);
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

  // primero las inserciones (mueven filas de SALDOS), después la reescritura de Hoja1.
  // Cada inserción corre hacia abajo todo lo que está después, así que las filas de
  // madre elegidas en el formulario (que se eligieron sobre el archivo sin tocar) se
  // van corrigiendo a medida que se inserta — igual que el mapeo.
  const ajustarDestinos = (filaInsertada) => {
    for (const d of Object.values(destinosElegidos)) {
      if (d && typeof d.madreFila === "number" && d.madreFila >= filaInsertada) d.madreFila += 1;
    }
  };
  for (const n of nuevas) {
    const destino = destinosElegidos[n.codigo];
    if (n.capitulo === "RESULTADOS" && !esProveedor(n)) {
      // adentro del bloque de su madre: el Anexo II lee el subtotal de la madre,
      // así que no se toca ningún estado (tocarlo contaría el importe dos veces)
      const fila = insertarHijaEnMadre(wb, mapeo, n, destino.madreFila, moneda, log);
      ajustarDestinos(fila);
      continue;
    }
    const { filaNueva, colValor } = insertarCuentaEnSaldos(wb, mapeo, n, moneda, log);
    ajustarDestinos(filaNueva);
    if (esProveedor(n)) continue;
    agregarLineaNota4(wb, moneda, destino.lineaModelo, n, filaNueva, colValor, log);
  }

  const ultimaFila = actualizarHoja1(wb, cuentasExport, moneda, log);
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
    PARAMS, derivarMapeoMaestro, actualizarHoja1, normalizarRangosVlookup,
    detectarNuevas, insertarCuentaEnSaldos, agregarLineaNota4,
    lineasDeNota4, procesarBalance, copiarPatronFila, actualizarSaldosManuales,
    madresResultados, insertarHijaEnMadre, filasQueAgrega,
  };
}
