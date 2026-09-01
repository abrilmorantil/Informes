// Configuración de los balances formales — una por moneda.
//
// Por qué existe: hoy la decisión de "qué cuenta de Onvio alimenta cada línea del balance" no
// está declarada en ninguna parte. Vive escrita a mano adentro del Excel, en el texto de las
// columnas de SALDOS, y el motor la deduce leyendo ese texto. Eso ya se desincronizó: en el
// maestro de pesos la fila 137 dice "213010010 - Patentes a pagar" en una columna y
// "212020002 - Retenciones SUSS a pagar" en la otra —dos cuentas distintas, las dos reales—
// y nadie se enteró hasta que se buscó a mano.
//
// Son DOS configuraciones y no una. Medido sobre los maestros reales: pesos nombra 522
// códigos y dólares 200, con sólo 163 en común, porque el de dólares está armado a nivel de
// cuenta madre y el de pesos tiene el detalle. Y los dos guardan la equivalencia de otra
// forma:
//
//   pesos    una columna trae la cuenta de Onvio y otra el código viejo del cliente
//            (`212020001 SICOSS a pagar` / `21407000 Sicoss a pagar`)
//   dólares  hay UNA sola columna, y lo que trae es el código viejo; la equivalencia con
//            Onvio vive aparte, en equivalencias_dolares.json — y ahí está guardada al revés
//            que en los otros lados: la clave es el código de Onvio y el valor el viejo.
//
// La configuración NO reemplaza al Excel: el maestro sigue teniendo la geometría (qué fila,
// qué línea de la Nota 4, qué subtotales). Lo que declara son las decisiones que hoy están
// implícitas en el texto, para poder verlas y corregirlas sin abrir el Excel.

const CFG_RE_CUENTA = /^\s*(\d{5,})\s*-\s*(.+?)\s*$/;
const CFG_RE_REF_SALDOS = /SALDOS!\$?([A-Z]{1,3})\$?(\d+)/i;

function cfgTexto(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return "";
    return v.result === undefined || v.result === null ? "" : String(v.result);
  }
  return String(v);
}

function cfgLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ¿La fila es un subtotal de familia? En RESULTADOS la cuenta madre y su primera subcuenta
// comparten fila: la madre va en D con el subtotal en G, la subcuenta en E con el VLOOKUP en
// F. Sin esta distinción, esas filas parecen "dos cuentas distintas en la misma fila" —da 3
// falsas alarmas en el maestro de pesos (filas 200, 413 y 416), y taparían a la única de
// verdad, que es la 137.
function cfgEsSubtotal(ws, fila, params, filasQueAgrega) {
  const colValor = (params && params.saldosColValor) || 7;
  for (const c of [colValor, colValor - 1]) {
    const v = ws.getCell(fila, c).value;
    if (!(v && typeof v === "object" && typeof v.formula === "string")) continue;
    if (typeof filasQueAgrega === "function" && filasQueAgrega(v.formula, fila).length) return true;
  }
  return false;
}

function cfgOtrasColumnas(ws, fila, colClave, colsPosibles) {
  const otras = [];
  for (const c of colsPosibles) {
    if (c === colClave) continue;
    const m = CFG_RE_CUENTA.exec(cfgTexto(ws.getCell(fila, c).value).trim());
    if (m) otras.push({ columna: c, code: m[1], description: m[2] });
  }
  return otras;
}

// Deriva la configuración leyendo el maestro. `mapeo` es lo que devuelve
// derivarMapeoMaestro(): el mismo índice que usa el motor, así la configuración describe lo
// que la app hace de verdad y no una segunda interpretación del archivo.
//
// `equivalencias` es equivalencias_dolares.json tal cual está guardado (Onvio -> viejo); sólo
// se usa en dólares, que es donde el maestro no trae la cuenta de Onvio.
function derivarConfigBalance(wb, moneda, mapeo, lineasNota4, params, opciones) {
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El maestro no tiene la hoja 'SALDOS'.");
  const o = opciones || {};
  const colsCuenta = (params && params.saldosColsCuenta) || [3, 4, 5];

  // Qué línea de la Nota 4 lee cada fila de SALDOS
  const notaPorFila = {};
  const hojaNota = wb.getWorksheet("Activo y Pasivo");
  for (const l of (lineasNota4 || [])) {
    if (!hojaNota) break;
    const cel = hojaNota.getCell(l.fila, l.colFormula);
    const m = cel.formula && CFG_RE_REF_SALDOS.exec(String(cel.formula));
    if (!m) continue;
    (notaPorFila[+m[2]] = notaPorFila[+m[2]] || []).push({ fila: l.fila, texto: l.texto });
  }

  // En dólares, qué cuentas de Onvio alimentan cada línea. Se da vuelta el archivo.
  const onvioPorLinea = {};
  for (const [deOnvio, v] of Object.entries(o.equivalencias || {})) {
    const destino = typeof v === "string" ? v : (v && (v.codigo || v.code));
    if (!destino || !/^\d{5,}$/.test(String(destino))) continue;
    (onvioPorLinea[String(destino)] = onvioPorLinea[String(destino)] || []).push(deOnvio);
  }

  const lineas = [];
  const avisos = [];
  for (const [codigo, f] of Object.entries(mapeo.cuentas)) {
    const otras = cfgOtrasColumnas(ws, f.fila, f.col, colsCuenta);
    const esSubtotal = cfgEsSubtotal(ws, f.fila, params, o.filasQueAgrega);

    let cliente = null, madre = null;
    for (const x of otras) {
      if (x.code === codigo) continue;
      if (x.code.length < codigo.length) {
        // el plan viejo tiene un dígito menos en el grupo del medio
        cliente = { code: x.code, description: x.description };
      } else if (esSubtotal) {
        // madre y subcuenta compartiendo fila: es el armado normal de RESULTADOS
        madre = { code: x.code, description: x.description };
      } else {
        avisos.push({
          tipo: "dos_cuentas_en_la_misma_fila", fila: f.fila,
          usa: { code: codigo, description: f.nombre },
          tambien: { code: x.code, description: x.description, columna: cfgLetra(x.columna) },
        });
      }
    }

    // De dónde sale el saldo de esta línea
    let onvio;
    if (moneda === "dolares") {
      onvio = onvioPorLinea[codigo] || (o.equivalencias ? [] : [codigo]);
      // si el propio código es una cuenta de Onvio y nadie la mapea, se alimenta sola
      if (!onvio.length && codigo.length >= 9) onvio = [codigo];
      if (!cliente && codigo.length < 9) cliente = { code: codigo, description: f.nombre };
    } else {
      onvio = [codigo];
    }

    lineas.push({
      linea: codigo,
      nombre: f.nombre,
      capitulo: f.capitulo || null,
      fila: f.fila,
      columna: cfgLetra(f.col),
      subtotal: esSubtotal || undefined,
      onvio,
      cliente,
      madre,
      nota4: (notaPorFila[f.fila] || []).map(x => ({ fila: x.fila, texto: x.texto })),
    });
  }

  lineas.sort((a, b) => a.fila - b.fila);

  // Líneas de la Nota 4 que no las alimenta ninguna cuenta: se imprimen siempre en cero.
  //
  // Hay que distinguir dos cosas que se parecen y no son lo mismo:
  //
  //   - la línea apunta a un SUBTOTAL (el bloque de proveedores es `SALDOS!I116 =
  //     SUM(G63:G114)`): está perfecto, la alimentan las 52 filas de abajo. Avisar por eso es
  //     ruido, y el ruido enseña a ignorar el panel.
  //   - la fila SÍ tiene una cuenta escrita, pero en OTRA columna que la que lee el VLOOKUP.
  //     Eso es un problema de verdad y no se ve mirando la hoja: en el maestro de pesos pasa
  //     en las filas 39, 40 y 43 —"- Socios", "- Zento S.A." y "- Títulos Públicos"—, donde la
  //     cuenta quedó en la D y la fórmula lee la C, que está vacía. Esas líneas no pueden
  //     traer un importe nunca. Hoy las dos cuentas están en cero, así que no se ve.
  const filasConCuenta = new Set(lineas.map(l => l.fila));
  const colDe = (letras) => letras.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  for (const l of (lineasNota4 || [])) {
    if (!hojaNota) break;
    const cel = hojaNota.getCell(l.fila, l.colFormula);
    const m = cel.formula && CFG_RE_REF_SALDOS.exec(String(cel.formula));
    if (!m) continue;
    const filaSaldos = +m[2];
    if (filasConCuenta.has(filaSaldos)) continue;

    // ¿la celda a la que apunta es un subtotal? entonces está bien
    const apuntada = ws.getCell(filaSaldos, colDe(m[1].toUpperCase()));
    if (apuntada.formula && typeof o.filasQueAgrega === "function" &&
        o.filasQueAgrega(apuntada.formula, filaSaldos).length) continue;

    // ¿la fila tiene la cuenta escrita en otra columna?
    const enOtraColumna = [];
    for (let c = 1; c <= ws.columnCount && c <= 10; c++) {
      const mm = CFG_RE_CUENTA.exec(cfgTexto(ws.getCell(filaSaldos, c).value).trim());
      if (mm) enOtraColumna.push({ columna: cfgLetra(c), code: mm[1], description: mm[2] });
    }
    avisos.push({
      tipo: enOtraColumna.length ? "cuenta_en_la_columna_equivocada" : "linea_sin_cuenta",
      nota4: l.fila, texto: l.texto, saldos: filaSaldos,
      cuentas: enOtraColumna,
    });
  }

  for (const d of (mapeo.duplicadas || [])) {
    avisos.push({ tipo: "codigo_repetido", code: d.codigo, fila: d.fila, filaPrevia: d.filaPrevia });
  }

  const sinOnvio = lineas.filter(l => !l.onvio.length).length;
  return {
    moneda,
    resumen: {
      lineas: lineas.length,
      conCliente: lineas.filter(l => l.cliente).length,
      conMadre: lineas.filter(l => l.madre).length,
      enNota4: lineas.filter(l => l.nota4.length).length,
      sinCuentaDeOnvio: sinOnvio,
    },
    lineas,
    avisos,
  };
}

if (typeof module !== "undefined") {
  module.exports = { derivarConfigBalance, cfgTexto, cfgLetra, cfgEsSubtotal };
}
