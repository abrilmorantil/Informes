// El Anexo II: en qué columna cae el gasto de cada cuenta madre.
//
// Es la decisión contable que hoy vive escrita a mano adentro del Excel y no se ve desde
// ningún lado. El Anexo II tiene un concepto por fila (ALOJAMIENTO, ALQUILERES, CARGAS
// SOCIALES...) y cuatro columnas —Administración, Comercialización, Exploración, Financieros—,
// y cada celda lee el SUBTOTAL de una cuenta madre de SALDOS: `+SALDOS!G322`. Nunca lee las
// subcuentas: por eso una subcuenta nueva se inserta DENTRO del bloque de su madre y el Anexo
// II no se toca — agregarle una referencia contaría el gasto dos veces.
//
// Un mismo concepto puede partirse en dos columnas según la cuenta. En el maestro real:
//
//     CARGAS SOCIALES   Administración   <- 421060000 Cargas sociales Adm.
//                       Exploración      <- 424500000 CARGAS SOCIALES CAMPO
//
// El invariante que hace que esto sea seguro: **cada madre entra exactamente una vez**. Dos
// veces cuenta el gasto doble; cero lo hace desaparecer del estado de resultados sin que nada
// lo diga. Medido sobre el maestro de pesos: 58 madres, las 58 una sola vez (33 en
// Administración, 25 en Exploración).
//
// Nombres con prefijo `a2` por el ámbito global único que comparten los scripts del sitio.

const A2_HOJA = "Anexo II";
const A2_FILA_TITULOS = 8;          // donde dicen ADMINISTRACION / COMERCIALIZACION / ...
const A2_COL_CONCEPTO = 2;          // la columna B
// Se arma en cada uso y no una sola vez al cargar el archivo: así toma el nombre de hoja
// que corresponda, y de paso no arrastra el `lastIndex` de la búsqueda anterior.
const a2ReRef = () => new RegExp(`\\+?\\s*${REF_DISTRIB}!\\$?([A-Z]{1,3})\\$?(\\d+)`, "gi");

function a2Texto(cell) {
  const v = cell && cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return "";
    return v.result === undefined || v.result === null ? "" : String(v.result);
  }
  return String(v);
}

function a2Letra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Las columnas de gasto, leídas del propio archivo. No se hardcodean: si el maestro cambiara,
// la configuración lo sigue.
function a2Columnas(wb) {
  const ws = wb.getWorksheet(A2_HOJA);
  if (!ws) return [];
  const cols = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const t = a2Texto(ws.getCell(A2_FILA_TITULOS, c)).trim();
    if (!t) continue;
    if (!/ADMINISTRACION|COMERCIALIZACION|EXPLORACION|FINANCIEROS/i.test(t)) continue;
    cols.push({ col: c, letra: a2Letra(c), nombre: t });
  }
  return cols;
}

// Los conceptos: las filas que tienen texto en B y son parte del cuerpo del anexo.
function a2Conceptos(wb) {
  const ws = wb.getWorksheet(A2_HOJA);
  if (!ws) return [];
  const cols = a2Columnas(wb);
  const conceptos = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const texto = a2Texto(ws.getCell(r, A2_COL_CONCEPTO)).trim();
    if (!texto) continue;
    // el cuerpo es donde hay celdas de gasto; la fila de totales tiene SUM de la columna
    let esCuerpo = false;
    for (const x of cols) {
      const f = ws.getCell(r, x.col).formula;
      if (!f) { esCuerpo = esCuerpo || false; continue; }
      if (new RegExp(REF_DISTRIB + "!", "i").test(String(f))) esCuerpo = true;
      else if (/^\s*\+?SUM\(/i.test(String(f))) { esCuerpo = false; break; }
    }
    // una fila sin fórmulas pero con concepto también es del cuerpo (todavía sin cuenta)
    const tieneSum = cols.some(x => {
      const f = ws.getCell(r, x.col).formula;
      return f && /^\s*\+?SUM\(/i.test(String(f));
    });
    if (tieneSum) continue;
    if (!esCuerpo && !/^-?\s*$/.test(texto) && texto.length > 2) esCuerpo = true;
    if (esCuerpo) conceptos.push({ fila: r, texto });
  }
  return conceptos;
}

// Dónde cae cada cuenta madre. Devuelve una fila por madre, con el concepto y la columna del
// Anexo II que la leen — o null si no la lee nadie.
function a2Mapa(wb, madres) {
  const ws = wb.getWorksheet(A2_HOJA);
  if (!ws) return { lineas: [], porFilaSaldos: new Map() };
  const cols = a2Columnas(wb);

  const porFilaSaldos = new Map();
  for (let r = 1; r <= ws.rowCount; r++) {
    for (const x of cols) {
      const f = ws.getCell(r, x.col).formula;
      if (!f) continue;
      const texto = String(f);
      const reRef = a2ReRef();
      let m;
      while ((m = reRef.exec(texto)) !== null) {
        const fila = +m[2];
        if (!porFilaSaldos.has(fila)) porFilaSaldos.set(fila, []);
        porFilaSaldos.get(fila).push({
          anexoFila: r, col: x.col, columna: x.nombre, letra: x.letra,
          concepto: a2Texto(ws.getCell(r, A2_COL_CONCEPTO)).trim(),
        });
      }
    }
  }

  const lineas = (madres || []).map(m => ({
    filaSaldos: m.fila, codigo: m.codigo, nombre: m.nombre,
    donde: porFilaSaldos.get(m.fila) || [],
  }));
  return { lineas, porFilaSaldos };
}

// El control que hace que esto sea seguro. Se corre antes y después de cada cambio.
function a2Verificar(wb, madres) {
  const { lineas } = a2Mapa(wb, madres);
  const dobles = lineas.filter(l => l.donde.length > 1);
  const sinAnexo = lineas.filter(l => l.donde.length === 0);
  return {
    ok: dobles.length === 0 && sinAnexo.length === 0,
    total: lineas.length,
    unaVez: lineas.filter(l => l.donde.length === 1).length,
    dobles, sinAnexo,
  };
}

// Saca la referencia a `SALDOS!G<filaSaldos>` de una celda. Si era la única, la celda queda
// vacía: dejar un "+" suelto la rompe.
function a2QuitarRef(ws, anexoFila, col, filaSaldos) {
  const cell = ws.getCell(anexoFila, col);
  const f = String(cell.formula || "");
  if (!f) return false;
  const re = new RegExp(`\\+?\\s*${REF_DISTRIB}!\\$?[A-Z]{1,3}\\$?${filaSaldos}(?![0-9])`, "ig");
  const nueva = f.replace(re, "").trim();
  if (nueva === f.trim()) return false;
  if (!nueva || /^\+*$/.test(nueva)) cell.value = null;
  else cell.value = { formula: nueva.replace(/^\+*/, "+") };
  return true;
}

// Y la agrega a otra celda, sumándola a lo que ya hubiera.
function a2AgregarRef(ws, anexoFila, col, filaSaldos, colSaldos) {
  const cell = ws.getCell(anexoFila, col);
  const ref = `${refDeHoja(hojaDistrib(ws.workbook).name)}!${colSaldos}${filaSaldos}`;
  const f = String(cell.formula || "").trim();
  if (!f) { cell.value = { formula: `+${ref}` }; return; }
  if (new RegExp(`${REF_DISTRIB}!\\$?[A-Z]{1,3}\\$?${filaSaldos}(?![0-9])`, "i").test(f)) return;  // ya está
  cell.value = { formula: `${f}+${ref}` };
}

// Mueve una cuenta madre a otro concepto y/o a otra columna del Anexo II.
//
// Se verifica ANTES y DESPUÉS: si el movimiento dejara la madre contada dos veces o fuera del
// anexo, se deshace. Es la única forma de que esto no pueda perder ni duplicar plata en
// silencio.
// `colSaldosDefecto` es la columna de SALDOS que se usa cuando la línea todavía no está en el
// anexo y no hay de dónde copiarla: G en pesos, C en dólares. Los dos maestros son archivos
// distintos y no comparten geometría.
function a2Mover({ wb, madres, filaSaldos, anexoFilaDestino, colDestino,
                   colSaldosDefecto = "G", log = () => {} }) {
  const ws = wb.getWorksheet(A2_HOJA);
  if (!ws) throw new Error(`El maestro no tiene la hoja '${A2_HOJA}'.`);

  const antes = a2Verificar(wb, madres);
  const linea = antes && a2Mapa(wb, madres).lineas.find(l => l.filaSaldos === filaSaldos);
  if (!linea) throw new Error("Esa cuenta madre no está en el maestro. NO se tocó el archivo.");
  if (linea.donde.length > 1) {
    throw new Error(`"${linea.nombre}" hoy entra ${linea.donde.length} veces al Anexo II. ` +
                    `Eso hay que resolverlo primero. NO se tocó el archivo.`);
  }

  const origen = linea.donde[0] || null;
  const cols = a2Columnas(wb);
  const destino = cols.find(x => x.col === colDestino);
  if (!destino) throw new Error("Esa no es una columna de gastos del Anexo II. NO se tocó el archivo.");

  // De qué columna de SALDOS se lee el subtotal: la misma que ya usaba, o G por defecto
  const colSaldos = origen ? (String(ws.getCell(origen.anexoFila, origen.col).formula)
    .match(new RegExp(`${REF_DISTRIB}!\\$?([A-Z]{1,3})\\$?${filaSaldos}(?![0-9])`, "i")) || [])[1] : null;

  const guardado = origen ? String(ws.getCell(origen.anexoFila, origen.col).formula || "") : null;
  const guardadoDestino = String(ws.getCell(anexoFilaDestino, colDestino).formula || "");

  if (origen) a2QuitarRef(ws, origen.anexoFila, origen.col, filaSaldos);
  a2AgregarRef(ws, anexoFilaDestino, colDestino, filaSaldos, colSaldos || colSaldosDefecto);

  // Se compara contra CÓMO ESTABA, no contra cero. En pesos las dos cosas son lo mismo porque
  // el maestro arranca con 0 dobles y 0 fuera del anexo, pero el de dólares tiene 6 líneas
  // legítimamente fuera (la diferencia de cambio, el ROBO, los otros ingresos): exigir cero
  // ahí rechazaba cualquier movimiento. Lo que no se puede es EMPEORAR.
  const despues = a2Verificar(wb, madres);
  if (despues.unaVez !== antes.unaVez ||
      despues.dobles.length > antes.dobles.length ||
      despues.sinAnexo.length > antes.sinAnexo.length) {
    // deshacer
    if (origen) ws.getCell(origen.anexoFila, origen.col).value = guardado ? { formula: guardado } : null;
    ws.getCell(anexoFilaDestino, colDestino).value = guardadoDestino ? { formula: guardadoDestino } : null;
    throw new Error(
      `El cambio dejaba el Anexo II inconsistente (${despues.dobles.length} cuenta(s) contadas dos ` +
      `veces, ${despues.sinAnexo.length} sin aparecer). Se deshizo, NO se tocó el archivo.`);
  }

  const conceptoDestino = a2Texto(ws.getCell(anexoFilaDestino, A2_COL_CONCEPTO)).trim();
  log(`  "${linea.codigo} ${linea.nombre}"` +
      (origen ? ` sale de "${origen.concepto}" (${origen.columna})` : "") +
      ` y pasa a "${conceptoDestino}" (${destino.nombre}).`);
  return { desde: origen, hasta: { anexoFila: anexoFilaDestino, columna: destino.nombre, concepto: conceptoDestino } };
}

if (typeof module !== "undefined") {
  module.exports = {
    a2Columnas, a2Conceptos, a2Mapa, a2Verificar, a2Mover, a2QuitarRef, a2AgregarRef, a2Letra,
  };
}
