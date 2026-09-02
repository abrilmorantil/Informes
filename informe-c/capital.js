// Control del capital en el Estado de Evolución del Patrimonio Neto.
//
// `Pat.Neto` NO lee el capital de la contabilidad: lo tiene escrito a mano en el maestro, en
// la fila "Saldos al inicio del ejercicio" más una fila por cada aumento. Mientras la cuenta
// "Capital Suscripto" no se mueva da igual —en junio 2026 coincidían exactamente en las dos
// monedas— pero cuando entra un aporte, el Activo lo refleja (viene en el export) y el
// Patrimonio Neto no. El balance se abre exactamente por esa diferencia.
//
// Y se abre EN SILENCIO, que es lo peor: los valores que quedan guardados en el archivo son
// los del maestro y cierran; la diferencia aparece recién cuando Excel recalcula al abrirlo.
// Pasó en la corrida de agosto 2026, con el capital subiendo 261.309.720,00 en pesos y
// 176.680,00 en dólares (el mismo aporte a TC 1.479,00): los dos balances salieron abiertos
// por esos importes exactos y no hubo ningún aviso.
//
// Por eso este control corre ANTES de habilitar la descarga y frena. El importe que falta lo
// calcula la app —es la diferencia contra la cuenta— pero la fecha del aporte la pone la
// usuaria, porque ese dato no está en ninguna parte del export.

const CAP_RE_CUENTA = /capital\s+suscripto/i;
const CAP_RE_INICIO = /saldos\s+al\s+inicio/i;
const CAP_RE_CIERRE = /^\s*saldos\s+al\s+\d/i;         // "Saldos al 30.06.2026"
const CAP_RE_AUMENTO = /aumento\s+de\s+capital/i;

function _capTexto(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined) return _capTexto(v.result);
    return "";
  }
  return String(v);
}

function _capNum(cell) {
  const v = cell && cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

// Una constante escrita a mano, no una fórmula: es lo que cuenta como "capital declarado".
function _capConstante(cell) {
  if (!cell || cell.formula) return null;
  return typeof cell.value === "number" ? cell.value : null;
}

// La geometría de Pat.Neto no se hardcodea: las dos monedas tienen la hoja armada distinto
// (en pesos la apertura está en la 12, el aumento en la 22 y el cierre en la 27; en dólares
// son la 12, la 17 y la 21). Se ubica todo por las etiquetas de la columna B, y la columna
// del capital sale de la fila del aumento que ya está cargado, que hace de modelo.
function ubicarPatNeto(ws) {
  let filaInicio = null, filaCierre = null, filaAumento = null;
  for (let r = 1; r <= ws.rowCount; r++) {
    const b = _capTexto(ws.getCell(r, 2).value).trim();
    if (!b) continue;
    if (filaInicio === null && CAP_RE_INICIO.test(b)) filaInicio = r;
    else if (CAP_RE_AUMENTO.test(b)) filaAumento = r;
    else if (CAP_RE_CIERRE.test(b)) filaCierre = r;
  }
  if (filaInicio === null || filaCierre === null) return null;

  let colCapital = null;
  if (filaAumento !== null) {
    for (let c = 3; c <= ws.columnCount; c++) {
      if (_capConstante(ws.getCell(filaAumento, c)) !== null) { colCapital = c; break; }
    }
  }
  // Apenas se arrastra el cierre al inicio, la hoja queda sin ninguna fila de aumento cargada
  // hasta que entre la del mes nuevo. En ese rato la columna del capital sale de la apertura,
  // que siempre lo tiene escrito a mano.
  if (colCapital === null) {
    for (let c = 3; c <= ws.columnCount; c++) {
      const v = _capConstante(ws.getCell(filaInicio, c));
      if (v !== null && v !== 0) { colCapital = c; break; }
    }
  }
  if (colCapital === null) return null;

  // Y la fila que hace de modelo para el cableado: la del aumento si esta cargada, y si no una
  // ranura vacia que ya lo tenga --que es justamente como queda la fila del mes anterior
  // despues de limpiarla, porque el arrastre le deja las formulas puestas.
  if (filaAumento === null) {
    for (let r = filaInicio + 1; r < filaCierre && filaAumento === null; r++) {
      if (_capTexto(ws.getCell(r, 2).value).trim()) continue;
      const cap = ws.getCell(r, colCapital);
      if (cap.formula) continue;                       // una fila de subtotales, no una ranura
      if (cap.value !== null && cap.value !== undefined && cap.value !== "") continue;
      for (let c = 1; c <= ws.columnCount; c++) {
        if (ws.getCell(r, c).formula) { filaAumento = r; break; }
      }
    }
  }
  if (filaAumento === null) return null;
  return { filaInicio, filaCierre, filaAumento, colCapital };
}

// Lo que Pat.Neto declara: la apertura más cada aumento cargado. Sólo constantes — las
// fórmulas de esa columna son subtotales y contarían dos veces.
function capitalDeclarado(ws, u) {
  let total = 0;
  const filas = [];
  for (let r = u.filaInicio; r < u.filaCierre; r++) {
    const v = _capConstante(ws.getCell(r, u.colCapital));
    if (v === null || v === 0) continue;
    total += v;
    filas.push({ fila: r, etiqueta: _capTexto(ws.getCell(r, 2).value).trim(), importe: v });
  }
  return { total: Math.round(total * 100) / 100, filas };
}

// De qué columna de Hoja1 sale el saldo. Se lee del propio VLOOKUP de SALDOS en vez de
// hardcodearla (en pesos es la 5 y en dólares la 4): si el maestro cambia, el control lo
// sigue solo.
function columnaValorDeHoja1(wb, porDefecto) {
  const ws = hojaDistrib(wb);
  if (ws) {
    for (let r = 1; r <= ws.rowCount; r++) {
      for (let c = 1; c <= ws.columnCount; c++) {
        const f = ws.getCell(r, c).formula;
        const m = f && new RegExp(
          `VLOOKUP\\([^,]+,\\s*${REF_SUMAS}![^,]+,\\s*(\\d+)\\s*,`, "i").exec(String(f));
        if (m) return +m[1];
      }
    }
  }
  return porDefecto;
}

// Lo que dice la contabilidad. Se lee de Hoja1 —donde el motor deja el export de ESTE mes— y
// no de SALDOS, porque ahí la cuenta suele ser una fórmula sin resultado guardado.
// El saldo viene acreedor (negativo); el capital se compara en positivo.
function capitalContable(wb, colValorPorDefecto) {
  const ws = hojaSumas(wb);
  if (!ws) return null;
  const col = columnaValorDeHoja1(wb, colValorPorDefecto);
  for (let r = 1; r <= ws.rowCount; r++) {
    const clave = _capTexto(ws.getCell(r, 1).value).trim();
    if (!CAP_RE_CUENTA.test(clave)) continue;
    const v = _capNum(ws.getCell(r, col));
    if (v === null) continue;
    return { valor: Math.abs(Math.round(v * 100) / 100), fila: r, clave, columna: col };
  }
  return null;
}

// El control. `falta` es lo que hay que agregarle a Pat.Neto para que el balance cierre.
function controlarCapital(wb, colValorPorDefecto) {
  const ws = wb.getWorksheet("Pat.Neto");
  if (!ws) return { ok: true, sinHoja: true };
  const u = ubicarPatNeto(ws);
  if (!u) {
    return { ok: false, sinUbicar: true,
             motivo: "No pude ubicar en 'Pat.Neto' la fila de apertura, la del aumento de " +
                     "capital y la de cierre, así que no puedo controlar el capital." };
  }
  const declarado = capitalDeclarado(ws, u);
  const contable = capitalContable(wb, colValorPorDefecto);
  if (!contable) {
    return { ok: false, sinCuenta: true, ubic: u, declarado,
             motivo: "No encontré la cuenta 'Capital Suscripto' en Hoja1." };
  }
  const falta = Math.round((contable.valor - declarado.total) * 100) / 100;
  return {
    ok: Math.abs(falta) < 0.005,
    ubic: u, declarado, contable, falta,
  };
}

// ------------------------------------------------------- agregar el aumento que falta

// Corre una fórmula simple de la misma hoja a otra fila ("+C22" -> "+C16"). Sólo se usa para
// copiar el cableado de la fila que hace de modelo, que en los dos maestros son referencias
// de una celda a otra de la misma fila.
function _capCorrerFormula(formula, delta) {
  return String(formula).replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g,
    (todo, d1, col, d2, fila) => (d2 ? todo : `${d1}${col}${d2}${+fila + delta}`));
}

// Que la fila nueva quede DENTRO de lo que suma la fila de cierre. En dólares no lo estaba:
// `C21` era `+C15+C17`, una lista explícita que se saltea la 16, mientras que su vecina `E21`
// sí era `SUM(E15:E17)`. Con esa lista, una línea agregada en la 16 no se contaba en la
// columna del capital integrado y el estado quedaba descuadrado contra sí mismo.
function asegurarTotalIncluya(ws, filaCierre, col, fila) {
  const cell = ws.getCell(filaCierre, col);
  const f = cell.formula;
  if (!f) return null;
  const letra = ws.getColumn(col).letter;

  const rango = new RegExp(`^SUM\\(\\$?${letra}\\$?(\\d+):\\$?${letra}\\$?(\\d+)\\)$`, "i");
  const m = rango.exec(String(f).trim());
  if (m) {
    const a = +m[1], b = +m[2];
    if (fila >= a && fila <= b) return null;                  // ya la incluye
    const nueva = `SUM(${letra}${Math.min(a, fila)}:${letra}${Math.max(b, fila)})`;
    cell.value = { formula: nueva };
    return { celda: `${letra}${filaCierre}`, antes: `=${f}`, despues: `=${nueva}` };
  }

  // lista explícita del tipo "+C15+C17"
  const refs = [...String(f).matchAll(new RegExp(`\\$?${letra}\\$?(\\d+)`, "gi"))].map(x => +x[1]);
  if (refs.length && String(f).replace(new RegExp(`[+\\s]|\\$?${letra}\\$?\\d+`, "gi"), "") === "") {
    if (refs.includes(fila)) return null;
    const a = Math.min(...refs, fila), b = Math.max(...refs, fila);
    const nueva = `SUM(${letra}${a}:${letra}${b})`;
    cell.value = { formula: nueva };
    return { celda: `${letra}${filaCierre}`, antes: `=${f}`, despues: `=${nueva}` };
  }
  return null;
}

// ------------------------------------------------------- arrastrar el cierre a la apertura

// El EEPN no se acumula solo. Cada mes, lo que quedó en "Saldos al <fecha>" tiene que pasar a
// "Saldos al inicio", y la línea del aumento del mes anterior tiene que desaparecer para que
// entre la del mes nuevo. Es lo que el cliente venía haciendo a mano:
//
//     junio:  inicio 14.700.098.884 + aumento 16/06 148.546.600 = cierre 14.848.645.484
//     julio:  inicio 14.848.645.484 + aumento 24/07 261.309.720 = cierre 15.109.955.204
//
// Ojo con qué se arrastra: sólo las columnas que están escritas a mano en la apertura, que son
// las del capital. "Resultados no asignados" NO se toca --J12 es la fórmula que lo trae de
// SALDOS y da el saldo al inicio del EJERCICIO, que no se mueve en todo el año-- y la
// "Ganancia (Pérdida) del ejercicio" es acumulada, así que arrastrarla la contaría dos veces.
//
// Qué filas se arrastran: las que tienen fecha ANTERIOR al mes que se está cerrando. Así la
// operación se puede repetir sin miedo: si se vuelve a correr el mismo mes, el aumento nuevo
// ya es de este mes y se queda donde está.
function _capFechaEtiqueta(etiqueta) {
  const m = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/.exec(String(etiqueta || ""));
  if (!m) return null;
  const mes = +m[2];
  let anio = +m[3];
  if (anio < 100) anio += 2000;
  if (mes < 1 || mes > 12) return null;
  return { anio, mes };
}

// Lo que vale una celda de la fila del aumento, incluso si es parte del cableado. En pesos
// F=+C y en dolares E=+C: son columnas que muestran el mismo importe con otro titulo. Hace
// falta porque en el maestro de dolares la apertura de "Capital suscripto" (E12) esta escrita
// a mano, no es una formula como en pesos, asi que si no se arrastra tambien se queda atras y
// el mes siguiente el estado abre por menos de lo que cerro.
function _capValorEfectivo(ws, fila, col, vuelta) {
  const cell = ws.getCell(fila, col);
  if (!cell.formula) return typeof cell.value === "number" ? cell.value : 0;
  if ((vuelta || 0) > 3) return 0;
  const m = /^\+?\$?([A-Z]{1,3})\$?(\d+)$/.exec(String(cell.formula).trim());
  if (!m) return 0;                       // una suma o algo mas complejo: no es cableado
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return _capValorEfectivo(ws, +m[2], c, (vuelta || 0) + 1);
}

function pasarCierreAlInicio(wb, periodo) {
  const ws = wb.getWorksheet("Pat.Neto");
  if (!ws) return { sinHoja: true, arrastradas: [], limpiadas: [] };
  const u = ubicarPatNeto(ws);
  if (!u) {
    throw new Error("No pude ubicar en 'Pat.Neto' la apertura y el cierre, así que no arrastré nada.");
  }
  if (!periodo || !periodo.anio || !periodo.mes) {
    throw new Error("No sé a qué mes se está cerrando, así que no arrastré nada.");
  }
  const antes = capitalDeclarado(ws, u).total;

  const aArrastrar = [], sinFecha = [], sonDelMes = [];
  for (let r = u.filaInicio + 1; r < u.filaCierre; r++) {
    const et = _capTexto(ws.getCell(r, 2).value).trim();
    if (!et || !CAP_RE_AUMENTO.test(et)) continue;
    const f = _capFechaEtiqueta(et);
    if (!f) { sinFecha.push({ fila: r, etiqueta: et }); continue; }
    if (f.anio > periodo.anio || (f.anio === periodo.anio && f.mes >= periodo.mes)) {
      sonDelMes.push({ fila: r, etiqueta: et });
      continue;
    }
    aArrastrar.push({ fila: r, etiqueta: et });
  }
  if (!aArrastrar.length) {
    return { arrastradas: [], limpiadas: [], sinFecha, sonDelMes, apertura: antes };
  }

  // 1) la apertura se lleva lo que sumaron esas filas, columna por columna
  const arrastradas = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const ap = ws.getCell(u.filaInicio, c);
    const viejo = _capConstante(ap);
    if (viejo === null) continue;                    // fórmulas y textos no se tocan
    let suma = 0;
    for (const x of aArrastrar) suma += _capValorEfectivo(ws, x.fila, c);
    if (!suma) continue;
    const val = Math.round((viejo + suma) * 100) / 100;
    ap.value = val;
    arrastradas.push({ celda: ws.getColumn(c).letter + u.filaInicio, antes: viejo, despues: val });
  }

  // 2) las filas viejas quedan libres, pero con el cableado puesto para el aumento que viene
  const limpiadas = [];
  for (const x of aArrastrar) {
    ws.getCell(x.fila, 2).value = null;
    for (let c = 3; c <= ws.columnCount; c++) {
      const cell = ws.getCell(x.fila, c);
      if (cell.formula) continue;
      if (cell.value !== null && cell.value !== undefined && cell.value !== "") cell.value = null;
    }
    limpiadas.push(x);
  }

  // 3) el control que hace que esto sea seguro: mover plata de una fila a otra no puede
  // cambiar el capital declarado. Si cambió, se arrastró algo que no correspondía.
  const u2 = ubicarPatNeto(ws);
  const despues = u2 ? capitalDeclarado(ws, u2).total : null;
  if (despues === null || Math.abs(despues - antes) >= 0.005) {
    throw new Error(`El arrastre del capital no cerró: "Pat.Neto" declaraba ${antes} y quedó en ` +
                    `${despues}. NO uses este archivo, avisá para revisarlo.`);
  }

  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  return { arrastradas, limpiadas, sinFecha, sonDelMes, apertura: despues };
}

// Escribe el aumento en la primera fila libre entre la apertura y el cierre, copiándole el
// cableado a la fila del aumento que ya estaba cargado. No inventa la fecha: `etiqueta` es lo
// que la usuaria escribió.
function agregarAumentoDeCapital(wb, { etiqueta, importe, colValorPorDefecto }) {
  const ws = wb.getWorksheet("Pat.Neto");
  if (!ws) throw new Error("El archivo no tiene la hoja 'Pat.Neto'.");
  const u = ubicarPatNeto(ws);
  if (!u) throw new Error("No pude ubicar la estructura de 'Pat.Neto'.");
  if (!etiqueta || !String(etiqueta).trim()) throw new Error("Falta el texto del aumento.");
  if (typeof importe !== "number" || !isFinite(importe) || importe === 0) {
    throw new Error("El importe del aumento tiene que ser un número distinto de cero.");
  }

  // Se prefiere una fila que YA tenga el cableado de la fila modelo: el maestro de pesos trae
  // seis ranuras así (16 a 21), preparadas justamente para movimientos de capital, y es donde
  // corresponde que el aumento aparezca. Recién si no hay ninguna se usa una fila vacía
  // cualquiera — que es el caso de dólares, donde no hay ranuras preparadas.
  const cableadoModelo = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    if (ws.getCell(u.filaAumento, c).formula) cableadoModelo.push(c);
  }
  // Una fila esta libre si no tiene etiqueta ni importe. No hace falta excluir la del aumento:
  // si todavia tiene uno cargado, tiene etiqueta; y si esta vacia es porque el arrastre la
  // limpio, que es justo donde corresponde escribir el del mes nuevo.
  const estaLibre = (r) => {
    const b = _capTexto(ws.getCell(r, 2).value).trim();
    const v = ws.getCell(r, u.colCapital).value;
    return !b && (v === null || v === undefined || v === "");
  };
  const yaCableada = (r) => cableadoModelo.length > 0 &&
    cableadoModelo.every(c => !!ws.getCell(r, c).formula);

  // De las ranuras preparadas se toma la ULTIMA, la mas pegada al cierre: es donde estuvo
  // siempre el aumento en los maestros (fila 22 en pesos) y donde queda al arrastrar el mes
  // anterior, asi el estado se ve igual todos los meses en vez de ir subiendo de fila.
  let libre = null;
  for (let r = u.filaCierre - 1; r > u.filaInicio && libre === null; r--) {
    if (estaLibre(r) && yaCableada(r)) libre = r;
  }
  for (let r = u.filaInicio + 1; r < u.filaCierre && libre === null; r++) {
    if (estaLibre(r)) libre = r;
  }
  if (libre === null) {
    throw new Error("No hay ninguna fila libre en 'Pat.Neto' entre la apertura y el cierre " +
                    "para cargar el aumento. Hay que hacerle lugar en el Excel.");
  }

  ws.getCell(libre, 2).value = String(etiqueta).trim();
  ws.getCell(libre, u.colCapital).value = Math.round(importe * 100) / 100;

  // el cableado de la fila modelo: en pesos F=+C y K=+F, en dólares E=+C y J=+E
  const delta = libre - u.filaAumento;
  const cableado = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    const modelo = ws.getCell(u.filaAumento, c);
    if (!modelo.formula) continue;
    const nueva = _capCorrerFormula(modelo.formula, delta);
    ws.getCell(libre, c).value = { formula: nueva };
    cableado.push(`${ws.getColumn(c).letter}${libre} = ${nueva}`);
  }

  // y que la fila de cierre la sume
  const totales = [];
  for (let c = 1; c <= ws.columnCount; c++) {
    if (!ws.getCell(u.filaAumento, c).formula && c !== u.colCapital) continue;
    const arreglo = asegurarTotalIncluya(ws, u.filaCierre, c, libre);
    if (arreglo) totales.push(arreglo);
  }

  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  return {
    fila: libre, etiqueta: String(etiqueta).trim(), importe: Math.round(importe * 100) / 100,
    cableado, totalesArreglados: totales,
    control: controlarCapital(wb, colValorPorDefecto),
  };
}

if (typeof module !== "undefined") {
  module.exports = {
    ubicarPatNeto, capitalDeclarado, capitalContable, columnaValorDeHoja1,
    controlarCapital, agregarAumentoDeCapital, asegurarTotalIncluya, pasarCierreAlInicio,
  };
}
