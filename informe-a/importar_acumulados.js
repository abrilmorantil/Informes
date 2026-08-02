// Toma la hoja "Gastos Acumulados" de un informe ya terminado y la usa como punto de
// partida del maestro.
//
// Hace falta porque esa hoja estuvo fuera de alcance mucho tiempo: el maestro la tiene
// congelada en un mes viejo mientras la app ya cerró varios meses más. Subiendo el último
// informe terminado se empareja de una vez, y de ahí en adelante `avanzarGastosAcumulados`
// la mantiene sola en cada cierre.
//
// El mes al que corresponde el archivo NO se pregunta: sale del rótulo de la columna del
// mes ("JUNIO"), que es lo que el archivo dice de sí mismo.

const MESES_IMP = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
                   "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

function _txtGA(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined) return _txtGA(v.result);
    return "";
  }
  return String(v);
}

const _normGA = (v) => _txtGA(v).normalize("NFD").replace(/[̀-ͯ]/g, "")
  .toUpperCase().replace(/\s+/g, " ").trim();

function _numGA(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

// Ubica la fila de rótulos y qué columna es cada cosa, buscando por su texto.
function ubicarGastosAcumulados(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 30); r++) {
    let colAcum = null, colMes = null, colAnteriores = null;
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const aqui = _normGA(ws.getCell(r, c).value);
      const arriba = _normGA(ws.getCell(r - 1, c).value);
      const abajo = _normGA(ws.getCell(r + 1, c).value);
      if (/^ENERO\s*-/.test(aqui)) colAcum = c;
      else if (MESES_IMP.includes(aqui)) colMes = c;
      else if (arriba === "ACUMULADO" && abajo === "ANTERIORES") colAnteriores = c;
    }
    if (colAcum && colMes) return { fila: r, colAcum, colMes, colAnteriores };
  }
  return null;
}

// Devuelve { mes, rotuloAcum, rotuloMes, proyectos:[{nombre, anteriores, acumulado, mes}] }
function leerGastosAcumulados(wb) {
  const ws = wb.getWorksheet("Gastos Acumulados");
  if (!ws) return { error: "El archivo no tiene la hoja 'Gastos Acumulados'." };
  const cols = ubicarGastosAcumulados(ws);
  if (!cols) {
    return { error: "No encontré los rótulos de 'Gastos Acumulados'. Tiene que haber una " +
                    "columna que diga 'ENERO - <mes>' y otra con el nombre del mes en curso." };
  }

  const rotuloAcum = _txtGA(ws.getCell(cols.fila, cols.colAcum).value).trim();
  const rotuloMes = _txtGA(ws.getCell(cols.fila, cols.colMes).value).trim();
  const mes = MESES_IMP.indexOf(_normGA(rotuloMes)) + 1;
  if (!mes) return { error: `No entiendo a qué mes corresponde la columna "${rotuloMes}".` };

  const proyectos = [];
  for (let r = cols.fila + 1; r <= ws.rowCount; r++) {
    const nombre = _txtGA(ws.getCell(r, 1).value).trim();
    if (!nombre) continue;
    const acumulado = _numGA(ws.getCell(r, cols.colAcum));
    const delMes = _numGA(ws.getCell(r, cols.colMes));
    if (acumulado === null && delMes === null) continue;      // títulos y totales
    if (/^TOTAL/i.test(nombre)) continue;
    proyectos.push({
      nombre,
      anteriores: cols.colAnteriores ? _numGA(ws.getCell(r, cols.colAnteriores)) : null,
      acumulado: acumulado || 0,
      mes: delMes || 0,
    });
  }
  if (!proyectos.length) return { error: "No encontré ningún proyecto con importes en esa hoja." };
  return { mes, rotuloAcum, rotuloMes, proyectos, columnas: cols };
}

// Escribe lo leído en el maestro y lo deja listo para el mes SIGUIENTE: el acumulado pasa a
// incluir el mes del archivo, y los rótulos corren uno.
//
// `anteriores` (la columna de años anteriores) se copia sólo si el archivo la trae, y las
// diferencias se informan: es una columna que la usuaria mantiene a mano.
function sembrarGastosAcumulados(wbMaestro, datos, anio) {
  const ws = wbMaestro.getWorksheet("Gastos Acumulados");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Gastos Acumulados'.");
  const cols = ubicarGastosAcumulados(ws);
  if (!cols) throw new Error("No encontré los rótulos de 'Gastos Acumulados' en el maestro.");
  if (datos.mes === 12) {
    throw new Error("El archivo es de diciembre: el cambio de ejercicio se hace a mano.");
  }

  const porNombre = new Map(datos.proyectos.map(p => [_normGA(p.nombre), p]));
  const actualizados = [], sinCorrespondencia = [], anterioresCambiados = [];

  for (let r = cols.fila + 1; r <= ws.rowCount; r++) {
    const nombre = _txtGA(ws.getCell(r, 1).value).trim();
    if (!nombre || /^TOTAL/i.test(nombre)) continue;
    const celdaAcum = ws.getCell(r, cols.colAcum);
    if (_numGA(celdaAcum) === null && !ws.getCell(r, cols.colMes).formula) continue;

    const p = porNombre.get(_normGA(nombre));
    if (!p) { sinCorrespondencia.push(nombre); continue; }

    const nuevo = Math.round((p.acumulado + p.mes) * 100) / 100;
    celdaAcum.value = nuevo;
    actualizados.push({ nombre, acumulado: p.acumulado, mes: p.mes, nuevo });

    if (cols.colAnteriores && p.anteriores !== null) {
      const celdaAnt = ws.getCell(r, cols.colAnteriores);
      const previo = _numGA(celdaAnt);
      if (previo === null || Math.abs(previo - p.anteriores) > 0.005) {
        anterioresCambiados.push({ nombre, de: previo, a: p.anteriores });
        celdaAnt.value = p.anteriores;
      }
    }
  }

  // los rótulos quedan apuntando al mes siguiente
  ws.getCell(cols.fila, cols.colAcum).value = `ENERO - ${MESES_IMP[datos.mes - 1]}`;
  ws.getCell(cols.fila, cols.colMes).value = MESES_IMP[datos.mes];
  ws.getCell(3, 1).value = `${MESES_IMP[datos.mes]} ${anio}`;

  // el total del acumulado es un número fijo en la fila de TOTALES: si no se rehace, queda con
  // el importe viejo y la hoja muestra importes correctos que no suman (ver el comentario de
  // `repararTotalGastosAcumulados` en motor.js).
  const total = repararTotalGastosAcumulados(ws, cols.colAcum);

  wbMaestro.calcProperties = wbMaestro.calcProperties || {};
  wbMaestro.calcProperties.fullCalcOnLoad = true;

  return {
    actualizados, sinCorrespondencia, anterioresCambiados, total,
    rotuloAcum: `ENERO - ${MESES_IMP[datos.mes - 1]}`,
    rotuloMes: MESES_IMP[datos.mes],
  };
}

if (typeof module !== "undefined") {
  // en el navegador motor.js se carga antes que este archivo y comparten el ámbito global
  global.repararTotalGastosAcumulados = require("./motor.js").repararTotalGastosAcumulados;
  module.exports = { MESES_IMP, ubicarGastosAcumulados, leerGastosAcumulados,
                     sembrarGastosAcumulados };
}
