// Estado de Resultados (EE RR) — tercera salida del informe, en la misma corrida.
//
// No hay lógica nueva de mapeo de cuentas: sus cifras son, una por una, las mismas que
// el balance en dólares ya expone en las hojas `Anexo II` y `Resultados`. El problema es
// que ahí son FÓRMULAS: las calcula Excel al abrir el archivo, así que el motor no las
// tiene. Este módulo las reproduce en JavaScript.
//
// La clave es que no se hardcodea nada: se leen las fórmulas del propio maestro y se
// evalúan contra los saldos que la corrida acaba de calcular. Si mañana cambia la
// estructura del Anexo II, esto la sigue.
//
// Verificado contra junio 2026, con los valores que el maestro tiene cacheados:
//   Anexo II F102 (Administración) = 530.007,27   H102 (Exploración) = 616.736,47
//   Resultados C17 (ajuste por traducción) = -46.243,87   C21 (resultado) = -1.192.987,61

// --------------------------------------------------- evaluador mínimo de fórmulas
//
// Alcanza con lo que usan estas dos hojas: sumas y restas de referencias a SALDOS, al
// Anexo II y a la propia columna, más SUM() de un rango vertical y números sueltos.
function evaluarFormula(formula, ctx) {
  const texto = String(formula || "").trim();
  if (!texto) return 0;

  // SUM(C12:C20) vertical y SUM(F10:I10) horizontal: se recorre el rectángulo entero.
  // El Anexo II usa las dos formas (la columna TOTALES de cada fila suma F..I).
  const colAIdx = (t) => t.split("").reduce((n2, ch) => n2 * 26 + (ch.charCodeAt(0) - 64), 0);
  const idxACol = (n2) => { let t = ""; while (n2 > 0) { const r = (n2 - 1) % 26; t = String.fromCharCode(65 + r) + t; n2 = (n2 - r - 1) / 26; } return t; };
  let f = texto.replace(/SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)/gi, (_, c1, r1, c2, r2) => {
    const fDesde = Math.min(+r1, +r2), fHasta = Math.max(+r1, +r2);
    const cDesde = Math.min(colAIdx(c1), colAIdx(c2)), cHasta = Math.max(colAIdx(c1), colAIdx(c2));
    let total = 0;
    for (let r = fDesde; r <= fHasta; r++) {
      for (let c = cDesde; c <= cHasta; c++) total += ctx.local(idxACol(c), r);
    }
    return `(${total})`;
  });

  // referencias a otras hojas y a la propia, reemplazadas por su valor
  f = f.replace(/'?([A-Za-z_][A-Za-z0-9_ .]*)'?!\$?([A-Z]{1,3})\$?(\d+)/g,
    (_, hoja, col, fila) => `(${ctx.hoja(hoja.trim(), col, +fila)})`);
  f = f.replace(/(?<![A-Z0-9_$!.)])\$?([A-Z]{1,3})\$?(\d+)(?!\s*\()/g,
    (_, col, fila) => `(${ctx.local(col, +fila)})`);

  // sólo queda aritmética simple; si aparece otra cosa, se avisa en vez de inventar
  if (!/^[-+*/().\d\s]*$/.test(f)) return { error: `no pude evaluar: ${texto}` };
  try {
    // eslint-disable-next-line no-new-func
    const v = Function(`"use strict";return (${f || 0});`)();
    return typeof v === "number" && isFinite(v) ? v : 0;
  } catch (e) {
    return { error: `no pude evaluar: ${texto}` };
  }
}

// --------------------------------------------------- los totales del estado
//
// `valorDeFilaSaldos(fila)` devuelve el importe que la corrida dejó en esa fila de
// SALDOS. Sale de los destinos que ya calculó `resolverDestinosDolares`.
function totalesEstadoResultados(wb, valorDeFilaSaldos) {
  const a2 = wb.getWorksheet("Anexo II");
  const res = wb.getWorksheet("Resultados");
  if (!a2 || !res) throw new Error("El maestro de dólares no tiene las hojas 'Anexo II' y 'Resultados'.");

  const avisos = [];
  const numero = (v) => (typeof v === "number" && isFinite(v) ? v : 0);

  // Hojas que la app NO toca —Anexo I sobre todo— y de las que el Anexo II sí se
  // alimenta: `Anexo II!F20 (DEPRECIACIONES) = +'Anexo I'!I24`, la amortización del
  // período, que se lleva a mano. Como el motor nunca las escribe, su valor guardado
  // sigue siendo el bueno y se lee tal cual.
  const usadasDeOtrasHojas = [];
  const valorOtraHoja = (nombre, col, fila) => {
    const ws = wb.getWorksheet(nombre);
    if (!ws) { avisos.push(`El maestro no tiene la hoja '${nombre}', que el Anexo II necesita.`); return 0; }
    const cell = ws.getCell(`${col}${fila}`);
    const v = numero(cell.result !== undefined ? cell.result : cell.value);
    usadasDeOtrasHojas.push({ hoja: nombre, celda: `${col}${fila}`, valor: v });
    return v;
  };

  // 1) El Anexo II: cada celda es la suma de las filas de SALDOS que referencia.
  const valorAnexo = (col, fila) => {
    const cell = a2.getCell(`${col}${fila}`);
    if (cell.formula) {
      const v = evaluarFormula(cell.formula, {
        hoja: (h, c, r) => {
          if (/^SALDOS$/i.test(h)) return numero(valorDeFilaSaldos(r));
          if (/^Anexo II$/i.test(h)) return numero(valorAnexo(c, r));
          return valorOtraHoja(h, c, r);
        },
        local: (c, r) => numero(valorAnexo(c, r)),
      });
      if (v && v.error) { avisos.push(`Anexo II!${col}${fila}: ${v.error}`); return 0; }
      return numero(v);
    }
    return numero(cell.value);
  };

  // la fila de totales es la que suma la columna entera (SUM(F10:F101))
  let filaTotales = null;
  a2.eachRow((row, r) => {
    const f = row.getCell(6).formula || "";     // columna F
    if (/^SUM\(F\d+:F\d+\)$/i.test(f)) filaTotales = r;
  });
  if (!filaTotales) throw new Error("No encontré la fila de totales del Anexo II.");

  const anexo = {
    fila: filaTotales,
    totales: valorAnexo("E", filaTotales),
    administracion: valorAnexo("F", filaTotales),
    comercializacion: valorAnexo("G", filaTotales),
    exploracion: valorAnexo("H", filaTotales),
    financieros: valorAnexo("I", filaTotales),
  };

  // 2) La hoja Resultados, evaluada con esos totales.
  const cacheRes = new Map();
  const valorRes = (col, fila) => {
    const k = `${col}${fila}`;
    if (cacheRes.has(k)) return cacheRes.get(k);
    cacheRes.set(k, 0);                     // corta una referencia circular
    const cell = res.getCell(k);
    let v;
    if (cell.formula) {
      v = evaluarFormula(cell.formula, {
        hoja: (h, c, r) => {
          if (/^SALDOS$/i.test(h)) return numero(valorDeFilaSaldos(r));
          if (/^Anexo II$/i.test(h)) return numero(valorAnexo(c, r));
          if (/^Balance$/i.test(h)) return 0;   // son textos de encabezado, no importes
          return valorOtraHoja(h, c, r);
        },
        local: (c, r) => numero(valorRes(c, r)),
      });
      if (v && v.error) { avisos.push(`Resultados!${k}: ${v.error}`); v = 0; }
    } else {
      v = numero(cell.value);
    }
    cacheRes.set(k, numero(v));
    return numero(v);
  };

  // Se ubican las líneas por su ETIQUETA, no por número de fila, para que sobreviva a
  // que el archivo cambie de forma.
  const etiqueta = (r) => {
    const v = res.getCell(r, 1).value;
    const t = (v && typeof v === "object") ? (v.result || "") : (v || "");
    return String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
  };
  // `excluir` hace falta porque hay etiquetas que se contienen: la fila
  // "Pérdida antes de impuesto a las Ganancias" matchea también "IMPUESTO A LAS
  // GANANCIAS" y, sin filtrar, el impuesto tomaba el valor del resultado.
  const buscar = (frases, excluir) => {
    for (let r = 1; r <= res.rowCount; r++) {
      const e = etiqueta(r);
      if (!e) continue;
      if (excluir && excluir.some(x => e.indexOf(x) >= 0)) continue;
      if (frases.some(f => e.indexOf(f) >= 0)) return r;
    }
    return null;
  };

  const filaGanancia = buscar(["GANANCIA NETA"]);
  const filaOtros = buscar(["OTROS INGRESOS Y EGRESOS"]);
  const filaExtraord = buscar(["RESULTADOS EXTRAORDINARIOS"]);
  const filaFinanc = buscar(["RESULTADOS FINANCIEROS Y POR TENENCIA"]);
  const filaAntes = buscar(["ANTES DE IMPUESTO"]);
  const filaImpuesto = buscar(["IMPUESTO A LAS GANANCIAS"], ["ANTES DE", "DESPUES DE"]);
  const filaEjercicio = buscar(["DEL EJERCICIO"], ["ANTES DE"]);

  const val = (fila) => (fila ? valorRes("C", fila) : 0);

  const totales = {
    anexo,
    ingresosOperacion: val(filaGanancia),
    // el EE RR muestra los gastos en negativo; el Anexo II los trae en positivo
    gastosOperacion: -anexo.exploracion,
    gastosAdministracion: -anexo.administracion,
    extraordinarios: val(filaExtraord),
    ajusteTraduccion: val(filaFinanc),
    otrosIngresos: val(filaOtros),
    antesDeImpuestos: val(filaAntes),
    impuesto: val(filaImpuesto),
    resultadoEjercicio: val(filaEjercicio),
    // de qué celdas ajenas al motor se tomó un importe (Anexo I, sobre todo)
    tomadoDeOtrasHojas: usadasDeOtrasHojas.filter(x => Math.abs(x.valor) > 0.005),
    avisos,
  };
  totales.totalResultadoOperacion = totales.gastosOperacion + totales.gastosAdministracion + totales.extraordinarios;
  return totales;
}

// --------------------------------------------------- controles de consistencia
//
// Misma filosofía que el resto del sistema: si algo no concilia, se avisa; no se fuerza.
function verificarEERR(t) {
  const avisos = [...(t.avisos || [])];
  const cerca = (a, b) => Math.abs(a - b) < 0.02;

  if (!cerca(t.anexo.comercializacion, 0) || !cerca(t.anexo.financieros, 0)) {
    avisos.push(
      `El Anexo II trae gastos de comercialización (${t.anexo.comercializacion.toFixed(2)}) ` +
      `o financieros (${t.anexo.financieros.toFixed(2)}) distintos de cero. El Estado de ` +
      `Resultados no tiene una línea para ellos, así que su total NO concilia con la hoja ` +
      `Resultados por esa diferencia. Hay que decidir dónde exponerlos.`);
  }

  const sumaEstado = t.totalResultadoOperacion + t.ajusteTraduccion + t.otrosIngresos;
  if (!cerca(sumaEstado, t.antesDeImpuestos)) {
    avisos.push(`El resultado antes de impuestos no cierra: el estado da ${sumaEstado.toFixed(2)} ` +
                `y la hoja Resultados ${t.antesDeImpuestos.toFixed(2)}.`);
  }
  if (!cerca(t.resultadoEjercicio, t.antesDeImpuestos - t.impuesto)) {
    avisos.push(`El resultado del ejercicio no cierra contra el resultado antes de impuestos.`);
  }
  return { concilia: avisos.length === 0, avisos };
}

if (typeof module !== "undefined") {
  module.exports = { evaluarFormula, totalesEstadoResultados, verificarEERR };
}
