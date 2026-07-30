// Lee un EE RR ya terminado y saca de ahí las cifras del mes, para poder arrancar con la
// columna MES ANTERIOR cargada sin esperar a que la app haya emitido el mes previo.
//
// Del archivo se toma la columna MES ACTUAL: lo que para ese informe era "el mes" es,
// para el siguiente, "el mes anterior". Las líneas se buscan por su ETIQUETA, no por
// número de fila, porque el que se venía armando a mano no tiene por qué coincidir fila
// a fila con el que genera la app.

const LINEAS_EERR = [
  { campo: "ingresosOperacion",    re: /^INGRESOS DE OPERACION/ },
  { campo: "gastosOperacion",      re: /^GASTOS DE OPERACION/ },
  { campo: "gastosAdministracion", re: /^GASTOS DE ADMINISTRACION/ },
  { campo: "extraordinarios",      re: /^RESULTADOS? EXTRAORDINARIO/ },
  { campo: "ajusteTraduccion",     re: /^AJUSTE POR TRADUCCION/ },
  { campo: "otrosIngresos",        re: /^OTROS INGRESOS/ },
  { campo: "antesDeImpuestos",     re: /^RESULTADO ANTES DE IMPUESTO/ },
  // en la hoja el impuesto se muestra NEGATIVO, porque esa columna se suma; adentro se
  // guarda con el signo de la hoja Resultados, que lo resta
  { campo: "impuesto",             re: /^IMPUESTO A LAS GANANCIAS/, invertir: true },
  { campo: "resultadoEjercicio",   re: /^RESULTADO DEL EJERCICIO/ },
];

// Las que tienen que estar sí o sí para dar el archivo por bueno. Son las que se leen del
// archivo y no se pueden deducir de nada: los subtotales, en cambio, si vienen sin calcular
// se rehacen.
const OBLIGATORIAS_EERR = ["gastosOperacion", "gastosAdministracion"];

function _normE(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/\s+/g, " ").trim();
}

function _numE(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/\s/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", "."));
    if (isFinite(n)) return n;
  }
  return null;
}

// Devuelve { totales, avisos, hoja, columna, etiquetaColumna } o { totales:null, avisos }.
function leerEERRDeArchivo(libro) {
  for (const nombre of libro.SheetNames) {
    const hoja = libro.Sheets[nombre];
    if (!hoja || !hoja["!ref"]) continue;
    const r = XLSX.utils.decode_range(hoja["!ref"]);
    const celda = (f, c) => hoja[XLSX.utils.encode_cell({ r: f, c })];
    const val = (f, c) => { const x = celda(f, c); return x ? x.v : null; };

    // 1) la columna del mes: se busca el encabezado "MES ACTUAL"
    let colValor = null, filaEncabezado = null, etiquetaColumna = null;
    for (let f = r.s.r; f <= Math.min(r.e.r, r.s.r + 30) && colValor === null; f++) {
      for (let c = r.s.c; c <= r.e.c; c++) {
        if (/^MES ACTUAL/.test(_normE(val(f, c)))) {
          colValor = c; filaEncabezado = f; etiquetaColumna = String(val(f, c)).trim();
          break;
        }
      }
    }
    if (colValor === null) continue;

    // 2) cada línea, por su etiqueta. Se prueban TODAS las celdas a la izquierda, no sólo
    //    la primera con algo: la columna A lleva el número de línea (67, 68, 69...) y
    //    quedarse con el primer texto no vacío daba "68" en vez de "Gastos de Operación".
    const totales = {};
    const encontradas = [];
    for (let f = filaEncabezado; f <= r.e.r; f++) {
      for (let c = r.s.c; c < colValor; c++) {
        const etiqueta = _normE(val(f, c));
        if (!etiqueta) continue;
        for (const L of LINEAS_EERR) {
          if (totales[L.campo] !== undefined || !L.re.test(etiqueta)) continue;
          const n = _numE(val(f, colValor));
          if (n === null) continue;
          totales[L.campo] = L.invertir ? -n : n;
          encontradas.push(L.campo);
        }
      }
    }

    // Los subtotales del EE RR son fórmulas, así que si el archivo no pasó por Excel vienen
    // sin valor. Se rehacen con las mismas cuentas que hace la hoja.
    const derivadas = [];
    const componentes = ["gastosOperacion", "gastosAdministracion", "extraordinarios",
                         "ajusteTraduccion", "otrosIngresos"];
    const hayComponentes = componentes.some(k => totales[k] !== undefined);
    if (totales.antesDeImpuestos === undefined && hayComponentes) {
      totales.antesDeImpuestos = componentes.reduce((a, k) => a + (totales[k] || 0), 0);
      derivadas.push("resultado antes de impuestos");
    }
    if (totales.resultadoEjercicio === undefined && totales.antesDeImpuestos !== undefined) {
      totales.resultadoEjercicio = totales.antesDeImpuestos - (totales.impuesto || 0);
      derivadas.push("resultado del ejercicio");
    }

    const faltan = OBLIGATORIAS_EERR.filter(k => totales[k] === undefined);
    if (faltan.length === OBLIGATORIAS_EERR.length) continue;   // no era esta hoja

    const avisos = [];
    if (faltan.length) {
      avisos.push(`No encontré en el archivo: ${faltan.join(", ")}. Sin eso el EE RR del ` +
                  `mes que viene queda incompleto.`);
    }
    if (derivadas.length) {
      avisos.push(`El archivo no traía calculado el ${derivadas.join(" ni el ")}: se rehízo ` +
                  `con las mismas cuentas que hace la hoja. Si lo abrís y guardás en Excel ` +
                  `antes de subirlo, se toma tal cual está.`);
    }
    // lo que no aparece va en cero, como en el archivo que genera la app
    for (const L of LINEAS_EERR) if (totales[L.campo] === undefined) totales[L.campo] = 0;

    avisos.push(`Se tomó la columna "${etiquetaColumna}" de la hoja "${nombre}". ` +
                `Es la que corresponde: lo que ahí es el mes actual pasa a ser el mes anterior.`);
    avisos.push(...verificarEERRImportado(totales));
    return { totales, encontradas, avisos, hoja: nombre, columna: colValor, etiquetaColumna };
  }
  return { totales: null, encontradas: [], hoja: null, columna: null, etiquetaColumna: null,
           avisos: ["No encontré un Estado de Resultados en el archivo. Tiene que tener una " +
                    "columna rotulada \"MES ACTUAL\" y las líneas de gastos de operación, " +
                    "de administración y el resultado del ejercicio."] };
}

// El propio EE RR se controla a sí mismo: si las cuentas no cierran, lo más probable es que
// se haya leído la columna equivocada o que falte una línea.
function verificarEERRImportado(t) {
  const avisos = [];
  const cerca = (a, b) => Math.abs(a - b) <= 1;
  const sumaAntes = t.gastosOperacion + t.gastosAdministracion + t.extraordinarios +
                    t.ajusteTraduccion + t.otrosIngresos;
  if (t.antesDeImpuestos && !cerca(sumaAntes, t.antesDeImpuestos)) {
    avisos.push(`Ojo: las líneas suman ${sumaAntes.toFixed(2)} pero el archivo dice que el ` +
                `resultado antes de impuestos es ${t.antesDeImpuestos.toFixed(2)}.`);
  }
  const esperado = (t.antesDeImpuestos || sumaAntes) - t.impuesto;
  if (!cerca(esperado, t.resultadoEjercicio)) {
    avisos.push(`Ojo: el resultado del ejercicio del archivo (${t.resultadoEjercicio.toFixed(2)}) ` +
                `no coincide con lo que dan sus propias líneas (${esperado.toFixed(2)}).`);
  }
  return avisos;
}

// Sirve tanto .xlsx como .xls: el informe que se venía armando a mano suele ser .xls.
function abrirLibroEERR(buffer) {
  return XLSX.read(buffer, { type: "array" });
}

if (typeof module !== "undefined") {
  module.exports = { LINEAS_EERR, OBLIGATORIAS_EERR, leerEERRDeArchivo,
                     verificarEERRImportado, abrirLibroEERR };
}
