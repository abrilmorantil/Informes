// Las fechas que los informes muestran ADENTRO tienen que salir del período que se emite, no
// de la fecha en que se corre la app ni de lo que quedó escrito la vez pasada. Casi todos los
// informes arrastran saldos de períodos anteriores, así que un maestro se reusa mes a mes y
// sus fechas quedan viejas sin que nada avise.
//
// Dos decisiones que hacen que esto sea seguro:
//
// 1) Se reescriben SÓLO las fechas que coinciden con el cierre que el archivo tiene hoy.
//    Cualquier otra fecha es un dato distinto y se deja como está. Gracias a esa regla el
//    "Aumento de capital 16/06/2026" de `Pat.Neto` —un hecho real, no el período— queda
//    intacto sin necesidad de nombrarlo. Una lista de celdas a tocar habría que mantenerla a
//    mano y se desactualiza en cuanto el archivo cambia; así el criterio viaja con el dato.
//
//    El precio de esa regla es que una fecha que quedó atrasada no se recupera sola: si el
//    título "AL 30.04.2026" del `Anexo I` se quedó en abril, en junio ya no coincide con el
//    cierre y no se lo toca más. Por eso los maestros tienen que arrancar con sus fechas al
//    día; de ahí en adelante se mantienen solas. (Ese título es el de la columna NETO
//    RESULTANTE, que se recalcula todos los meses: no es una columna de comparación.)
//
// 2) La redacción NO se rearma desde una plantilla: se conserva la del archivo y sólo cambian
//    día, mes y año. El maestro de pesos dice "Al 30 de junio 2026" y el de dólares "Al 30 de
//    junio de 2026"; el `Anexo I` de pesos escribe "Junio" con mayúscula y el de dólares
//    "junio" con minúscula. Cada uno tiene que seguir diciendo lo suyo.
//
// Las fechas con barras (dd/mm/aaaa) no se tocan nunca: en estos archivos ese formato lo usan
// los hechos puntuales, no los períodos.

const FP_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

const fpSinTildes = (s) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "");

// "setiembre" convive con "septiembre" en los archivos de la usuaria
function fpIndiceMes(nombre) {
  const n = fpSinTildes(nombre).toLowerCase();
  const i = FP_MESES.findIndex(m => fpSinTildes(m) === n);
  if (i >= 0) return i + 1;
  return n === "setiembre" ? 9 : 0;
}

function fpUltimoDia(anio, mes) {
  return new Date(Date.UTC(anio, mes, 0)).getUTCDate();
}

// "2026-07-31" -> {anio, mes, dia}. Acepta también "2026-07" (toma el último día del mes).
function fpPartes(periodo) {
  const m = /^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$/.exec(String(periodo || "").trim());
  if (!m) return null;
  const anio = +m[1], mes = +m[2];
  if (mes < 1 || mes > 12) return null;
  const dia = m[3] ? +m[3] : fpUltimoDia(anio, mes);
  if (dia < 1 || dia > fpUltimoDia(anio, mes)) return null;
  return { anio, mes, dia };
}

const fpISO = (p) =>
  `${p.anio}-${String(p.mes).padStart(2, "0")}-${String(p.dia).padStart(2, "0")}`;

const fpIguales = (a, b) => !!a && !!b && a.anio === b.anio && a.mes === b.mes && a.dia === b.dia;

// Copia el estilo de mayúsculas del texto que se reemplaza: "Junio"/"junio"/"JUNIO".
function fpComoEl(modelo, texto) {
  const soloLetras = fpSinTildes(modelo).replace(/[^A-Za-z]/g, "");
  if (soloLetras && soloLetras === soloLetras.toUpperCase()) return texto.toUpperCase();
  if (/^[A-ZÁÉÍÓÚÑ]/.test(modelo)) return texto.charAt(0).toUpperCase() + texto.slice(1);
  return texto;
}

// Copia el relleno con ceros: si decía "06" el mes nuevo va con dos dígitos, si decía "6" no.
const fpComoNum = (modelo, n) =>
  String(n).padStart(String(modelo).length >= 2 ? 2 : 1, "0");

// "30 de junio de 2026" / "30 de junio 2026" (el "de" antes del año se conserva tal cual)
const FP_RE_LARGA = /(\d{1,2})(\s+de\s+)([A-Za-zÁÉÍÓÚÑáéíóúñ]+)(\s+(?:de\s+)?)(\d{4})/g;
// "30.06.2026" / "30.6.2026" / "31.5.26" — el año va con cuatro dígitos o con dos.
// Los dos dígitos hacen falta por el encabezado "AL 31.5.26" del `Anexo I` de pesos, que es
// el título de la columna NETO RESULTANTE: la columna se recalcula todos los meses, así que
// el título tiene que seguir al período. Mientras el año de dos dígitos no se reconocía, esa
// celda no la veía nadie y se quedó en mayo.
const FP_RE_PUNTOS = /(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})(?!\d)/g;

// Un año de dos dígitos es de este siglo: en estos archivos no hay fechas anteriores a 2000.
const fpAnio = (t) => (String(t).length === 2 ? 2000 + +t : +t);

// Reescribe en `texto` las fechas que sean exactamente `viejo`. Devuelve el texto nuevo, o
// null si no cambió nada. `otras` recibe las fechas que se encontraron y NO se tocaron.
function fpReescribirTexto(texto, viejo, nuevo, otras) {
  let hubo = false;
  const anotar = (p) => { if (otras) otras.push(fpISO(p)); };

  let salida = String(texto).replace(FP_RE_LARGA, (todo, d, sep1, mes, sep2, anio) => {
    const nroMes = fpIndiceMes(mes);
    if (!nroMes) return todo;
    const hallada = { anio: +anio, mes: nroMes, dia: +d };
    if (!fpIguales(hallada, viejo)) { anotar(hallada); return todo; }
    hubo = true;
    return fpComoNum(d, nuevo.dia) + sep1 + fpComoEl(mes, FP_MESES[nuevo.mes - 1]) +
           sep2 + String(nuevo.anio);
  });

  salida = salida.replace(FP_RE_PUNTOS, (todo, d, m, anio) => {
    const hallada = { anio: fpAnio(anio), mes: +m, dia: +d };
    if (!fpIguales(hallada, viejo)) { anotar(hallada); return todo; }
    hubo = true;
    // el año sale con los mismos dígitos que tenía: "31.5.26" no se convierte en "30.6.2026"
    const anioNuevo = String(anio).length === 2
      ? String(nuevo.anio).slice(-2) : String(nuevo.anio);
    return `${fpComoNum(d, nuevo.dia)}.${fpComoNum(m, nuevo.mes)}.${anioNuevo}`;
  });

  return hubo ? salida : null;
}

// ---------------------------------------------------------------- celdas y libros (ExcelJS)

function fpFechaDeCelda(cell) {
  const v = cell.value;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return { anio: v.getUTCFullYear(), mes: v.getUTCMonth() + 1, dia: v.getUTCDate() };
  }
  return null;
}

function fpTextoDeCelda(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "string") return v;
  if (typeof v === "object" && v.richText) return v.richText.map(t => t.text).join("");
  return null;
}

// Reescribe una celda si lleva el cierre viejo. Las celdas con fórmula NO se tocan: toman su
// valor de otra y cambiarlas romperia el arrastre que el archivo ya tiene armado.
function fpReescribirCelda(cell, viejo, nuevo, otras) {
  if (cell.formula) return null;

  const fecha = fpFechaDeCelda(cell);
  if (fecha) {
    if (!fpIguales(fecha, viejo)) { if (otras) otras.push(fpISO(fecha)); return null; }
    const antes = fpISO(fecha);
    cell.value = new Date(Date.UTC(nuevo.anio, nuevo.mes - 1, nuevo.dia));
    return { antes, despues: fpISO(nuevo) };
  }

  const texto = fpTextoDeCelda(cell);
  if (texto === null) return null;
  const nuevoTexto = fpReescribirTexto(texto, viejo, nuevo, otras);
  if (nuevoTexto === null) return null;
  const antes = texto;
  cell.value = nuevoTexto;
  return { antes, despues: nuevoTexto };
}

// Recorre el libro entero. Devuelve qué cambió y qué fechas quedaron sin tocar, para que la
// app pueda mostrarlo: una fecha que no se actualiza sin que nadie se entere es justamente
// el problema que esto viene a resolver.
function fpReescribirLibro(wb, viejo, nuevo) {
  const cambios = [], otras = [];
  wb.eachSheet((ws) => {
    for (let r = 1; r <= ws.rowCount; r++) {
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = ws.getCell(r, c);
        const cambio = fpReescribirCelda(cell, viejo, nuevo, otras);
        if (cambio) {
          cambios.push(Object.assign({ hoja: ws.name, celda: `${ws.getColumn(c).letter}${r}` }, cambio));
        }
      }
    }
  });
  return { cambios, otrasFechas: [...new Set(otras)].sort() };
}

// El cierre que el archivo tiene HOY, leído de una celda que se sabe que lo lleva.
// Se prueba primero como fecha real y después como texto.
function fpCierreDeCelda(cell) {
  if (!cell) return null;
  const fecha = fpFechaDeCelda(cell);
  if (fecha) return fecha;
  const texto = fpTextoDeCelda(cell) ||
    (cell.value && typeof cell.value === "object" && typeof cell.value.result === "string"
      ? cell.value.result : null);
  if (!texto) return null;

  FP_RE_LARGA.lastIndex = 0;
  const l = FP_RE_LARGA.exec(texto);
  if (l && fpIndiceMes(l[3])) return { anio: +l[5], mes: fpIndiceMes(l[3]), dia: +l[1] };
  FP_RE_PUNTOS.lastIndex = 0;
  const p = FP_RE_PUNTOS.exec(texto);
  if (p) return { anio: fpAnio(p[3]), mes: +p[2], dia: +p[1] };
  return null;
}

// Cómo se muestra un período en pantalla: "julio 2026 (cierre 31/07/2026)".
function fpDescribir(p) {
  if (!p) return "(sin período)";
  return `${FP_MESES[p.mes - 1]} ${p.anio} (cierre ` +
         `${String(p.dia).padStart(2, "0")}/${String(p.mes).padStart(2, "0")}/${p.anio})`;
}

// El mes siguiente al de un período, cerrando en su último día. Es lo que se propone cuando
// hay un cierre anterior registrado.
function fpMesSiguiente(p) {
  const mes = p.mes === 12 ? 1 : p.mes + 1;
  const anio = p.mes === 12 ? p.anio + 1 : p.anio;
  return { anio, mes, dia: fpUltimoDia(anio, mes) };
}

if (typeof module !== "undefined") {
  module.exports = {
    FP_MESES, fpIndiceMes, fpUltimoDia, fpPartes, fpISO, fpIguales,
    fpReescribirTexto, fpReescribirCelda, fpReescribirLibro,
    fpCierreDeCelda, fpDescribir, fpMesSiguiente,
  };
}
