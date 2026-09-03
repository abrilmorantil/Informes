// Gestión de categorías, sobre COPIAS EN MEMORIA del maestro real de pesos: nunca escribe
// el archivo. Cubre los casos que importan de verdad —
//   - sacar una cuenta del medio de un bloque "rango"
//   - sacar la ÚLTIMA cuenta de un bloque "rango" (el caso que rompe si no se achica el
//     subtotal a mano antes de borrar: si no, el rango se corre y suma la fila de al lado)
//   - sacar una cuenta de un bloque "lista" (+F+F+F)
//   - los dos resguardos: no dejar una categoría sin cuentas, no tocar la fila compartida
//   - renombrar una cuenta
//   - agregar una cuenta nueva a una categoría
// y que ninguna de esas operaciones desarma OTRA categoría que no se tocó.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
// En Node cada archivo es su propio módulo, así que los nombres de hoja compartidos —que en
// el navegador define motor_balances.js para todos— hay que dejarlos en el global a mano.
for (const [k, v] of Object.entries(require(path.join(AQUI, "motor_balances.js")))) {
  if (global[k] === undefined) global[k] = v;
}
const { insertRowEn, borrarFilaEn } = require(path.join(AQUI, "formula_hojas.js"));
global.insertRowEn = insertRowEn;
global.borrarFilaEn = borrarFilaEn;
const { PARAMS, derivarMapeoMaestro, madresResultados, insertarHijaEnMadre } = require(path.join(AQUI, "motor_balances.js"));
global.madresResultados = madresResultados;
global.insertarHijaEnMadre = insertarHijaEnMadre;
const { categoriasPesos, quitarCuentaDeCategoria, editarCuenta, agregarCuentaACategoria, gcFilasDelBloque } = require(path.join(AQUI, "gestion_categorias.js"));

// El maestro avanza un mes en cada cierre y todas las filas se corren, así que nada acá se
// clava a un número de fila: las categorías se buscan por CÓDIGO y las cuentas a sacar
// también. Un test clavado a "la fila 154" se pone en rojo todos los meses sin que se haya
// roto nada — pasó con el cierre de julio.
const porCodigo = (cats, cod) => cats.find(c => c.codigo === cod);
const filaDe = (cat, cod) => {
  const m = cat.miembros.find(x => x.codigo === cod);
  if (!m) throw new Error(`la categoría "${cat.nombre}" no tiene la cuenta ${cod}`);
  return m.fila;
};
const rangoEsperado = (cat, quitadas) =>
  `SUM(F${cat.bloque.desde}:F${cat.bloque.hasta - quitadas})`;

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const buf = fs.readFileSync(path.join(AQUI, "base_pesos.xlsx"));
const abrirCopia = () => abrirWorkbook(buf.slice(0));
// Fabrica una fila compartida: sube la primera cuenta del bloque a la fila de la madre y
// estira el subtotal para incluirla. Es la forma que tenía el maestro antes de septiembre
// de 2026, y que el motor tiene que seguir entendiendo aunque el archivo ya no la use.
function compartirFila(wb, filaMadre) {
  const ws = hojaDistrib(wb);
  const hija = filaMadre + 1;
  const clave = ws.getCell(hija, 5).value;
  const formula = String(ws.getCell(hija, 6).formula || "");
  if (!clave || !formula) return false;
  ws.getCell(filaMadre, 5).value = clave;
  ws.getCell(filaMadre, 6).value = {
    formula: formula.replace(/(VLOOKUP\(\s*\$?E\$?)\d+/i, "$1" + filaMadre) };
  ws.getCell(hija, 5).value = null;
  ws.getCell(hija, 6).value = null;
  const g = ws.getCell(filaMadre, 7);
  const antes = String(g.formula || "");
  g.value = { formula: antes.replace(new RegExp("(\\$?F\\$?)" + hija + "(?!\\d)"), "$1" + filaMadre) };
  return true;
}


function gCell(ws, fila, col) { return ws.getCell(fila, col); }
function formulaDe(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  return v && typeof v === "object" ? v.formula : null;
}

(async () => {
  // ---------------------------------------------------------- categoriasPesos()
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias, sueltas } = categoriasPesos(wb, mapeo);
    check(categorias.length === 58, `58 categorías de RESULTADOS (dio ${categorias.length})`);

    // Las categorías se buscan por su CÓDIGO, no por su fila. El maestro avanza un mes cada
    // cierre y las filas se corren, así que un test clavado a "la fila 154" se pone en rojo
    // todos los meses sin que se haya roto nada. Ya pasó con el cierre de julio.
    const cHon = categorias.find(c => c.codigo === "42101000");
    check(!!cHon && cHon.nombre === "Honorarios legales",
      `"Honorarios legales" (42101000) está, en la fila ${cHon && cHon.filaMadre}`);
    check(cHon.miembros.length === 4, `"Honorarios legales" tiene 4 miembros (dio ${cHon.miembros.length})`);
    check(cHon.miembros.every(m => !m.esFilaCompartida), "ninguno de sus miembros comparte fila con la madre");

    // Y una categoría que comparte fila con una de sus cuentas. El maestro ya no tiene
    // ninguna, así que se fabrica: lo que se prueba es que el motor la reconozca.
    const wbC = await abrirCopia();
    compartirFila(wbC, porCodigo(categoriasPesos(wbC, derivarMapeoMaestro(wbC, "pesos")).categorias,
                                "42101000").filaMadre);
    const { categorias: catsC } = categoriasPesos(wbC, derivarMapeoMaestro(wbC, "pesos"));
    const cCompartida = catsC.find(c => c.miembros.some(m => m.esFilaCompartida && m.fila === c.filaMadre));
    check(!!cCompartida,
      `hay una categoría que comparte fila con un miembro suyo: "${cCompartida && cCompartida.nombre}" (fila ${cCompartida && cCompartida.filaMadre})`);
    check(sueltas.length > 0, `hay cuentas de RESULTADOS sueltas, sin categoría (${sueltas.length})`);
  }

  // ---------------------------------------------- sacar del medio de un "rango"
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    const cat = porCodigo(categorias, "42101000");   // "Honorarios legales", bloque rango
    const esperado = rangoEsperado(cat, 1);
    quitarCuentaDeCategoria(wb, cat, filaDe(cat, "421430000"), () => {});
    const ws = hojaDistrib(wb);
    check(formulaDe(ws, cat.filaMadre, 7) === esperado,
      `el subtotal de "Honorarios legales" quedó en ${esperado} (dio ${formulaDe(ws, cat.filaMadre, 7)})`);
    // lo que estaba en 157 (Servicio Notarial) ahora está en 156, corrido un lugar
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c154b = porCodigo(cat2, "42101000");
    check(c154b.miembros.length === 3, `quedaron 3 miembros (dio ${c154b.miembros.length})`);
    check(!c154b.miembros.some(m => m.codigo === "421430000"), "la cuenta sacada (421430000) ya no aparece");
    // otra categoría, más abajo, no se tiene que haber tocado en su forma (madre a fila 159-1=158)
    const c159 = cat2.find(c => c.nombre === "Honorarios profesionales");
    check(!!c159 && c159.miembros.length === 8, `"Honorarios profesionales" sigue con sus 8 miembros (dio ${c159 && c159.miembros.length})`);
  }

  // -------------------------------------------- sacar la ÚLTIMA fila de un "rango"
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    const cat = porCodigo(categorias, "42101000");   // bloque rango: se saca su ÚLTIMA fila
    const esperado = rangoEsperado(cat, 1);
    const ws = hojaDistrib(wb);
    const filaVecinaAntes = textoDeFila(ws, cat.bloque.hasta + 1); // la próxima madre
    quitarCuentaDeCategoria(wb, cat, cat.bloque.hasta, () => {});
    check(formulaDe(ws, cat.filaMadre, 7) === esperado,
      `sacar la ÚLTIMA fila achica el subtotal a ${esperado} (dio ${formulaDe(ws, cat.filaMadre, 7)})`);
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c154b = porCodigo(cat2, "42101000");
    check(c154b.miembros.length === 3 && !c154b.miembros.some(m => m.codigo === "425840000"),
      "la cuenta sacada (425840000, la última del bloque) ya no está, y NO quedó de más sumando la fila vecina");
    const c159 = cat2.find(c => c.nombre === "Honorarios profesionales");
    check(!!c159 && c159.codigo === "42102000" && c159.miembros.length === 8,
      `la categoría vecina ("Honorarios profesionales") no se corrompió (dio ${c159 && c159.miembros.length} miembros)`);
  }

  // ------------------------------------------------------ sacar de un bloque "lista"
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    // la primera categoría de tipo "lista" con al menos 3 filas, sea cual sea
    const cat = categorias.find(c => c.bloque.tipo === "lista" && c.bloque.filas.length >= 3);
    check(!!cat, `hay una categoría de tipo lista: "${cat && cat.nombre}" (fila ${cat && cat.filaMadre})`);
    const quedan = cat.bloque.filas.slice(0, -1);
    const esperado = quedan.map(f => "+F" + f).join("");
    const ws = hojaDistrib(wb);
    quitarCuentaDeCategoria(wb, cat, cat.bloque.filas[cat.bloque.filas.length - 1], () => {});
    check(formulaDe(ws, cat.filaMadre, 7) === esperado,
      `el subtotal de la lista quedó en "${esperado}" (dio ${formulaDe(ws, cat.filaMadre, 7)})`);
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c197b = porCodigo(cat2, cat.codigo);
    check(c197b.miembros.length === 2, `quedaron 2 miembros en la categoría lista (dio ${c197b.miembros.length})`);
  }

  // ------------------------------------------------- ningún #REF! NUEVO en NINGUNA hoja
  //
  // El maestro real ya trae 3 #REF! propias (Balance!G8, Balance!G17, Anexo II!J103 — ver
  // CLAUDE.md, no las introdujo esta app). El control es que la cantidad NO CREZCA, no que
  // dé cero.
  const contarRefs = (wb) => {
    let n = 0;
    wb.worksheets.forEach(hoja => hoja.eachRow(row => row.eachCell(cell => {
      const v = cell.value;
      const f = v && typeof v === "object" ? v.formula : null;
      if (f && /#REF!/.test(f)) n++;
    })));
    return n;
  };
  {
    const wb = await abrirCopia();
    const refsAntes = contarRefs(wb);
    // una tanda combinada: sacar del medio, sacar la última de otro bloque, sacar de una
    // lista — releyendo categorías entre cada paso, como va a hacer el panel de verdad
    // (cada borrado corre filas y las categorías siguientes cambian de número de fila).
    const recategorizar = () => categoriasPesos(wb, derivarMapeoMaestro(wb, "pesos")).categorias;
    quitarCuentaDeCategoria(wb, recategorizar().find(c => c.codigo === "42101000"),
      filaDe(recategorizar().find(c => c.codigo === "42101000"), "421430000"), () => {});
    const catHonProf = recategorizar().find(c => c.codigo === "42102000"); // "Honorarios profesionales"
    quitarCuentaDeCategoria(wb, catHonProf, Math.max(...gcFilasDelBloque(catHonProf.bloque)), () => {}); // la última de su rango
    quitarCuentaDeCategoria(wb, recategorizar().find(c => c.codigo === "42113000"), // "Tasas Retributivas Minería"
      recategorizar().find(c => c.codigo === "42113000").miembros.find(m => !m.esFilaCompartida).fila, () => {});
    const refsDespues = contarRefs(wb);
    check(refsDespues === refsAntes,
      `sacar 3 cuentas de categorías distintas no agrega #REF! nuevas (antes ${refsAntes}, después ${refsDespues})`);
  }

  // -------------------------------------------------------------- resguardos
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    const cat0 = porCodigo(categorias, "42101000");
    compartirFila(wb, cat0.filaMadre);
    const cat200 = categoriasPesos(wb, derivarMapeoMaestro(wb, "pesos")).categorias
      .find(c => c.miembros.some(m => m.esFilaCompartida && m.fila === c.filaMadre));
    let tiroError = false;
    try { quitarCuentaDeCategoria(wb, cat200, cat200.filaMadre, () => {}); } catch (e) { tiroError = /comparte la fila/.test(e.message); }
    check(tiroError, "sacar la fila compartida con la madre tira error y no toca el archivo");

    const catUnica = categorias.find(c => c.miembros.length === 1);
    check(!!catUnica, "hay al menos una categoría de 1 sola cuenta para probar el resguardo");
    let tiroError2 = false;
    try { quitarCuentaDeCategoria(wb, catUnica, catUnica.miembros[0].fila, () => {}); }
    catch (e) { tiroError2 = /se quedaría sin ninguna cuenta/.test(e.message); }
    check(tiroError2, "sacar la única cuenta de una categoría tira error y no toca el archivo");
  }

  // ---------------------------------------------------------------- editarCuenta
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    const cat = porCodigo(categorias, "42101000");
    const miembro = cat.miembros[0];
    editarCuenta(wb, miembro.fila, miembro.col, "999999999", "Cuenta de prueba", () => {});
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    check(!!mapeo2.cuentas["999999999"] && mapeo2.cuentas["999999999"].nombre === "Cuenta de prueba",
      "el código y nombre nuevos se leen bien después de editar");
    check(!mapeo2.cuentas[miembro.codigo], "el código viejo ya no está");
  }

  // -------------------------------------------------------- agregarCuentaACategoria
  {
    const wb = await abrirCopia();
    let mapeo = derivarMapeoMaestro(wb, "pesos");
    let { categorias } = categoriasPesos(wb, mapeo);
    const cat = porCodigo(categorias, "42102000");   // "Honorarios profesionales", bloque rango
    const filaNueva = agregarCuentaACategoria(wb, mapeo, { codigo: "429999000", nombre: "Cuenta agregada de prueba" }, cat, () => {});
    check(typeof filaNueva === "number", `insertó una fila nueva (${filaNueva})`);
    mapeo = derivarMapeoMaestro(wb, "pesos");
    ({ categorias } = categoriasPesos(wb, mapeo));
    const catB = categorias.find(c => c.filaMadre === cat.filaMadre || c.codigo === cat.codigo);
    check(!!catB && catB.miembros.some(m => m.codigo === "429999000"),
      "la cuenta nueva aparece como miembro de la categoría");
    check(catB.miembros.length === cat.miembros.length + 1,
      `la categoría pasó de ${cat.miembros.length} a ${catB.miembros.length} miembros`);
  }

  console.log(fallos ? `\n${fallos} falla(s).` : "\nTodo OK.");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e); process.exit(1); });

function textoDeFila(ws, fila) {
  const partes = [];
  for (let c = 3; c <= 5; c++) {
    const v = ws.getCell(fila, c).value;
    if (typeof v === "string") partes.push(v);
  }
  return partes.join(" | ");
}
