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
const { insertRowEn, borrarFilaEn } = require(path.join(AQUI, "formula_hojas.js"));
global.insertRowEn = insertRowEn;
global.borrarFilaEn = borrarFilaEn;
const { PARAMS, derivarMapeoMaestro, madresResultados, insertarHijaEnMadre } = require(path.join(AQUI, "motor_balances.js"));
global.madresResultados = madresResultados;
global.insertarHijaEnMadre = insertarHijaEnMadre;
const { categoriasPesos, quitarCuentaDeCategoria, editarCuenta, agregarCuentaACategoria, gcFilasDelBloque } = require(path.join(AQUI, "gestion_categorias.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const buf = fs.readFileSync(path.join(AQUI, "base_pesos.xlsx"));
const abrirCopia = () => abrirWorkbook(buf.slice(0));

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
    const c154 = categorias.find(c => c.filaMadre === 154);
    check(!!c154 && c154.codigo === "42101000" && c154.nombre === "Honorarios legales",
      'categoría 154 identificada ("Honorarios legales")');
    check(c154.miembros.length === 4, `"Honorarios legales" tiene 4 miembros (dio ${c154.miembros.length})`);
    check(c154.miembros.every(m => !m.esFilaCompartida), "ninguno de sus miembros comparte fila con la madre");
    const c200 = categorias.find(c => c.filaMadre === 200);
    check(c200.miembros.some(m => m.esFilaCompartida && m.fila === 200),
      '"Vacaciones" (fila 200) marca su miembro de fila compartida');
    check(sueltas.length > 0 && sueltas.every(s => s.fila < 154 || true),
      `hay cuentas de RESULTADOS sueltas, sin categoría (${sueltas.length})`);
  }

  // ---------------------------------------------- sacar del medio de un "rango"
  {
    const wb = await abrirCopia();
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const { categorias } = categoriasPesos(wb, mapeo);
    const cat = categorias.find(c => c.filaMadre === 154); // rango 155-158
    quitarCuentaDeCategoria(wb, cat, 156, () => {});
    const ws = wb.getWorksheet("SALDOS");
    check(formulaDe(ws, 154, 7) === "SUM(F155:F157)",
      `el subtotal de "Honorarios legales" quedó en SUM(F155:F157) (dio ${formulaDe(ws, 154, 7)})`);
    // lo que estaba en 157 (Servicio Notarial) ahora está en 156, corrido un lugar
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c154b = cat2.find(c => c.filaMadre === 154);
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
    const cat = categorias.find(c => c.filaMadre === 154); // rango 155-158, hasta=158
    const ws = wb.getWorksheet("SALDOS");
    const filaVecinaAntes = textoDeFila(ws, 159); // la próxima madre, para confirmar que no se corrompe
    quitarCuentaDeCategoria(wb, cat, 158, () => {});
    check(formulaDe(ws, 154, 7) === "SUM(F155:F157)",
      `sacar la ÚLTIMA fila (158) achica el subtotal a SUM(F155:F157) (dio ${formulaDe(ws, 154, 7)})`);
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c154b = cat2.find(c => c.filaMadre === 154);
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
    const cat = categorias.find(c => c.filaMadre === 197); // lista [197,198,199], 197 compartida
    check(cat.bloque.tipo === "lista", "la categoría de prueba es de tipo lista");
    const ws = wb.getWorksheet("SALDOS");
    quitarCuentaDeCategoria(wb, cat, 199, () => {});
    check(formulaDe(ws, 197, 7) === "+F197+F198",
      `el subtotal de la lista quedó en "+F197+F198" (dio ${formulaDe(ws, 197, 7)})`);
    const mapeo2 = derivarMapeoMaestro(wb, "pesos");
    const { categorias: cat2 } = categoriasPesos(wb, mapeo2);
    const c197b = cat2.find(c => c.filaMadre === 197);
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
    quitarCuentaDeCategoria(wb, recategorizar().find(c => c.codigo === "42101000"), 156, () => {});
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
    const cat200 = categorias.find(c => c.filaMadre === 200);
    let tiroError = false;
    try { quitarCuentaDeCategoria(wb, cat200, 200, () => {}); } catch (e) { tiroError = /comparte la fila/.test(e.message); }
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
    const cat = categorias.find(c => c.filaMadre === 154);
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
    const cat = categorias.find(c => c.filaMadre === 159); // "Honorarios profesionales", rango
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
