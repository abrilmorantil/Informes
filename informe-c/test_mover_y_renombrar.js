// Las dos operaciones nuevas del panel de configuración del Informe 04:
//
//   renombrarCategoria    — cambiarle el código o el nombre a una cuenta madre
//   moverCuentaACategoria — meter una cuenta suelta adentro de una cuenta madre
//
// Lo que hay que probar no es que "anden": es que NO puedan romper el balance. Mover una cuenta
// es insertar y después borrar, y si el borrado se descubriera imposible recién al final, la
// cuenta quedaría dos veces y el Anexo II sumaría de más sin que nada lo avise. Por eso el
// resguardo va antes de tocar el archivo, y acá se comprueba justamente eso.
//
//   node informe-c/test_mover_y_renombrar.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const fh = require(path.join(AQUI, "formula_hojas.js"));
for (const k of Object.keys(fh)) global[k] = fh[k];
const M = require(path.join(AQUI, "motor_balances.js"));
for (const [k, v] of Object.entries(M)) { if (global[k] === undefined) global[k] = v; }
const clas = require(path.join(AQUI, "clasificacion.js"));
for (const k of Object.keys(clas)) global[k] = clas[k];
const gest = require(path.join(AQUI, "gestion_categorias.js"));
for (const k of Object.keys(gest)) global[k] = gest[k];
const a2 = require(path.join(AQUI, "anexo2.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const buf = fs.readFileSync(path.join(AQUI, "base_pesos.xlsx"));
const abrirCopia = () => abrirWorkbook(buf.slice(0));
const cats = (wb) => gest.categoriasPesos(wb, M.derivarMapeoMaestro(wb, "pesos")).categorias;
const porCodigo = (l, c) => l.find(x => String(x.codigo) === String(c));
const txt = (c) => {
  const v = c.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
};

(async () => {
  console.log("=== 1) renombrar una categoría ===");
  {
    const wb = await abrirCopia();
    const cat = porCodigo(cats(wb), "42101000");           // "Honorarios legales"
    const ws = M.hojaDistrib(wb);
    const antesMiembros = cat.miembros.map(m => m.codigo).join(",");
    const antesSubtotal = String(ws.getCell(cat.filaMadre, 7).formula || "");

    const r = gest.renombrarCategoria(wb, cat, "42101999", "Honorarios legales y notariales", () => {});
    check(r.nuevo === "42101999 - Honorarios legales y notariales",
      `el rótulo quedó "${r.nuevo}"`);
    check(txt(ws.getCell(cat.filaMadre, 4)) === r.nuevo, "y así está escrito en la columna D");

    const despues = cats(wb);
    const catN = porCodigo(despues, "42101999");
    check(!!catN, "la categoría se reconoce con el código nuevo");
    check(catN && catN.miembros.map(m => m.codigo).join(",") === antesMiembros,
      "sus cuentas son exactamente las mismas");
    check(String(ws.getCell(cat.filaMadre, 7).formula || "") === antesSubtotal,
      "y el subtotal no se tocó");
    check(a2.a2Verificar(wb, M.madresResultados(wb, "pesos")).ok,
      "el Anexo II sigue sano: cada madre entra exactamente una vez");
  }

  console.log("\n=== 2) renombrar no toca el Balance de sumas y saldos ===");
  {
    // El rótulo de la madre es el código del CLIENTE y no existe en Onvio. Si se lo tratara
    // como a una cuenta —renombrando también en la zona de pegado— se estaría renombrando la
    // cuenta equivocada, o ninguna.
    const wb = await abrirCopia();
    const h1 = M.hojaSumas(wb);
    const antes = [];
    h1.eachRow({ includeEmpty: false }, (row, r) => antes.push(r + ":" + txt(row.getCell(1))));
    gest.renombrarCategoria(wb, porCodigo(cats(wb), "42101000"), "42101999", "Otro nombre", () => {});
    const despues = [];
    h1.eachRow({ includeEmpty: false }, (row, r) => despues.push(r + ":" + txt(row.getCell(1))));
    check(antes.join("|") === despues.join("|"),
      "la zona de pegado del export quedó intacta");
  }

  console.log("\n=== 3) lo que renombrar NO deja hacer ===");
  {
    const wb = await abrirCopia();
    const lista = cats(wb);
    const cat = porCodigo(lista, "42101000");
    const otra = lista.find(c => String(c.codigo) !== "42101000");
    let e1 = null;
    try { gest.renombrarCategoria(wb, cat, String(otra.codigo), "Lo que sea", () => {}); }
    catch (e) { e1 = e.message; }
    check(!!e1 && /ya lo usa/.test(e1), "usar el código de otra categoría se rechaza");
    check(txt(M.hojaDistrib(wb).getCell(cat.filaMadre, 4)).indexOf("42101000") === 0,
      "y el archivo quedó como estaba");

    let e2 = null;
    try { gest.renombrarCategoria(wb, cat, "abc", "Nombre", () => {}); } catch (e) { e2 = e.message; }
    check(!!e2 && /código de cuenta válido/.test(e2), "un código que no es un número se rechaza");

    let e3 = null;
    try { gest.renombrarCategoria(wb, cat, "42101999", "  ", () => {}); } catch (e) { e3 = e.message; }
    check(!!e3 && /Falta el nombre/.test(e3), "sin nombre se rechaza");
  }

  console.log("\n=== 4) mover una cuenta suelta adentro de una categoría ===");
  {
    const wb = await abrirCopia();
    const mapeo = M.derivarMapeoMaestro(wb, "pesos");
    const lista = gest.categoriasPesos(wb, mapeo);
    const enCategoria = new Set();
    for (const c of lista.categorias) for (const m of c.miembros) enCategoria.add(String(m.codigo));
    const madres = new Set(M.madresResultados(wb, "pesos").map(m => String(m.codigo)));

    // una suelta que se pueda mover: ni madre, ni referenciada por nadie
    const candidata = lista.sueltas.find(s =>
      !enCategoria.has(String(s.codigo)) && !madres.has(String(s.codigo)) &&
      fh.quienReferenciaLaFila(wb, M.hojaDistrib(wb).name, s.fila).length === 0);
    check(!!candidata, `hay una cuenta suelta que se puede mover: ${candidata && candidata.codigo}`);
    if (!candidata) { console.log(`\n${fallos} FALLA(S)`); process.exit(1); }

    const destino = porCodigo(lista.categorias, "42101000");
    const antesN = destino.miembros.length;
    const totalAntes = Object.keys(mapeo.cuentas).length;

    gest.moverCuentaACategoria(wb, mapeo, candidata.fila, destino, () => {});

    const despues = cats(wb);
    const destinoN = porCodigo(despues, "42101000");
    check(destinoN.miembros.length === antesN + 1,
      `la categoría pasó de ${antesN} a ${destinoN.miembros.length} cuentas`);
    check(destinoN.miembros.some(m => String(m.codigo) === String(candidata.codigo)),
      `y ${candidata.codigo} está adentro`);

    const mapeoDespues = M.derivarMapeoMaestro(wb, "pesos");
    check(Object.keys(mapeoDespues.cuentas).length === totalAntes,
      `el balance sigue teniendo ${totalAntes} cuentas: se movió, no se duplicó`);
    const filas = Object.values(mapeoDespues.cuentas).filter(f => f.nombre === candidata.nombre);
    check(filas.length === 1, "la cuenta aparece una sola vez en todo el archivo");
    check(a2.a2Verificar(wb, M.madresResultados(wb, "pesos")).ok,
      "el Anexo II sigue sano");
  }

  console.log("\n=== 5) lo que mover NO deja hacer ===");
  {
    const wb = await abrirCopia();
    const mapeo = M.derivarMapeoMaestro(wb, "pesos");
    const lista = gest.categoriasPesos(wb, mapeo);
    const destino = porCodigo(lista.categorias, "42101000");
    const ws = M.hojaDistrib(wb);

    // una fila que alguien referencia: se rechaza ANTES de tocar nada
    const atada = lista.sueltas.find(s =>
      fh.quienReferenciaLaFila(wb, ws.name, s.fila).length > 0);
    if (atada) {
      const antesTotal = Object.keys(M.derivarMapeoMaestro(wb, "pesos").cuentas).length;
      let e = null;
      try { gest.moverCuentaACategoria(wb, mapeo, atada.fila, destino, () => {}); }
      catch (x) { e = x.message; }
      check(!!e && /#REF!/.test(e),
        `mover una cuenta que otra fórmula lee se rechaza: ${atada.codigo}`);
      check(Object.keys(M.derivarMapeoMaestro(wb, "pesos").cuentas).length === antesTotal,
        "y el archivo quedó exactamente como estaba (no se insertó nada)");
    } else {
      console.log("  (salteado) no hay ninguna cuenta suelta referenciada por otra fórmula");
    }

    // una cuenta madre no se mueve
    const madre = M.madresResultados(wb, "pesos").find(m => m.fila !== destino.filaMadre);
    let e2 = null;
    try { gest.moverCuentaACategoria(wb, mapeo, madre.fila, destino, () => {}); }
    catch (x) { e2 = x.message; }
    check(!!e2 && /ES una cuenta madre/.test(e2), "mover una categoría adentro de otra se rechaza");

    // una que ya está adentro
    let e3 = null;
    try { gest.moverCuentaACategoria(wb, mapeo, destino.miembros[0].fila, destino, () => {}); }
    catch (x) { e3 = x.message; }
    check(!!e3 && /ya integra/.test(e3), "mover una cuenta a la categoría donde ya está se rechaza");
  }

  console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e.stack); process.exit(1); });
