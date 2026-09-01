// El panel de configuración de cuentas (panel_config_cuentas.js) sobre el maestro real de
// pesos: confirma que arma el resumen y los pendientes (config_balances.js) y que renderiza
// una categoría y sus cuentas sin tirar excepción. La lógica de sacar/agregar/editar cuentas
// se prueba aparte, en test_gestion_categorias.js.
//
// La detección de "dos cuentas en la misma fila" se prueba con una fila ARMADA a propósito, no
// contra el defecto que traía el maestro. Antes se apoyaba en la fila 137 —que decía
// "213010010 Patentes a pagar" en una columna y "212020002 Retenciones SUSS a pagar" en la
// otra— y al arreglarla el test se puso en rojo sin que nada se hubiera roto. Un test que se
// cae cuando se arregla el archivo mide el archivo, no el código.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

global.ExcelJS = require(path.join(AQUI, "..", "informe-a", "vendor", "exceljs.min.js"));
const { abrirWorkbook } = require(path.join(AQUI, "..", "informe-a", "formula_utils.js"));
const { insertRowEn, borrarFilaEn } = require(path.join(AQUI, "formula_hojas.js"));
global.insertRowEn = insertRowEn;
global.borrarFilaEn = borrarFilaEn;
const { PARAMS, derivarMapeoMaestro, lineasDeNota4, filasQueAgrega, madresResultados, insertarHijaEnMadre } = require(path.join(AQUI, "motor_balances.js"));
global.madresResultados = madresResultados;
global.insertarHijaEnMadre = insertarHijaEnMadre;
const { derivarConfigBalance } = require(path.join(AQUI, "config_balances.js"));
const { categoriasPesos, quitarCuentaDeCategoria, editarCuenta, agregarCuentaACategoria } = require(path.join(AQUI, "gestion_categorias.js"));
global.categoriasPesos = categoriasPesos;
global.quitarCuentaDeCategoria = quitarCuentaDeCategoria;
global.editarCuenta = editarCuenta;
global.agregarCuentaACategoria = agregarCuentaACategoria;
global.derivarMapeoMaestro = derivarMapeoMaestro;
global.derivarConfigBalance = derivarConfigBalance;
global.lineasDeNota4 = lineasDeNota4;
global.filasQueAgrega = filasQueAgrega;
global.PARAMS = PARAMS;
global.abrirWorkbook = abrirWorkbook;
const { pccResumenHtml, pccPendientesHtml, gcMiembroHtml, gcCategoriaHtml, gcEsc } = require(path.join(AQUI, "panel_config_cuentas.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

(async () => {
  const wb = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const mapeo = derivarMapeoMaestro(wb, "pesos");
  const lineasNota4 = lineasDeNota4(wb);
  const cfg = derivarConfigBalance(wb, "pesos", mapeo, lineasNota4, PARAMS.pesos, { filasQueAgrega });
  const { categorias } = categoriasPesos(wb, mapeo);

  // No se fija un número exacto de líneas: el maestro crece cada vez que se le conecta una
  // cuenta que estaba suelta, y un test que se cae por eso mide el archivo, no el código.
  check(cfg.lineas.length > 400, `el balance de pesos tiene ${cfg.lineas.length} líneas`);
  check(cfg.resumen.conMadre === 3, `3 casos de madre/hija resueltos solos (dio ${cfg.resumen.conMadre})`);

  // El maestro no tiene que tener filas con dos cuentas distintas. Si aparece alguna, se
  // muestra acá para poder mirarla.
  const conflictos = cfg.avisos.filter(a => a.tipo === "dos_cuentas_en_la_misma_fila");
  check(conflictos.length === 0,
    conflictos.length
      ? `hay ${conflictos.length} fila(s) con dos cuentas: ` +
        conflictos.map(c => `${c.fila} (${c.usa.code} / ${c.tambien.code})`).join(", ")
      : "ninguna fila del maestro tiene dos cuentas distintas escritas");

  // Y que la detección funciona: se arma una fila con dos cuentas y tiene que saltar.
  const wbX = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const wsX = wbX.getWorksheet("SALDOS");
  wsX.getCell(137, 4).value = "212020002 - Retenciones SUSS a pagar";   // como estaba antes
  const cfgX = derivarConfigBalance(wbX, "pesos", derivarMapeoMaestro(wbX, "pesos"),
    lineasDeNota4(wbX), PARAMS.pesos, { filasQueAgrega });
  const c137 = cfgX.avisos.find(a => a.tipo === "dos_cuentas_en_la_misma_fila" && a.fila === 137);
  check(!!c137, "con dos cuentas escritas en una fila, el panel lo marca");
  check(c137 && c137.usa.code === "213010010" && c137.tambien.code === "212020002",
    "y dice cuál usa y cuál es la otra");

  const resumenHtml = pccResumenHtml(cfg, cfg.avisos.length, categorias.length);
  check(resumenHtml.includes(String(cfg.resumen.lineas)), "el resumen HTML muestra el total de líneas");
  check(resumenHtml.includes(String(cfg.avisos.length)), "el resumen HTML muestra la cantidad de pendientes");
  check(resumenHtml.includes(String(categorias.length)), "el resumen HTML muestra la cantidad de categorías");

  const pendientesHtml = pccPendientesHtml(cfgX);
  check(pendientesHtml.includes("213010010") && pendientesHtml.includes("212020002"),
    "el panel dibuja las dos cuentas en conflicto");
  // El maestro no tiene que tener pendientes. Si aparece alguno, se lista para poder verlo.
  check(cfg.avisos.length === 0,
    cfg.avisos.length
      ? `quedan ${cfg.avisos.length} pendiente(s): ` +
        cfg.avisos.map(a => `${a.tipo} (${a.texto || "fila " + a.fila})`).join(", ")
      : "el maestro no tiene pendientes: cada línea del balance sabe de qué cuenta sale");

  // Y que el panel sabe dibujar una línea sin cuenta, con un caso armado.
  const wbY = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  wbY.getWorksheet("SALDOS").getCell(40, 3).value = null;      // se le saca la cuenta a Zento
  const cfgY = derivarConfigBalance(wbY, "pesos", derivarMapeoMaestro(wbY, "pesos"),
    lineasDeNota4(wbY), PARAMS.pesos, { filasQueAgrega });
  check(pccPendientesHtml(cfgY).includes("Zento"),
    "si una línea se queda sin cuenta, el panel la lista");

  const cat154 = categorias.find(c => c.filaMadre === 154);
  const miembroHtml = gcMiembroHtml(cat154, cat154.miembros[0]);
  check(miembroHtml.includes(cat154.miembros[0].codigo), "una cuenta miembro muestra su código en el HTML");
  check(miembroHtml.includes('data-accion="quitar"'), "una cuenta que se puede sacar tiene el botón Quitar");

  const cat200 = categorias.find(c => c.filaMadre === 200);
  const compartida = cat200.miembros.find(m => m.esFilaCompartida);
  const miembroCompartidoHtml = gcMiembroHtml(cat200, compartida);
  check(!miembroCompartidoHtml.includes('data-accion="quitar"'),
    "una cuenta que comparte fila con su categoría NO tiene botón Quitar");

  const catHtml = gcCategoriaHtml(cat154);
  check(catHtml.includes(cat154.nombre) && catHtml.includes(String(cat154.miembros.length)),
    "el HTML de una categoría muestra su nombre y la cantidad de cuentas");

  check(gcEsc('<script>&"</script>') === "&lt;script&gt;&amp;&quot;&lt;/script&gt;",
    "gcEsc escapa HTML (nombres de cuenta con caracteres raros no rompen el panel)");

  console.log(fallos ? `\n${fallos} falla(s).` : "\nTodo OK.");
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error("ERROR:", e); process.exit(1); });
