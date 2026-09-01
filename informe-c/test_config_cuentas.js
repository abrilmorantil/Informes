// El panel de configuración de cuentas (panel_config_cuentas.js) sobre el maestro real de
// pesos: confirma que arma el resumen y los pendientes (config_balances.js) — el pendiente
// conocido de la fila 137 ("Patentes a pagar" vs "Retenciones SUSS a pagar") — y que renderiza
// una categoría y sus cuentas sin tirar excepción. La lógica de sacar/agregar/editar cuentas
// se prueba aparte, en test_gestion_categorias.js.
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

  check(cfg.lineas.length === 444, `444 líneas en el balance de pesos (dio ${cfg.lineas.length})`);
  check(cfg.resumen.conMadre === 3, `3 casos de madre/hija resueltos solos (dio ${cfg.resumen.conMadre})`);

  const conflicto137 = cfg.avisos.find(a => a.tipo === "dos_cuentas_en_la_misma_fila" && a.usa.code === "213010010");
  check(!!conflicto137, 'la fila 137 queda marcada como pendiente ("213010010")');
  check(conflicto137 && conflicto137.tambien.code === "212020002",
    "el conflicto de la fila 137 muestra la otra cuenta (212020002 Retenciones SUSS a pagar)");

  const resumenHtml = pccResumenHtml(cfg, cfg.avisos.length, categorias.length);
  check(resumenHtml.includes(String(cfg.resumen.lineas)), "el resumen HTML muestra el total de líneas");
  check(resumenHtml.includes(String(cfg.avisos.length)), "el resumen HTML muestra la cantidad de pendientes");
  check(resumenHtml.includes(String(categorias.length)), "el resumen HTML muestra la cantidad de categorías");

  const pendientesHtml = pccPendientesHtml(cfg);
  check(pendientesHtml.includes("213010010") && pendientesHtml.includes("212020002"),
    "el panel de pendientes muestra las dos cuentas en conflicto de la fila 137");
  check(pendientesHtml.includes("Zento S.A."), 'el panel de pendientes lista la línea de Nota 4 "Zento S.A." sin cuenta');

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
