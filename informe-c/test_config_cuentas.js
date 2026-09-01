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
const { categoriasPesos, quitarCuentaDeCategoria, editarCuenta, agregarCuentaACategoria,
        corregirColumnaDeCuenta } = require(path.join(AQUI, "gestion_categorias.js"));
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
const { pccResumenHtml, pccPendientesHtml, pccSinCuentaHtml, pccCuentasDeHoja1,
        gcMiembroHtml, gcCategoriaHtml, gcEsc } = require(path.join(AQUI, "panel_config_cuentas.js"));

// Los pendientes son los avisos MENOS las líneas sin cuenta asignada, que no son un error
// sino un estado válido y tienen su propio bloque. Es la misma cuenta que hace el panel.
const pendientesDe = (cfg) => cfg.avisos.filter(a => a.tipo !== "linea_sin_cuenta");
const sinCuentaDe = (cfg) => cfg.avisos.filter(a => a.tipo === "linea_sin_cuenta");

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

  // El encabezado muestra SOLO los pendientes: los otros contadores no pedían ninguna acción
  // y diluían al único que sí.
  const resumenHtml = pccResumenHtml(cfg, pendientesDe(cfg).length, categorias.length);
  check(resumenHtml.includes("Pendientes de revisar") && resumenHtml.includes(">0<"),
    "el encabezado muestra los pendientes, y hoy son 0");
  check(resumenHtml.includes("check ok") && !resumenHtml.includes("check bad"),
    "y en verde, porque no hay ninguno");
  check(!resumenHtml.includes(String(categorias.length)) || categorias.length === 0,
    "ya no muestra el resto de los contadores");

  const conPendientes = pccResumenHtml(cfg, 3, categorias.length);
  check(conPendientes.includes("check bad") && conPendientes.includes(">3<"),
    "y si hubiera pendientes, los muestra en rojo");

  const pendientesHtml = pccPendientesHtml(cfgX);
  check(pendientesHtml.includes("213010010") && pendientesHtml.includes("212020002"),
    "el panel dibuja las dos cuentas en conflicto");
  // El maestro no tiene que tener pendientes. Si aparece alguno, se lista para poder verlo.
  const pend = pendientesDe(cfg);
  check(pend.length === 0,
    pend.length
      ? `quedan ${pend.length} pendiente(s): ` +
        pend.map(a => `${a.tipo} (${a.texto || "fila " + a.fila})`).join(", ")
      : "el maestro no tiene pendientes: cada línea del balance sabe de qué cuenta sale");

  // --- las líneas SIN cuenta asignada: no son pendientes, y se les puede asignar una
  console.log("\n=== líneas sin cuenta asignada ===");
  const sinCuenta = sinCuentaDe(cfg);
  check(sinCuenta.length > 0 && sinCuenta.some(a => /Socios/i.test(a.texto || "")),
    `"- Socios" figura como línea sin cuenta asignada (hay ${sinCuenta.length})`);
  check(!pccPendientesHtml(cfg).includes("Socios"),
    "y NO sale entre los pendientes: no es un error");

  const onvio = pccCuentasDeHoja1(wb, mapeo);
  check(onvio.length > 0 && onvio.every(c => /^\d{5,}$/.test(c.codigo) && c.nombre),
    `las cuentas sugeridas salen de Hoja1 (${onvio.length} libres)`);
  // Ofrecer una cuenta que ya alimenta otra línea la contaría dos veces y abriría el balance.
  check(!onvio.some(c => mapeo.cuentas[c.codigo]),
    "y deja afuera las que ya están puestas en otra línea del balance");
  check(pccCuentasDeHoja1(wb, mapeo).length < pccCuentasDeHoja1(wb, null).length,
    "sin ese filtro la lista sería más larga: el filtro hace algo");
  const scHtml = pccSinCuentaHtml(cfg, onvio);
  check(scHtml.includes("Socios") && scHtml.includes("cc_cod_") && scHtml.includes("cc_nom_"),
    "el bloque la muestra con los campos para asignarle una cuenta");
  check(scHtml.includes("<datalist") && scHtml.includes(onvio[0].codigo),
    "con las cuentas libres de Hoja1 como sugerencia");
  // El código se escribe: un desplegable cerrado con Hoja1 no alcanza, porque Hoja1 no es el
  // plan de Onvio sino sólo las cuentas que el balance ya recibe.
  check(/<input[^>]+id="cc_cod_/.test(scHtml),
    "y el código se puede escribir, no sólo elegir de la lista");
  check(!scHtml.includes("status-msg bad"),
    "el bloque no va en rojo: es un estado válido, no un problema");
  check(pccSinCuentaHtml({ avisos: [] }, onvio) === "",
    "si no hay ninguna línea sin cuenta, el bloque no se dibuja");

  // Asignarle una cuenta es la misma operación que corregir: se escribe en la columna que el
  // VLOOKUP de esa fila usa de clave, y la línea deja de figurar.
  const filaSocios = sinCuenta.find(a => /Socios/i.test(a.texto || "")).saldos;
  const wbZ = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  const elegida = pccCuentasDeHoja1(wbZ, derivarMapeoMaestro(wbZ, "pesos"))[0];
  corregirColumnaDeCuenta(wbZ, filaSocios, elegida.codigo, elegida.nombre);
  const cfgZ = derivarConfigBalance(wbZ, "pesos", derivarMapeoMaestro(wbZ, "pesos"),
    lineasDeNota4(wbZ), PARAMS.pesos, { filasQueAgrega });
  check(!sinCuentaDe(cfgZ).some(a => a.saldos === filaSocios),
    `al asignarle ${elegida.codigo}, la línea deja de figurar como sin cuenta`);

  // Y que el panel sabe dibujar una línea sin cuenta, con un caso armado.
  const wbY = await abrirWorkbook(fs.readFileSync(path.join(AQUI, "base_pesos.xlsx")));
  wbY.getWorksheet("SALDOS").getCell(40, 3).value = null;      // se le saca la cuenta a Zento
  const cfgY = derivarConfigBalance(wbY, "pesos", derivarMapeoMaestro(wbY, "pesos"),
    lineasDeNota4(wbY), PARAMS.pesos, { filasQueAgrega });
  check(pccSinCuentaHtml(cfgY, onvio).includes("Zento") || pccPendientesHtml(cfgY).includes("Zento"),
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
