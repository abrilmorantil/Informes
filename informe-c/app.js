// Une todo para el Informe C: estado en GitHub, alta inicial del maestro de pesos,
// y el ciclo de carga (export → destinos de cuentas nuevas → borrador → aprobar).
//
// Por ahora corre solo PESOS. Dólares queda para cuando estén validadas las
// equivalencias de códigos viejos (ver informe_balances/equivalencias_dolares_propuestas.json).

const $ = (id) => document.getElementById(id);
const mostrar = (id, v) => $(id).classList.toggle("hidden", !v);

let estado = null;
let bufferBase = null;

let cuentasExport = null;
let nuevasPendientes = [];   // las que necesitan destino (no proveedores)
let wbBorrador = null;
let resumenBorrador = null;
let listasDestino = null;    // {conceptos, lineasNota4}
let logLineas = [];
const log = (m) => { logLineas.push(String(m)); $("logBody").textContent = logLineas.join("\n"); };

let altaBuffer = null;
let clasificacion = null;   // la del Informe B, si esta disponible
let bufferBaseUsd = null;   // el maestro de dolares, si ya esta cargado
let equivalencias = {};     // decisiones de dolares ya confirmadas
let pendientesUsd = [];     // cuentas de dolares esperando destino
let cuentasMaestroUsd = [];
let wbBorradorUsd = null;
let altaBufferUsd = null;
let totalesEERR = null;        // los totales del Estado de Resultados de esta corrida
let avisosEERR = [];
let movimientoEERR = null;     // el balance de comprobación en dólares de esta corrida

// --------------------------------------------------------------- arranque

async function arrancar() {
  mostrar("cargando", true);
  try {
    if (!hasGhcSettings()) {
      mostrar("cargando", false);
      mostrar("cardSinConfig", true);
      return;
    }
    const e = await ghcLeerEstado();
    if (!e) {
      mostrar("cargando", false);
      mostrar("cardAlta", true);
      return;
    }
    estado = e.estado;
    // la clasificación de cuentas ya configurada en el Informe B (mismo repositorio)
    try {
      const mb = await ghcLeerMappingB();
      if (mb) clasificacion = indexarClasificacion(mb);
    } catch (err) { clasificacion = null; }
    const b = await ghcLeerBase();
    if (!b) throw new Error("Hay un estado guardado pero falta base_pesos.xlsx en el repositorio.");
    bufferBase = b.buffer;
    const bu = await ghcLeerBaseUsd();
    bufferBaseUsd = bu ? bu.buffer : null;
    equivalencias = await ghcLeerEquivalencias();
    mostrar("cardAltaUsd", !bufferBaseUsd);

    const h = estado.historial || [];
    $("txtUltimaCarga").textContent = h.length
      ? `Última carga: ${new Date(h[h.length - 1].fecha).toLocaleDateString("es-AR")} (${h[h.length - 1].cuentas} cuentas)`
      : "Todavía no se cargó ningún export.";

    mostrar("cargando", false);
    mostrar("cardEerrAnterior", true);
    renderEerrAnterior();
    mostrar("cardExport", true);
    await proponerPeriodo();
    mostrar("cardGuardados", true);
    renderHistorial();
  } catch (err) {
    mostrar("cargando", false);
    mostrar("cardSinConfig", true);
    $("cardSinConfig").innerHTML =
      `<h2>No pude leer el estado guardado</h2>
       <p class="footer-note">${err.message}</p>
       <div style="margin-top:14px;"><button onclick="abrirConfig()">Revisar configuración</button></div>`;
  }
}

function renderHistorial() {
  const h = (estado && estado.historial) || [];
  const ultimo = (estado && estado.ultimo_periodo_cerrado) ||
    (h.length ? h[h.length - 1].periodo : null);
  if (ultimo) {
    const p = fpPartes(ultimo);
    $("txtUltimoCierre").innerHTML = `Último período cerrado: <b>${p ? fpDescribir(p) : ultimo}</b>.`;
  }
  if (!h.length) return;
  // Las cargas viejas no guardaban el período: se muestran con guión en vez de inventarlo.
  $("historialBody").innerHTML = h.slice().reverse().map(x => `
    <tr>
      <td>${x.periodo || "—"}</td>
      <td>${x.fecha ? new Date(x.fecha).toLocaleDateString("es-AR") : "—"}</td>
      <td class="num">${x.cuentas ?? "—"}</td>
      <td class="num">${x.nuevas ?? 0}</td>
      <td class="num">${x.totalPesos != null ? x.totalPesos.toFixed(2) : "—"}</td>
    </tr>`).join("");
  mostrar("cardHistorial", true);
}

// ------------------------------------------------------- volver a bajar lo guardado

// Los maestros viven en el repositorio y no en la máquina de la usuaria, así que sin esto la
// única forma de tener el archivo bueno a mano era volver a correr el informe entero.
function bajarBuffer(buffer, nombre) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return nombre;
}

// El nombre lleva el período del último cierre, para que no queden dos "base_pesos.xlsx" en
// la carpeta de descargas sin manera de distinguirlos.
function sufijoDelUltimoCierre() {
  const h = (estado && estado.historial) || [];
  const ultimo = (estado && estado.ultimo_periodo_cerrado) ||
    (h.length ? h[h.length - 1].periodo : null);
  const p = ultimo ? fpPartes(ultimo) : null;
  return p ? `${p.anio}-${String(p.mes).padStart(2, "0")}` : "guardado";
}

async function bajarGuardado(cual) {
  const st = $("guardadosStatus");
  st.innerHTML = '<div class="status-msg">Bajando del repositorio…</div>';
  try {
    const suf = sufijoDelUltimoCierre();
    if (cual === "pesos") {
      const b = await ghcLeerBase();
      if (!b) throw new Error("No hay un maestro de pesos guardado.");
      st.innerHTML = `<div class="status-msg ok">Bajó <b>` +
        `${bajarBuffer(b.buffer, `SCA_Balance_${suf}_pesos.xlsx`)}</b>.</div>`;
    } else {
      const b = await ghcLeerBaseUsd();
      if (!b) throw new Error("No hay un maestro de dólares guardado.");
      st.innerHTML = `<div class="status-msg ok">Bajó <b>` +
        `${bajarBuffer(b.buffer, `SCA_Balance_${suf}_dolares.xlsx`)}</b>.</div>`;
    }
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

if ($("btnBajarPesos")) $("btnBajarPesos").addEventListener("click", () => bajarGuardado("pesos"));
if ($("btnBajarDolares")) $("btnBajarDolares").addEventListener("click", () => bajarGuardado("dolares"));

// --------------------------------------------------------------- configuración

function abrirConfig() {
  const s = loadGhcSettings();
  $("cfgToken").value = s.token || "";
  $("cfgRepo").value = s.repo || "";
  $("cfgRama").value = s.rama || "main";
  $("cfgCarpeta").value = s.carpeta || "informe-c";
  $("configStatus").innerHTML = "";
  mostrar("modalConfig", true);
}
function cerrarConfig() { mostrar("modalConfig", false); }

async function guardarConfig() {
  saveGhcSettings({
    token: $("cfgToken").value.trim(),
    repo: $("cfgRepo").value.trim(),
    rama: $("cfgRama").value.trim() || "main",
    carpeta: $("cfgCarpeta").value.trim() || "informe-c",
  });
  const st = $("configStatus");
  st.innerHTML = '<div class="status-msg">Probando la conexión…</div>';
  try {
    await ghcLeerEstado();
    st.innerHTML = '<div class="status-msg ok">Conectado correctamente.</div>';
    setTimeout(() => { cerrarConfig(); location.reload(); }, 800);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

// --------------------------------------------------------------- alta inicial

$("fileBase").addEventListener("change", async () => {
  const f = $("fileBase").files[0];
  if (!f) return;
  $("txtBase").textContent = f.name;
  $("altaStatus").innerHTML = "";
  try {
    if (/\.xls$/i.test(f.name)) {
      throw new Error(
        "Este es el .xls viejo. Abrilo en Excel y guardalo como 'Libro de Excel (.xlsx)' " +
        "primero: el formato viejo no conserva las fórmulas al procesarlo acá."
      );
    }
    altaBuffer = await f.arrayBuffer();
    const wb = await abrirWorkbook(altaBuffer.slice(0));
    for (const hoja of [HOJA_DISTRIB, HOJA_SUMAS, "Anexo II", "Activo y Pasivo", "Balance"]) {
      if (!wb.getWorksheet(hoja)) throw new Error(`El archivo no tiene la hoja '${hoja}'. ¿Es el balance formal en pesos?`);
    }
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const n = Object.keys(mapeo.cuentas).length;
    $("altaDeteccion").innerHTML = `
      <div class="check ok" style="display:block;">
        <div class="name">Maestro de pesos reconocido: ${n} cuentas en Distribución por línea</div>
        <div class="detail">${mapeo.duplicadas.length
          ? `Hay ${mapeo.duplicadas.length} código(s) que figuran dos veces (${mapeo.duplicadas.map(d => d.codigo).join(", ")}). Se usa la primera de cada uno. Solo importa si alguno llega a venir en el export; mientras tanto son restos del plan de cuentas viejo.`
          : "Sin cuentas repetidas."}</div>
      </div>`;
    mostrar("altaDeteccion", true);
    $("btnGuardarAlta").disabled = false;
  } catch (e) {
    $("altaStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnGuardarAlta").disabled = true;
  }
});

$("btnGuardarAlta").addEventListener("click", async () => {
  const st = $("altaStatus");
  mostrar("spinnerAlta", true);
  $("btnGuardarAlta").disabled = true;
  try {
    await ghcGuardarTodo({
      bufferBase: altaBuffer,
      estado: { historial: [] },
      mensaje: "Balance Pesos: carga inicial del maestro",
    });
    st.innerHTML = '<div class="status-msg ok">Guardado. Ya podés cargar el export.</div>';
    setTimeout(() => location.reload(), 900);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnGuardarAlta").disabled = false;
  } finally {
    mostrar("spinnerAlta", false);
  }
});

$("fileBaseUsd").addEventListener("change", async () => {
  const f = $("fileBaseUsd").files[0];
  if (!f) return;
  $("txtBaseUsd").textContent = f.name;
  $("altaUsdStatus").innerHTML = "";
  try {
    if (/\.xls$/i.test(f.name)) {
      throw new Error("Es el .xls viejo. Abrilo en Excel y guardalo como 'Libro de Excel (.xlsx)' primero.");
    }
    altaBufferUsd = await f.arrayBuffer();
    const wb = await abrirWorkbook(altaBufferUsd.slice(0));
    for (const hoja of [HOJA_DISTRIB, HOJA_SUMAS, "Balance"]) {
      if (!wb.getWorksheet(hoja)) {
        throw new Error("El archivo no tiene la hoja '" + hoja + "'. Es el balance formal en dolares?");
      }
    }
    const cm = cuentasDelMaestro(wb, "dolares");
    $("altaUsdStatus").innerHTML =
      '<div class="status-msg ok">Maestro de dolares reconocido: ' + cm.length + ' cuentas en Distribución por línea.</div>';
    $("btnGuardarAltaUsd").disabled = false;
  } catch (e) {
    $("altaUsdStatus").innerHTML = '<div class="status-msg bad">' + e.message + '</div>';
    $("btnGuardarAltaUsd").disabled = true;
  }
});

$("btnGuardarAltaUsd").addEventListener("click", async () => {
  $("btnGuardarAltaUsd").disabled = true;
  try {
    await ghcGuardarBaseUsd(altaBufferUsd, "Balance Dolares: carga inicial del maestro");
    $("altaUsdStatus").innerHTML = '<div class="status-msg ok">Guardado.</div>';
    setTimeout(() => location.reload(), 900);
  } catch (e) {
    $("altaUsdStatus").innerHTML = '<div class="status-msg bad">' + e.message + '</div>';
    $("btnGuardarAltaUsd").disabled = false;
  }
});

// --------------------------------------------------------------- ciclo de carga

$("fileExport").addEventListener("change", () => {
  const f = $("fileExport").files[0];
  if (f) $("txtExport").textContent = f.name;
  $("btnProcesar").disabled = !f;
});

async function leerExport() {
  const buf = await $("fileExport").files[0].arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return parseExportBalances(
    XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }),
    ws["!merges"]
  );
}

$("btnProcesar").addEventListener("click", async () => {
  ["cardNuevas", "cardResultado", "cardCierre"].forEach(id => mostrar(id, false));
  $("exportStatus").innerHTML = "";
  mostrar("spinner", true);
  $("btnProcesar").disabled = true;
  try {
    const exp = await leerExport();
    cuentasExport = exp.cuentas;
    if (exp.discrepanciasCapitulo.length) {
      $("exportStatus").innerHTML =
        `<div class="status-msg bad">Ojo: estas cuentas están en una sección que no coincide con su código:
         ${exp.discrepanciasCapitulo.map(d => d.codigo).join(", ")}. Se usó la sección del export.</div>`;
    }

    // para saber qué es nuevo hay que mirar el maestro
    const wb = await abrirWorkbook(bufferBase.slice(0));
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const nuevas = detectarNuevas(cuentasExport, mapeo, "pesos");
    listasDestino = { madres: madresResultados(wb, "pesos"), lineasNota4: lineasDeNota4(wb) };

    // no se pregunta ni por los proveedores ni por lo que el Informe B ya tiene clasificado
    const provNuevos = nuevas.filter(n => esProveedor(clasificacion, n.codigo));
    const resueltas = [];
    nuevasPendientes = [];
    for (const n of nuevas) {
      if (esProveedor(clasificacion, n.codigo)) continue;
      const m = n.capitulo === "RESULTADOS"
        ? madreEnArchivo(clasificacion, listasDestino.madres, n.codigo) : null;
      if (m) resueltas.push(`<b>${n.codigo}</b> ${n.nombre} → ${m.nombre}`);
      else nuevasPendientes.push(n);
    }

    if (provNuevos.length) {
      $("exportStatus").innerHTML +=
        `<div class="status-msg ok">${provNuevos.length} proveedor(es) nuevo(s) se insertan solos en el detalle del pasivo.</div>`;
    }
    if (resueltas.length) {
      $("exportStatus").innerHTML +=
        `<div class="status-msg ok">${resueltas.length} cuenta(s) nueva(s) ya estaban clasificadas en BALCOMPROBDOLARES y se ubican solas: ${resueltas.join(", ")}.</div>`;
    }
    if (!clasificacion) {
      $("exportStatus").innerHTML +=
        `<div class="status-msg">No pude leer la clasificación de BALCOMPROBDOLARES (mapping.json), así que voy a preguntar por cada cuenta nueva.</div>`;
    }

    if (nuevasPendientes.length) {
      renderNuevas();
      mostrar("cardNuevas", true);
    } else {
      await correrMotor({});
    }
  } catch (e) {
    $("exportStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    mostrar("spinner", false);
    $("btnProcesar").disabled = false;
  }
});

function renderNuevas() {
  $("nuevasBody").innerHTML = nuevasPendientes.map((n, i) => {
    const esResultado = n.capitulo === "RESULTADOS";
    const destino = esResultado
      ? `<select class="selMadre" data-idx="${i}">
           <option value="">— ¿de qué cuenta madre es? —</option>
           ${listasDestino.madres.map(m => `<option value="${m.fila}">${m.codigo} - ${m.nombre}</option>`).join("")}
         </select>`
      : `<select class="selLinea" data-idx="${i}">
           <option value="">— debajo de qué línea de la Nota 4 —</option>
           ${listasDestino.lineasNota4.map((l, j) => `<option value="${j}">${l.texto} (fila ${l.fila})</option>`).join("")}
         </select>`;
    return `
      <tr class="pending-row">
        <td>${n.codigo}</td>
        <td>${n.nombre}</td>
        <td>${n.capitulo}</td>
        <td class="num">${n.saldo_ars.toFixed(2)}</td>
        <td>${destino}</td>
      </tr>`;
  }).join("");
  conBuscadorTodos(".selMadre", "Buscar cuenta madre…");
  conBuscadorTodos(".selLinea", "Buscar línea de la Nota 4…");
}

$("btnConfirmarNuevas").addEventListener("click", async () => {
  const destinos = {};
  for (let i = 0; i < nuevasPendientes.length; i++) {
    const n = nuevasPendientes[i];
    if (n.capitulo === "RESULTADOS") {
      const sc = document.querySelector(`.selMadre[data-idx="${i}"]`);
      if (!sc.value) {
        $("nuevasStatus").innerHTML = '<div class="status-msg bad">Falta elegir la cuenta madre de al menos una cuenta.</div>';
        return;
      }
      destinos[n.codigo] = { madreFila: parseInt(sc.value, 10) };
    } else {
      const sl = document.querySelector(`.selLinea[data-idx="${i}"]`);
      if (!sl.value) {
        $("nuevasStatus").innerHTML = '<div class="status-msg bad">Falta elegir la línea de la Nota 4 de al menos una cuenta.</div>';
        return;
      }
      destinos[n.codigo] = { lineaModelo: listasDestino.lineasNota4[parseInt(sl.value, 10)] };
    }
  }
  $("nuevasStatus").innerHTML = "";
  $("btnConfirmarNuevas").disabled = true;
  try {
    await correrMotor(destinos);
    mostrar("cardNuevas", false);
  } catch (e) {
    $("nuevasStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    $("btnConfirmarNuevas").disabled = false;
  }
});

async function correrMotor(destinos) {
  // el período se confirma antes de generar nada: de él salen las fechas de adentro de los
  // informes y el nombre de los archivos
  periodoConfirmado = fpPartes(($("periodoBalance").value || "").trim());
  if (!periodoConfirmado) {
    throw new Error("Falta confirmar el período que se emite, con el formato 2026-07-31. " +
                    "De ahí salen las fechas de los balances y el nombre de los archivos.");
  }

  logLineas = [];
  // siempre desde una copia limpia del maestro guardado
  const wb = await abrirWorkbook(bufferBase.slice(0));
  log("PESOS");
  const { resumen } = procesarBalance({
    wb, cuentasExport, moneda: "pesos", destinosElegidos: destinos, clasificacion, log,
  });
  ponerFechasDelPeriodo(wb, "pesos", periodoConfirmado);
  arrastrarCapital(wb, "pesos", periodoConfirmado);
  wbBorrador = wb;
  resumenBorrador = resumen;

  // Dolares, si el maestro esta cargado. Si faltan decisiones se muestra la pantalla de
  // revision, y el balance en pesos igual queda listo para descargar.
  wbBorradorUsd = null;
  totalesEERR = null;          // que no quede el de la corrida anterior
  avisosEERR = [];
  let resumenUsd = null;
  if (bufferBaseUsd) {
    log("\nDOLARES");
    const wbu = await abrirWorkbook(bufferBaseUsd.slice(0));
    try {
      const r = procesarDolares({ wb: wbu, cuentasExport, clasificacion, equivalencias, log });
      ponerFechasDelPeriodo(wbu, "dólares", periodoConfirmado);
      arrastrarCapital(wbu, "dólares", periodoConfirmado);
      wbBorradorUsd = wbu;
      resumenUsd = r.resumen;
      calcularEERR(wbu, r);
      mostrar("cardRevisionUsd", false);
    } catch (e) {
      if (!e.pendientesDolares) throw e;
      pendientesUsd = e.pendientesDolares;
      cuentasMaestroUsd = e.cuentasMaestro;
      renderRevisionUsd();
      mostrar("cardRevisionUsd", true);
      log("  Faltan ubicar " + pendientesUsd.length + " cuenta(s) en dolares: ver el paso de revision.");
      // sin el balance en dolares resuelto no hay EE RR: sus cifras salen de ahí
      avisosEERR = [`primero hay que resolver las ${pendientesUsd.length} cuenta(s) ` +
                    `pendientes del balance en dólares, de ahí salen sus cifras.`];
    }
  }

  revisarCapital();
  renderQueSeDescarga();
  renderResultado(resumen, resumenUsd);
  mostrar("cardResultado", true);
  mostrar("cardCierre", true);
}

// --------------------------------------------------------------- revision de dolares
//
// Se pregunta UNA sola vez por cuenta: la respuesta queda guardada en GitHub y las
// corridas siguientes ya no la vuelven a pedir. No se pre-selecciona ningun destino a
// proposito: los emparejamientos automaticos por parecido de nombre fallan la mitad de
// las veces (proponen meter la amortizacion acumulada dentro del bien de uso).

function renderRevisionUsd() {
  // Donde el código del plan viejo Y el nombre coinciden, se deja la propuesta ya
  // elegida: son justo las que la usuaria igual iba a buscar a mano. No se aplica sola
  // —queda a la vista para que la confirme— y donde no hay propuesta no se elige nada.
  let conPropuesta = 0;
  $("revisionUsdBody").innerHTML = pendientesUsd.map((c, i) => {
    const prop = proponerDestino(c, cuentasMaestroUsd);
    if (prop) conPropuesta++;
    const opciones = cuentasMaestroUsd.map(f =>
      '<option value="' + f.codigo + '"' + (prop && f.codigo === prop.codigo ? ' selected' : '') +
      '>' + f.codigo + ' - ' + f.nombre + '</option>').join("");
    return '<tr class="pending-row">' +
      '<td>' + c.codigo + '</td>' +
      '<td>' + c.nombre + '</td>' +
      '<td class="num">' + c.saldo_usd.toFixed(2) + '</td>' +
      '<td><select class="selUsd" data-idx="' + i + '">' +
        '<option value=""' + (prop ? '' : ' selected') + '>— elegí a qué línea del balance va —</option>' +
        '<option value="__nada__">No cargarla en el balance en dólares</option>' +
        opciones +
      '</select>' +
      (prop ? '<div class="buscador-aviso">Propuesta: mismo código en el plan viejo (' +
              prop.codigo + ') y nombre parecido. Revisala.</div>' : '') +
      '</td>' +
    '</tr>';
  }).join("");
  $("revisionUsdResumen").innerHTML = conPropuesta
    ? '<div class="status-msg ok">' + conPropuesta + ' de ' + pendientesUsd.length +
      ' vienen con una propuesta ya elegida, porque el código del plan viejo y el nombre ' +
      'coinciden. Revisalas igual. Las otras ' + (pendientesUsd.length - conPropuesta) +
      ' no tienen equivalente claro: esas las elegís vos.</div>'
    : "";
  // acá son 203 opciones: el buscador es imprescindible
  conBuscadorTodos(".selUsd", "Buscar cuenta del balance en dólares…");
}

$("btnConfirmarUsd").addEventListener("click", async () => {
  const selects = document.querySelectorAll(".selUsd");
  const nuevas = {};
  for (let i = 0; i < selects.length; i++) {
    if (!selects[i].value) {
      $("revisionUsdStatus").innerHTML =
        '<div class="status-msg bad">Falta decidir al menos una cuenta. Si alguna no va al balance, eleg&iacute; "No cargarla".</div>';
      return;
    }
    nuevas[pendientesUsd[i].codigo] = selects[i].value === "__nada__" ? null : selects[i].value;
  }
  $("revisionUsdStatus").innerHTML = "";
  $("btnConfirmarUsd").disabled = true;
  try {
    equivalencias = Object.assign({}, equivalencias, nuevas);
    await ghcGuardarEquivalencias(equivalencias, "Balance Dolares: equivalencias confirmadas");
    const wbu = await abrirWorkbook(bufferBaseUsd.slice(0));
    const r = procesarDolares({ wb: wbu, cuentasExport, clasificacion, equivalencias, log });
    wbBorradorUsd = wbu;
    calcularEERR(wbu, r);
    mostrar("cardRevisionUsd", false);
    revisarCapital();
  renderQueSeDescarga();
    renderResultado(resumenBorrador, r.resumen);
    $("cierreStatus").innerHTML =
      '<div class="status-msg ok">Decisiones guardadas. No te las vuelve a preguntar.</div>';
  } catch (e) {
    $("revisionUsdStatus").innerHTML = '<div class="status-msg bad">' + e.message + '</div>';
  } finally {
    $("btnConfirmarUsd").disabled = false;
  }
});

function renderResultado(r, rUsd) {
  // Cada línea dice de qué balance habla: los dos se procesan en la misma corrida y
  // sus hojas se llaman igual, así que un aviso sin moneda manda a mirar el archivo
  // equivocado (las filas que nombra no existen en el otro).
  const items = [
    { name: "Pesos: total del export", value: r.total.toFixed(2), ok: true,
      detail: "Excel recalcula los estados al abrir el archivo; el control L23 del Balance tiene que dar 0." },
    { name: "Pesos: cuentas cargadas en Balance de sumas y saldos", value: String(r.cuentas), ok: true,
      detail: "La zona de pegado quedó actualizada con el export fresco." },
    { name: "Pesos: cuentas nuevas insertadas", value: String(r.nuevas), ok: true,
      detail: "Con su fila en Distribución por línea y su referencia en los estados." },
    { name: "Pesos: cuentas sin enganchar", value: String(r.noEnganchadas.length), ok: r.noEnganchadas.length === 0,
      detail: r.noEnganchadas.length ? r.noEnganchadas.join(", ") : "Todas las cuentas del export quedaron enganchadas a Distribución por línea." },
  ];
  $("checksBody").innerHTML = items.map(c => `
    <div class="check ${c.ok ? "ok" : "bad"}">
      <div><div class="name">${c.name}</div><div class="detail">${c.detail}</div></div>
      <div style="display:flex; align-items:center; gap:10px;">
        <span class="value">${c.value}</span>
        <span class="badge ${c.ok ? "ok" : "bad"}">${c.ok ? "OK" : "Revisar"}</span>
      </div>
    </div>`).join("");

  if (rUsd) {
    items.push(
      { name: "Dolares: total del export", value: rUsd.total.toFixed(2), ok: true,
        detail: "Es la cifra que tiene que dar el balance en dolares." },
      { name: "Dolares: lineas con importe", value: String(rUsd.lineasConImporte), ok: true,
        detail: "Sobre " + rUsd.cuentasMaestro + " cuentas del maestro. Los proveedores van agrupados en una sola linea." });
    if (rUsd.ignoradas.length) {
      items.push({ name: "Dolares: cuentas dejadas afuera", value: String(rUsd.ignoradas.length),
        ok: true, detail: rUsd.ignoradas.join(", ") });
    }
  }

  const avisos = [];
  if (rUsd && rUsd.sinFilaEnHoja1 && rUsd.sinFilaEnHoja1.length) {
    avisos.push(
      '<div class="check bad" style="display:block;">' +
      '<div class="name">Dolares: ' + rUsd.sinFilaEnHoja1.length +
      ' cuenta(s) del balance no tienen fila donde pegar el importe</div>' +
      '<div class="detail">' + rUsd.sinFilaEnHoja1.join(", ") +
      '. Estan en Distribución por línea pero no en Balance de sumas y saldos, asi que su importe queda en cero.</div></div>');
  }
  if (r.duplicadas.length) {
    avisos.push(`
      <div class="check bad" style="display:block;">
        <div class="name">Pesos: ${r.duplicadas.length} código(s) usados por dos cuentas distintas, pendientes de resolver</div>
        <div class="detail">
          ${r.duplicadas.map(d => `<b>${d.codigo}</b> (filas ${d.filaPrevia} y ${d.fila})`).join(", ")}.
          Son filas del archivo <b>de pesos</b>, no del de dólares: los dos tienen una hoja
          llamada SALDOS y sus filas no coinciden. Mientras no se decida cuál es la correcta
          se usa la primera; hoy las dos están en cero, así que no altera ningún importe.
        </div>
      </div>`);
  }
  $("avisosBody").innerHTML = avisos.join("");
}

// --------------------------------------------------------- Estado de Resultados
//
// Los totales del EE RR son los mismos que las hojas Anexo II y Resultados exponen por
// fórmula. Como Excel todavía no las calculó, se reproducen acá a partir de los saldos
// que la corrida acaba de resolver.
function calcularEERR(wbUsd, resultadoDolares) {
  totalesEERR = null;
  avisosEERR = [];
  // el balance de comprobación en dólares que acompaña al estado de resultados
  movimientoEERR = (resultadoDolares && resultadoDolares.movimiento) || null;
  try {
    const porFila = new Map();
    for (const d of resultadoDolares.destinos.values()) {
      porFila.set(d.fila, d.aportes.reduce((a, x) => a + x.saldo, 0));
    }
    totalesEERR = totalesEstadoResultados(wbUsd, (f) => porFila.get(f) || 0);
    const v = verificarEERR(totalesEERR);
    avisosEERR = v.avisos;
    log("\nESTADO DE RESULTADOS");
    log(`  Gastos de operación ${totalesEERR.gastosOperacion.toFixed(2)} | ` +
        `administración ${totalesEERR.gastosAdministracion.toFixed(2)} | ` +
        `resultado del ejercicio ${totalesEERR.resultadoEjercicio.toFixed(2)}`);
    for (const t of (totalesEERR.tomadoDeOtrasHojas || [])) {
      log(`  ${t.hoja}!${t.celda} = ${t.valor.toFixed(2)} — es una hoja que la app no toca, ` +
          `se usa el valor que tiene el maestro.`);
    }
    avisosEERR.forEach(a => log(`  ⚠ ${a}`));
  } catch (e) {
    avisosEERR = [e.message];
    log(`  ⚠ No pude armar el Estado de Resultados: ${e.message}`);
  }
}

// ------------------------------------------------------------------ período que se emite

// Los maestros se reusan mes a mes: sus fechas son las del mes pasado hasta que alguien las
// cambia, y nadie avisa. Además el nombre del archivo salía con la fecha de HOY, que no tiene
// nada que ver con el período. Por eso el período se confirma siempre antes de generar, y de
// él salen tanto las fechas de adentro como el nombre de los archivos. Ver `fechas.js`.
let periodoConfirmado = null;

// El cierre que cada maestro declara hoy, leído de `Balance!F5`.
async function cierreDelMaestro(buffer) {
  if (!buffer) return null;
  const wb = await abrirWorkbook(buffer.slice(0));
  const ws = wb.getWorksheet("Balance");
  return ws ? fpCierreDeCelda(ws.getCell("F5")) : null;
}

// Se propone el mes siguiente al que el maestro trae. Es lo más fiel: el maestro sabe hasta
// dónde llegó. Si no se puede leer, se deja vacío y se pide a mano en vez de inventar.
async function proponerPeriodo() {
  const campo = $("periodoBalance");
  if (!campo) return;
  let propuesto = null;
  try {
    const cierre = await cierreDelMaestro(bufferBase);
    if (cierre) propuesto = fpMesSiguiente(cierre);
  } catch (e) { /* se pide a mano */ }
  if (propuesto) campo.value = fpISO(propuesto);
  campo.addEventListener("input", renderPeriodo);
  renderPeriodo();
}

// Sólo se dice algo cuando la fecha no sirve. Si es válida no hace falta explicar nada: el
// campo ya muestra el período, y contar ademas de dónde salió la propuesta y cómo se van a
// llamar los archivos era texto que se lee una vez y despues estorba todos los meses.
function renderPeriodo() {
  const campo = $("periodoBalance"), caja = $("periodoBalanceInfo");
  if (!campo || !caja) return;
  caja.innerHTML = fpPartes(campo.value.trim()) ? ""
    : `<span class="status-msg bad">Poné la fecha de cierre del período, con el formato ` +
      `<b>2026-07-31</b>.</span>`;
}

// Pone en el libro las fechas del período confirmado. Devuelve qué cambió, para el log.
function ponerFechasDelPeriodo(wb, moneda, nuevo) {
  const ws = wb.getWorksheet("Balance");
  const viejo = ws ? fpCierreDeCelda(ws.getCell("F5")) : null;
  if (!viejo) {
    log(`\n⚠ ${moneda.toUpperCase()}: no pude leer del archivo a qué fecha está cerrado, ` +
        `así que las fechas de adentro quedaron como estaban. Hay que revisarlas a mano.`);
    return null;
  }
  if (fpIguales(viejo, nuevo)) {
    log(`\n${moneda.toUpperCase()}: las fechas ya decían ${fpDescribir(nuevo)}, no hubo que tocarlas.`);
    return { cambios: [], otrasFechas: [] };
  }
  const r = fpReescribirLibro(wb, viejo, nuevo);
  const hojas = [...new Set(r.cambios.map(c => c.hoja))];
  log(`\n${moneda.toUpperCase()}: las fechas pasaron a ${fpDescribir(nuevo)} — ` +
      `${r.cambios.length} celda(s) en ${hojas.join(", ")}.`);
  return r;
}

// --------------------------------------------------------- arrastre del capital (Pat.Neto)

// El EEPN no se acumula solo: lo que quedó en "Saldos al <fecha>" tiene que abrir el mes
// siguiente, y la línea del aumento del mes anterior tiene que salir para que entre la nueva.
// Se hacía a mano todos los meses. Ver `pasarCierreAlInicio` en capital.js.
function arrastrarCapital(wb, moneda, periodo) {
  let r;
  try {
    r = pasarCierreAlInicio(wb, periodo);
  } catch (e) {
    log(`
⚠ ${moneda.toUpperCase()}: no pude arrastrar el capital de "Pat.Neto": ${e.message}`);
    return;
  }
  if (r.sinHoja) return;
  for (const x of (r.sinFecha || [])) {
    log(`
⚠ ${moneda.toUpperCase()}: la línea "${x.etiqueta}" (fila ${x.fila} de Pat.Neto) no ` +
        `tiene una fecha que pueda leer, así que la dejé donde estaba. Revisala.`);
  }
  if (!r.arrastradas.length) return;
  log(`
${moneda.toUpperCase()}: en "Pat.Neto" pasé el cierre del mes anterior a la apertura — ` +
      r.arrastradas.map(x => `${x.celda} de ${fmtImporte(x.antes)} a ${fmtImporte(x.despues)}`).join(", ") +
      `, y saqué ` + r.limpiadas.map(x => `"${x.etiqueta}"`).join(", ") +
      `. El capital declarado sigue en ${fmtImporte(r.apertura)}.`);
}

// --------------------------------------------------------- control del capital (Pat.Neto)

// `Pat.Neto` arma el capital con números escritos a mano, así que cuando entra un aporte el
// Activo lo refleja y el Patrimonio Neto no: el balance sale abierto exactamente por esa
// diferencia. Ver el encabezado de `capital.js`. Acá se controla antes de dejar descargar.
let capitalPendiente = [];

const fmtImporte = (n) => (typeof n === "number"
  ? n.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  : String(n));

function revisarCapital() {
  capitalPendiente = [];
  const libros = [["pesos", wbBorrador], ["dólares", wbBorradorUsd]];
  for (const [moneda, wb] of libros) {
    if (!wb) continue;
    let c;
    try { c = controlarCapital(wb); } catch (e) { c = { ok: false, motivo: e.message }; }
    if (c.ok || c.sinHoja) continue;
    capitalPendiente.push({ moneda, wb, control: c });
    log(`\n⚠ ${moneda.toUpperCase()}: al capital de "Pat.Neto" le faltan ` +
        (c.falta !== undefined
          ? `${fmtImporte(c.falta)} contra la cuenta Capital Suscripto.`
          : `datos: ${c.motivo}`));
  }
  renderCapital();
}

function renderCapital() {
  const caja = $("capitalBox");
  if (!caja) return;
  mostrar("cardCapital", capitalPendiente.length > 0);
  if (!capitalPendiente.length) { caja.innerHTML = ""; return; }

  const noResueltos = capitalPendiente.filter(p => p.control.falta === undefined);
  const filas = capitalPendiente.filter(p => p.control.falta !== undefined).map(p => `
    <tr>
      <td><b>${p.moneda}</b></td>
      <td class="num">${fmtImporte(p.control.declarado.total)}</td>
      <td class="num">${fmtImporte(p.control.contable.valor)}</td>
      <td class="num"><b>${fmtImporte(p.control.falta)}</b></td>
    </tr>`).join("");

  caja.innerHTML = `
    <div class="status-msg bad">La cuenta <b>Capital Suscripto</b> no coincide con lo que
      declara <b>Pat.Neto</b>, así que el balance saldría abierto por esa diferencia.</div>
    <table class="tabla">
      <thead><tr><th>Balance</th><th class="num">Pat.Neto</th>
        <th class="num">Cuenta</th><th class="num">Falta</th></tr></thead>
      <tbody>${filas}</tbody>
    </table>
    ${noResueltos.length ? noResueltos.map(p =>
      `<div class="status-msg bad">${p.moneda}: ${p.control.motivo}</div>`).join("") : ""}
    <p class="footer-note">Si es un aumento de capital, poné el texto con la fecha del aporte
      y se carga en cada balance con el importe de su moneda. Si no lo es, hay que revisar el
      maestro.</p>
    <label>Texto de la línea
      <input type="text" id="txtAumentoCapital" placeholder="Aumento de capital 31/07/2026"
             style="max-width:320px;">
    </label>
    <button id="btnCargarCapital" class="secundario">Cargar el aumento</button>
    <div id="capitalStatus"></div>`;

  $("btnCargarCapital").addEventListener("click", () => {
    const st = $("capitalStatus");
    const etiqueta = $("txtAumentoCapital").value.trim();
    if (!etiqueta) {
      st.innerHTML = `<div class="status-msg bad">Falta el texto de la línea.</div>`;
      return;
    }
    const hechos = [];
    try {
      for (const p of capitalPendiente) {
        if (p.control.falta === undefined) continue;
        const r = agregarAumentoDeCapital(p.wb, { etiqueta, importe: p.control.falta });
        hechos.push(`<b>${p.moneda}</b>: ${fmtImporte(r.importe)} en la fila ${r.fila} de Pat.Neto` +
          (r.totalesArreglados.length
            ? `, y se corrigió ${r.totalesArreglados.map(t => t.celda).join(", ")} para que la sume`
            : ""));
        log(`\n${p.moneda.toUpperCase()}: "${etiqueta}" por ${fmtImporte(r.importe)} ` +
            `cargado en Pat.Neto fila ${r.fila}. El capital quedó en ` +
            `${fmtImporte(r.control.declarado.total)}, igual que la cuenta.`);
        r.totalesArreglados.forEach(t =>
          log(`  La fila de cierre no sumaba esa fila: ${t.celda} pasó de ${t.antes} a ${t.despues}.`));
      }
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
      return;
    }
    revisarCapital();
    renderQueSeDescarga();
    const caja2 = $("capitalStatus");
    if (caja2) {
      caja2.innerHTML = `<div class="status-msg ok">Cargado: ${hechos.join("; ")}. ` +
        `Ya se puede descargar.</div>`;
    } else {
      $("descargaStatus").innerHTML =
        `<div class="status-msg ok">Cargado: ${hechos.join("; ")}. Ya se puede descargar.</div>`;
    }
  });
}

// --------------------------------------------------------------- descarga y cierre

async function descargar(wb, sufijo) {
  // La hoja de trazabilidad se arma recién acá, sobre el archivo terminado: es una foto de
  // ESTA corrida. Por eso no se guarda en el maestro, donde el mes que viene sería mentira.
  try {
    const moneda = sufijo === "dolares" ? "dolares" : "pesos";
    const r = tzEscribirHoja(wb, moneda);
    log(`\n${sufijo.toUpperCase()}: hoja "De dónde sale cada saldo" con ${r.filas} cuenta(s)` +
        (r.cortes ? `, ${r.cortes} marcada(s) porque el camino se corta.` : "."));
  } catch (e) {
    log(`\n⚠ ${sufijo.toUpperCase()}: no pude armar la hoja de trazabilidad (${e.message}). ` +
        `El balance se descarga igual.`);
  }
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  // el nombre sale del PERÍODO, no de la fecha de hoy: un balance de julio bajado en agosto
  // se llamaba "2026-08" y no había forma de distinguirlo del de agosto
  const p = periodoConfirmado;
  a.download = "SCA_Balance_" + p.anio + "-" + String(p.mes).padStart(2, "0") +
    "_" + sufijo + ".xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return a.download;
}

// Qué archivos van a bajar. Antes no se decía en ningún lado, así que si el EE RR no salía
// —queda en null cuando `calcularEERR` falla— no bajaba y no había manera de darse cuenta:
// el motivo iba sólo al log.
function textoDeDescarga({ hayPesos, hayDolares, hayEERR, avisos }) {
  if (capitalPendiente.length) {
    return `<div class="status-msg bad">No baja nada: el capital de "Pat.Neto" no coincide ` +
      `con la cuenta Capital Suscripto en ` +
      capitalPendiente.map(p => p.moneda).join(" y ") +
      `, así que el balance saldría abierto. Hay que resolverlo arriba.</div>`;
  }
  const items = [];
  if (hayPesos) items.push(["ok", "Balance en pesos"]);
  if (hayDolares) items.push(["ok", "Balance en dólares"]);
  if (hayEERR) {
    items.push(["ok", "Estado de Resultados (EE RR)"]);
  } else if (hayDolares) {
    items.push(["bad", "Estado de Resultados: NO se puede armar" +
      ((avisos && avisos.length) ? " — " + avisos.join(" ") : "")]);
  } else {
    items.push(["", "Estado de Resultados: hace falta el balance en dólares, " +
      "porque sus cifras salen de ahí." +
      ((avisos && avisos.length) ? " " + avisos.join(" ") : "")]);
  }
  return items.map(([c, t]) =>
    `<div class="status-msg ${c}">${c === "ok" ? "Baja: " : ""}${t}</div>`).join("");
}

// los tres archivos no se llaman parecido, así que quedan lejos entre sí en la carpeta de
// descargas: hay que decir el nombre de cada uno
function resumenDeDescarga(nombres, hayEERR) {
  return `<div class="status-msg ok">` +
    (nombres.length === 1 ? "Bajó 1 archivo:" : `Bajaron ${nombres.length} archivos:`) + `<br>` +
    nombres.map(n => "<b>" + n + "</b>").join("<br>") +
    (hayEERR ? "" : "<br>El EE RR no bajó, por lo que dice arriba.") +
    `<br>Si el navegador preguntó si permitís varias descargas, hay que aceptarlo o ` +
    `bajan sólo las primeras.</div>`;
}

function renderQueSeDescarga() {
  const caja = $("quePasaAlDescargar");
  if (!caja) return;
  caja.innerHTML = textoDeDescarga({
    hayPesos: !!wbBorrador, hayDolares: !!wbBorradorUsd,
    hayEERR: !!totalesEERR, avisos: avisosEERR,
  });
}

$("btnDescargar").addEventListener("click", async () => {
  const st = $("descargaStatus");
  st.innerHTML = "";
  // un balance que no cierra no se emite: es justo lo que este control viene a evitar
  if (capitalPendiente.length) {
    st.innerHTML = `<div class="status-msg bad">No se descarga nada todavía: el capital de ` +
      `"Pat.Neto" no coincide con la cuenta Capital Suscripto, así que el balance saldría ` +
      `abierto. Está el detalle más arriba.</div>`;
    return;
  }
  const bajaron = [];
  try {
    if (wbBorrador) bajaron.push(await descargar(wbBorrador, "pesos"));
    // el navegador bloquea las descargas seguidas de un mismo clic si van muy pegadas
    if (wbBorradorUsd) { await esperar(400); bajaron.push(await descargar(wbBorradorUsd, "dolares")); }
    if (totalesEERR) { await esperar(400); bajaron.push(descargarEERR()); }
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">Falló una descarga: ${e.message}</div>`;
    return;
  }
  st.innerHTML = resumenDeDescarga(bajaron, !!totalesEERR);
});

const esperar = (ms) => new Promise(r => setTimeout(r, ms));

// El EE RR sale como archivo aparte, con el layout del que se venía armando a mano.
function descargarEERR() {
  const anterior = (estado && estado.eerrAnterior) || null;
  const datos = escribirLibroEERR({
    actual: totalesEERR, anterior, periodoFin: periodoDelEERR(),
    titulo: "SOUTHERN COPPER ARGENTINA S.R.L.",
    movimiento: movimientoEERR,
  });
  const blob = new Blob([datos], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `EE_RR_${periodoDelEERR()}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return a.download;
}

// --------------------------------------------------------- mes anterior del EE RR

// El EE RR compara contra el mes anterior, que sale de `estado.eerrAnterior`. Ese dato lo
// deja cada cierre, así que la primera vez no existe y la columna MES ANTERIOR sale vacía:
// para arrancar se importa del informe del mes pasado ya terminado.
// `detalle` sólo se pide al importar: ahí los tres importes son la única forma de ver que se
// leyó la columna correcta del informe que se sube, y conviene mirarlos antes de confirmar.
// En el estado permanente son ruido: se muestran cada vez que se abre la página sin que haya
// nada que decidir.
function textoEerrAnterior(ant, detalle) {
  if (!ant) {
    return '<div class="status-msg">Todavía no hay un mes anterior registrado, así que la ' +
      'columna MES ANTERIOR del EE RR va a salir vacía. Cargá el informe del mes pasado ' +
      'acá abajo, o va a completarse sola a partir del próximo cierre.</div>';
  }
  const f = (v) => (typeof v === "number" ? v.toLocaleString("es-AR",
    { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "—");
  return '<div class="status-msg ok">Mes anterior registrado: <b>' + (ant.periodo || "sin fecha") +
    '</b>.' + (detalle
      ? '<br>Gastos de operación ' + f(ant.gastosOperacion) +
        ' | administración ' + f(ant.gastosAdministracion) +
        ' | resultado del ejercicio ' + f(ant.resultadoEjercicio)
      : '') + '</div>';
}

function renderEerrAnterior() {
  const caja = $("eerrAnteriorBox");
  if (!caja) return;
  caja.innerHTML = textoEerrAnterior((estado && estado.eerrAnterior) || null);
}

let eerrImportado = null;

if ($("fileEerr")) {
  $("fileEerr").addEventListener("change", async (ev) => {
    const archivo = ev.target.files[0];
    const st = $("importarEerrStatus");
    const btn = $("btnImportarEerr");
    eerrImportado = null;
    btn.disabled = true;
    if (!archivo) { $("txtEerr").textContent = "Elegí el EE RR terminado (.xlsx o .xls)"; st.innerHTML = ""; return; }
    $("txtEerr").textContent = archivo.name;
    st.innerHTML = '<div class="status-msg">Leyendo el archivo…</div>';
    try {
      const r = leerEERRDeArchivo(abrirLibroEERR(await archivo.arrayBuffer()));
      if (!r.totales) {
        st.innerHTML = `<div class="status-msg bad">${r.avisos.join(" ")}</div>`;
        return;
      }
      eerrImportado = r.totales;
      const hayDudas = r.avisos.some(a => /^Ojo|^No encontré/.test(a));
      st.innerHTML = `<div class="status-msg ${hayDudas ? "bad" : "ok"}">` +
        textoEerrAnterior(Object.assign({ periodo: "el que pongas abajo" }, r.totales), true)
          .replace(/^<div class="status-msg[^"]*">/, "").replace(/<\/div>$/, "") + "</div>" +
        r.avisos.map(a => `<div class="status-msg${/^Ojo|^No encontré/.test(a) ? " bad" : ""}">${a}</div>`).join("");
      $("btnImportarEerr").disabled = false;
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">No pude leer el archivo: ${e.message}</div>`;
    }
  });
}

if ($("btnImportarEerr")) {
  $("btnImportarEerr").addEventListener("click", async () => {
    const st = $("importarEerrStatus");
    const periodo = $("periodoEerr").value.trim();
    if (!eerrImportado) return;
    if (!periodo) {
      st.innerHTML = '<div class="status-msg bad">Poné la fecha de cierre de ese informe (por ejemplo 2026-05-31).</div>';
      return;
    }
    if (estado && estado.eerrAnterior && !confirm(
        `Ya hay un mes anterior registrado (${estado.eerrAnterior.periodo}). Reemplazarlo por ${periodo}?`)) {
      return;
    }
    $("btnImportarEerr").disabled = true;
    st.innerHTML = '<div class="status-msg">Guardando…</div>';
    try {
      const nuevo = Object.assign({}, estado, {
        eerrAnterior: Object.assign({ periodo }, eerrImportado),
      });
      await ghcGuardarEstado(nuevo, `EE RR: registra el mes anterior (${periodo})`);
      estado = nuevo;
      renderEerrAnterior();
      st.innerHTML = '<div class="status-msg ok">Listo. El EE RR que emitas ahora va a salir ' +
        'con la columna MES ANTERIOR completa.</div>';
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    } finally {
      $("btnImportarEerr").disabled = false;
    }
  });
}

// El período del EE RR es el cierre del mes que se está cargando: el que se confirmó al
// generar. Antes salía del historial —o sea, del mes ANTERIOR— y si no había historial se
// deducía de la fecha de hoy, que no tiene nada que ver con el período que se emite.
function periodoDelEERR() {
  if (periodoConfirmado) return fpISO(periodoConfirmado);
  const h = (estado && estado.historial) || [];
  const ult = h.length ? h[h.length - 1] : null;
  if (ult && ult.periodoFin) return ult.periodoFin;
  const hoy = new Date();
  const fin = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
  const p = (n) => String(n).padStart(2, "0");
  return `${fin.getFullYear()}-${p(fin.getMonth() + 1)}-${p(fin.getDate())}`;
}

$("fileAprobar").addEventListener("change", () => {
  const f = $("fileAprobar").files[0];
  if (f) $("txtAprobar").textContent = f.name;
  $("btnAprobar").disabled = !f;
});

$("fileAprobarUsd").addEventListener("change", () => {
  const f = $("fileAprobarUsd").files[0];
  if (f) $("txtAprobarUsd").textContent = f.name;
});

$("btnAprobar").addEventListener("click", async () => {
  const st = $("cierreStatus");
  const f = $("fileAprobar").files[0];
  if (!f) return;
  const ok = confirm(
    "Vas a guardar este archivo como el maestro definitivo de pesos.\n\n" +
    "La próxima carga va a partir de él. ¿Confirmás?"
  );
  if (!ok) return;

  mostrar("spinnerCierre", true);
  $("btnAprobar").disabled = true;
  st.innerHTML = "";
  try {
    const buf = await f.arrayBuffer();
    const wb = await abrirWorkbook(buf.slice(0));
    for (const hoja of [HOJA_DISTRIB, HOJA_SUMAS, "Balance"]) {
      if (!wb.getWorksheet(hoja)) throw new Error(`El archivo no tiene la hoja '${hoja}'. ¿Subiste el borrador correcto? NO se guardó nada.`);
    }

    // El EE RR de este mes queda guardado: es el "MES ANTERIOR" de la corrida siguiente.
    const eerrAnterior = totalesEERR
      ? snapshotEERR(totalesEERR, periodoDelEERR())
      : (estado && estado.eerrAnterior) || null;
    // Se guarda QUÉ PERÍODO se cerró, no sólo la fecha en que se apretó el botón: el
    // historial mostraba "30/7/2026" y no había forma de saber a qué mes correspondía.
    const periodoCerrado = periodoDelEERR();
    const nuevoEstado = {
      ultimo_periodo_cerrado: periodoCerrado,
      historial: [...(estado.historial || []), {
        periodo: periodoCerrado,
        fecha: new Date().toISOString(),
        cuentas: resumenBorrador ? resumenBorrador.cuentas : null,
        nuevas: resumenBorrador ? resumenBorrador.nuevas : 0,
        totalPesos: resumenBorrador ? resumenBorrador.total : null,
      }],
      eerrAnterior,
    };
    await ghcGuardarTodo({
      bufferBase: buf,
      estado: nuevoEstado,
      mensaje: `Balance Pesos: cierre de ${periodoCerrado}`,
    });
    const fu = $("fileAprobarUsd").files[0];
    if (fu) {
      const bufu = await fu.arrayBuffer();
      const wbu = await abrirWorkbook(bufu.slice(0));
      if (!hojaDistrib(wbu)) throw new Error("El archivo de dolares no tiene la hoja 'SALDOS'.");
      await ghcGuardarBaseUsd(bufu, "Balance Dolares: nueva carga aprobada");
    }
    st.innerHTML = '<div class="status-msg ok">Guardado como maestro definitivo' +
      (fu ? " (pesos y dolares)" : " (pesos)") + '.</div>';
    setTimeout(() => location.reload(), 1200);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnAprobar").disabled = false;
  } finally {
    mostrar("spinnerCierre", false);
  }
});

arrancar();
