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
    mostrar("cardExport", true);
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
  if (!h.length) return;
  $("historialBody").innerHTML = h.slice().reverse().map(x => `
    <tr>
      <td>${x.fecha ? new Date(x.fecha).toLocaleDateString("es-AR") : "—"}</td>
      <td class="num">${x.cuentas ?? "—"}</td>
      <td class="num">${x.nuevas ?? 0}</td>
      <td class="num">${x.totalPesos != null ? x.totalPesos.toFixed(2) : "—"}</td>
    </tr>`).join("");
  mostrar("cardHistorial", true);
}

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
    for (const hoja of ["SALDOS", "Hoja1", "Anexo II", "Activo y Pasivo", "Balance"]) {
      if (!wb.getWorksheet(hoja)) throw new Error(`El archivo no tiene la hoja '${hoja}'. ¿Es el balance formal en pesos?`);
    }
    const mapeo = derivarMapeoMaestro(wb, "pesos");
    const n = Object.keys(mapeo.cuentas).length;
    $("altaDeteccion").innerHTML = `
      <div class="check ok" style="display:block;">
        <div class="name">Maestro de pesos reconocido: ${n} cuentas en SALDOS</div>
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
    for (const hoja of ["SALDOS", "Hoja1", "Balance"]) {
      if (!wb.getWorksheet(hoja)) {
        throw new Error("El archivo no tiene la hoja '" + hoja + "'. Es el balance formal en dolares?");
      }
    }
    const cm = cuentasDelMaestro(wb, "dolares");
    $("altaUsdStatus").innerHTML =
      '<div class="status-msg ok">Maestro de dolares reconocido: ' + cm.length + ' cuentas en SALDOS.</div>';
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
  logLineas = [];
  // siempre desde una copia limpia del maestro guardado
  const wb = await abrirWorkbook(bufferBase.slice(0));
  log("PESOS");
  const { resumen } = procesarBalance({
    wb, cuentasExport, moneda: "pesos", destinosElegidos: destinos, clasificacion, log,
  });
  wbBorrador = wb;
  resumenBorrador = resumen;

  // Dolares, si el maestro esta cargado. Si faltan decisiones se muestra la pantalla de
  // revision, y el balance en pesos igual queda listo para descargar.
  wbBorradorUsd = null;
  let resumenUsd = null;
  if (bufferBaseUsd) {
    log("\nDOLARES");
    const wbu = await abrirWorkbook(bufferBaseUsd.slice(0));
    try {
      const r = procesarDolares({ wb: wbu, cuentasExport, clasificacion, equivalencias, log });
      wbBorradorUsd = wbu;
      resumenUsd = r.resumen;
      mostrar("cardRevisionUsd", false);
    } catch (e) {
      if (!e.pendientesDolares) throw e;
      pendientesUsd = e.pendientesDolares;
      cuentasMaestroUsd = e.cuentasMaestro;
      renderRevisionUsd();
      mostrar("cardRevisionUsd", true);
      log("  Faltan ubicar " + pendientesUsd.length + " cuenta(s) en dolares: ver el paso de revision.");
    }
  }

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
    mostrar("cardRevisionUsd", false);
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
    { name: "Pesos: cuentas cargadas en Hoja1", value: String(r.cuentas), ok: true,
      detail: "La zona de pegado quedó actualizada con el export fresco." },
    { name: "Pesos: cuentas nuevas insertadas", value: String(r.nuevas), ok: true,
      detail: "Con su fila en SALDOS y su referencia en los estados." },
    { name: "Pesos: cuentas sin enganchar", value: String(r.noEnganchadas.length), ok: r.noEnganchadas.length === 0,
      detail: r.noEnganchadas.length ? r.noEnganchadas.join(", ") : "Todas las cuentas del export quedaron enganchadas a SALDOS." },
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
      '. Estan en SALDOS pero no en Hoja1, asi que su importe queda en cero.</div></div>');
  }
  if (r.duplicadas.length) {
    avisos.push(`
      <div class="check bad" style="display:block;">
        <div class="name">Pesos: una cuenta del export está repetida en la hoja SALDOS</div>
        <div class="detail">
          ${r.duplicadas.map(d => `<b>${d.codigo}</b> (filas ${d.filaPrevia} y ${d.fila})`).join(", ")}.
          Son filas del archivo <b>de pesos</b>, no del de dólares: los dos tienen una hoja
          llamada SALDOS y sus filas no coinciden. Como esa cuenta sí viene en el export, se
          carga en la primera de las dos; conviene revisar cuál corresponde.
        </div>
      </div>`);
  }
  $("avisosBody").innerHTML = avisos.join("");
}

// --------------------------------------------------------------- descarga y cierre

async function descargar(wb, sufijo) {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const hoy = new Date();
  a.href = url;
  a.download = "SCA_Balance_" + hoy.getFullYear() + "-" +
    String(hoy.getMonth() + 1).padStart(2, "0") + "_" + sufijo + ".xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

$("btnDescargar").addEventListener("click", async () => {
  if (wbBorrador) await descargar(wbBorrador, "pesos");
  if (wbBorradorUsd) await descargar(wbBorradorUsd, "dolares");
});

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
    for (const hoja of ["SALDOS", "Hoja1", "Balance"]) {
      if (!wb.getWorksheet(hoja)) throw new Error(`El archivo no tiene la hoja '${hoja}'. ¿Subiste el borrador correcto? NO se guardó nada.`);
    }

    const nuevoEstado = {
      historial: [...(estado.historial || []), {
        fecha: new Date().toISOString(),
        cuentas: resumenBorrador ? resumenBorrador.cuentas : null,
        nuevas: resumenBorrador ? resumenBorrador.nuevas : 0,
        totalPesos: resumenBorrador ? resumenBorrador.total : null,
      }],
    };
    await ghcGuardarTodo({
      bufferBase: buf,
      estado: nuevoEstado,
      mensaje: "Balance Pesos: nueva carga aprobada",
    });
    const fu = $("fileAprobarUsd").files[0];
    if (fu) {
      const bufu = await fu.arrayBuffer();
      const wbu = await abrirWorkbook(bufu.slice(0));
      if (!wbu.getWorksheet("SALDOS")) throw new Error("El archivo de dolares no tiene la hoja 'SALDOS'.");
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
