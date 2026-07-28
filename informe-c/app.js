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
    const b = await ghcLeerBase();
    if (!b) throw new Error("Hay un estado guardado pero falta base_pesos.xlsx en el repositorio.");
    bufferBase = b.buffer;

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
          ? `Ojo: hay ${mapeo.duplicadas.length} cuentas repetidas (${mapeo.duplicadas.map(d => d.codigo).join(", ")}). Se usa la primera de cada una; conviene limpiarlas en Excel cuando puedas.`
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
    nuevasPendientes = nuevas.filter(n => !/^2110/.test(n.codigo));
    listasDestino = { madres: madresResultados(wb, "pesos"), lineasNota4: lineasDeNota4(wb) };

    const provNuevos = nuevas.length - nuevasPendientes.length;
    if (provNuevos > 0) {
      $("exportStatus").innerHTML +=
        `<div class="status-msg ok">${provNuevos} proveedor(es) nuevo(s) se insertan solos en el detalle del pasivo.</div>`;
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
  const { resumen } = procesarBalance({
    wb, cuentasExport, moneda: "pesos", destinosElegidos: destinos, log,
  });
  wbBorrador = wb;
  resumenBorrador = resumen;
  renderResultado(resumen);
  mostrar("cardResultado", true);
  mostrar("cardCierre", true);
}

function renderResultado(r) {
  const items = [
    { name: "Total en pesos del export", value: r.total.toFixed(2), ok: true,
      detail: "Excel recalcula los estados al abrir el archivo; el control L23 del Balance tiene que dar 0." },
    { name: "Cuentas cargadas en Hoja1", value: String(r.cuentas), ok: true,
      detail: "La zona de pegado quedó reescrita con el export fresco." },
    { name: "Cuentas nuevas insertadas", value: String(r.nuevas), ok: true,
      detail: "Con su fila en SALDOS y su referencia en los estados." },
    { name: "Cuentas sin enganchar", value: String(r.noEnganchadas.length), ok: r.noEnganchadas.length === 0,
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

  const avisos = [];
  if (r.duplicadas.length) {
    avisos.push(`
      <div class="check bad" style="display:block;">
        <div class="name">SALDOS tiene cuentas repetidas</div>
        <div class="detail">${r.duplicadas.map(d => `${d.codigo} (filas ${d.filaPrevia} y ${d.fila})`).join(", ")}.
        Se usa la primera de cada una; conviene limpiar el archivo en Excel.</div>
      </div>`);
  }
  $("avisosBody").innerHTML = avisos.join("");
}

// --------------------------------------------------------------- descarga y cierre

$("btnDescargar").addEventListener("click", async () => {
  if (!wbBorrador) return;
  const buffer = await wbBorrador.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const hoy = new Date();
  a.download = `SCA_Balance_${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, "0")}_pesos.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

$("fileAprobar").addEventListener("change", () => {
  const f = $("fileAprobar").files[0];
  if (f) $("txtAprobar").textContent = f.name;
  $("btnAprobar").disabled = !f;
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
    st.innerHTML = '<div class="status-msg ok">Guardado como maestro definitivo.</div>';
    setTimeout(() => location.reload(), 1200);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnAprobar").disabled = false;
  } finally {
    mostrar("spinnerCierre", false);
  }
});

arrancar();
