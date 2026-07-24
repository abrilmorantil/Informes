// Pega todo junto: lee el balance maestro y el export de Onvio, corre el motor,
// pide las categorías de las cuentas nuevas y ofrece la descarga.

let mapeoBase = null;      // mapeo tal como está guardado (sin las cuentas de esta corrida)
let mapeoSha = null;       // sha de GitHub, para poder volver a escribirlo
let mapeoResultante = null;// mapeo devuelto por el motor
let lineas = null;
let pendientes = [];
let wbResultado = null;
let logLineas = [];

const $ = (id) => document.getElementById(id);

// --------------------------------------------------------------------
// Carga del mapeo (de GitHub si está configurado, si no del propio sitio)
// --------------------------------------------------------------------

async function cargarMapeo() {
  const status = $("mapeoStatus");
  if (hasGhSettings()) {
    try {
      const { mapeo, sha } = await ghGetMapeo();
      mapeoBase = mapeo;
      mapeoSha = sha;
      status.textContent = `Mapeo cargado desde GitHub (${Object.keys(mapeo.cuentas).length} cuentas).`;
      status.style.color = "";
      return;
    } catch (e) {
      status.textContent = "No pude leer el mapeo desde GitHub: " + e.message;
      status.style.color = "var(--warn)";
    }
  }
  try {
    const res = await fetch("mapeo_maestro.json");
    mapeoBase = await res.json();
    mapeoSha = null;
    if (!hasGhSettings()) {
      status.textContent = `Mapeo cargado del sitio (${Object.keys(mapeoBase.cuentas).length} cuentas). ` +
        `Configurá GitHub (⚙) si querés que las cuentas nuevas queden guardadas.`;
    }
  } catch (e) {
    status.textContent = "No encontré mapeo_maestro.json. Configurá GitHub o subilo junto al sitio.";
    status.style.color = "var(--warn)";
  }
}

// --------------------------------------------------------------------
// Configuración de GitHub
// --------------------------------------------------------------------

function openSettings() {
  const s = loadGhSettings();
  $("cfgToken").value = s.token || "";
  $("cfgRepo").value = s.repo || "";
  $("cfgBranch").value = s.branch || "main";
  $("cfgPath").value = s.path || "informe-a/mapeo_maestro.json";
  $("settingsStatus").innerHTML = "";
  $("settingsModal").classList.remove("hidden");
}

function closeSettings() {
  $("settingsModal").classList.add("hidden");
}

async function saveSettings() {
  saveGhSettings({
    token: $("cfgToken").value.trim(),
    repo: $("cfgRepo").value.trim(),
    branch: $("cfgBranch").value.trim() || "main",
    path: $("cfgPath").value.trim() || "informe-a/mapeo_maestro.json",
  });
  const status = $("settingsStatus");
  status.innerHTML = '<div class="status-msg">Probando conexión…</div>';
  try {
    await ghGetMapeo();
    status.innerHTML = '<div class="status-msg ok">Listo, conectado a GitHub correctamente.</div>';
    await cargarMapeo();
    setTimeout(closeSettings, 900);
  } catch (e) {
    status.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

$("howToToken").addEventListener("click", (ev) => {
  ev.preventDefault();
  alert(
    "Cómo generar el token:\n\n" +
    "1. Andá a github.com → foto de perfil → Settings.\n" +
    "2. Developer settings → Personal access tokens → Tokens (classic).\n" +
    "3. Generate new token (classic).\n" +
    "4. Marcá el scope 'repo' (acceso al repositorio).\n" +
    "5. Generá y copiá el token (empieza con 'ghp_') — GitHub lo muestra una sola vez.\n\n" +
    "Queda guardado solo en este navegador, en tu computadora."
  );
});

// --------------------------------------------------------------------
// Selección de archivos
// --------------------------------------------------------------------

function conectarDropzone(inputId, textoId, etiqueta) {
  $(inputId).addEventListener("change", () => {
    const f = $(inputId).files[0];
    if (f) $(textoId).textContent = `${etiqueta}: ${f.name}`;
    $("btnProcesar").disabled = !($("fileBalance").files.length && $("fileOnvio").files.length);
  });
}
conectarDropzone("fileBalance", "txtBalance", "Balance maestro");
conectarDropzone("fileOnvio", "txtOnvio", "Export de Onvio");

// --------------------------------------------------------------------
// Procesar
// --------------------------------------------------------------------

function log(msg) {
  logLineas.push(String(msg));
  $("logBody").textContent = logLineas.join("\n");
}

function resetPantalla() {
  logLineas = [];
  wbResultado = null;
  mapeoResultante = null;
  $("cardPendientes").classList.add("hidden");
  $("cardResultado").classList.add("hidden");
  $("cardDescarga").classList.add("hidden");
  $("descargaStatus").innerHTML = "";
  $("pendientesStatus").innerHTML = "";
}

async function leerBalance() {
  const buf = await $("fileBalance").files[0].arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

async function leerLineasOnvio() {
  const buf = await $("fileOnvio").files[0].arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  return parseOnvioExport(rows);
}

$("btnProcesar").addEventListener("click", async () => {
  resetPantalla();
  $("spinner").classList.remove("hidden");
  $("btnProcesar").disabled = true;
  try {
    if (!mapeoBase) await cargarMapeo();
    if (!mapeoBase) throw new Error("No hay mapeo cargado todavía.");

    lineas = await leerLineasOnvio();
    const deteccion = detectarPendientes(lineas, mapeoBase);
    pendientes = deteccion.pendientes;

    if (deteccion.sinCc.length) {
      alert(
        "Estos centros de costo del export no se pudieron identificar en el balance y " +
        "sus líneas no se van a cargar:\n\n" + deteccion.sinCc.join("\n")
      );
    }

    if (pendientes.length) {
      renderPendientes();
      $("cardPendientes").classList.remove("hidden");
    } else {
      await correrMotor({});
    }
  } catch (e) {
    alert("Error: " + e.message);
  } finally {
    $("spinner").classList.add("hidden");
    $("btnProcesar").disabled = false;
  }
});

function renderPendientes() {
  const cats = categoriasDisponibles(mapeoBase);
  const body = $("pendientesBody");
  body.innerHTML = "";
  pendientes.forEach((p, idx) => {
    const tr = document.createElement("tr");
    tr.className = "pending-row";
    const opciones = cats.map(c => `<option value="${c}">${c}</option>`).join("");
    tr.innerHTML = `
      <td>${p.codigo}</td>
      <td>${p.label}</td>
      <td>${p.cc_nombre}</td>
      <td class="num">${p.saldo.toFixed(2)}</td>
      <td>
        <select class="catSelect" data-idx="${idx}">
          <option value="">— elegí una categoría —</option>
          ${opciones}
        </select>
      </td>`;
    body.appendChild(tr);
  });
}

$("btnConfirmar").addEventListener("click", async () => {
  const selects = document.querySelectorAll(".catSelect");
  const elegidas = {};
  for (let i = 0; i < selects.length; i++) {
    const valor = selects[i].value;
    if (!valor) {
      $("pendientesStatus").innerHTML =
        '<div class="status-msg bad">Falta elegir la categoría de al menos una cuenta.</div>';
      return;
    }
    elegidas[pendientes[i].codigo] = valor;
  }
  $("pendientesStatus").innerHTML = "";
  $("btnConfirmar").disabled = true;
  try {
    await correrMotor(elegidas);
    $("cardPendientes").classList.add("hidden");
  } catch (e) {
    $("pendientesStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    $("btnConfirmar").disabled = false;
  }
});

async function correrMotor(categoriasElegidas) {
  logLineas = [];
  // Se relee el balance en cada corrida: el workbook se modifica en el lugar, así que
  // reusarlo tras un intento fallido arrastraría cambios a medio aplicar.
  const wb = await leerBalance();
  const periodo = $("periodo").value.trim() || null;

  const { mapeo, resumen } = procesar({
    wb, lineas, mapeo: mapeoBase, categoriasElegidas, periodo, log,
  });

  wbResultado = wb;
  mapeoResultante = mapeo;

  renderChecks(resumen);
  $("cardResultado").classList.remove("hidden");
  $("cardDescarga").classList.remove("hidden");

  const btnGuardar = $("btnGuardarMapeo");
  if (resumen.nuevas > 0 || resumen.clasificadas > 0) {
    btnGuardar.classList.remove("hidden");
  } else {
    btnGuardar.classList.add("hidden");
  }
}

function renderChecks(resumen) {
  const totalSaldo = lineas.reduce((s, l) => s + l.saldo, 0);
  const items = [
    {
      name: "Total en USD del export de Onvio",
      detail: "Es la cifra que tiene que dar la hoja. Excel la recalcula al abrir el archivo.",
      value: totalSaldo.toFixed(2), ok: true,
    },
    {
      name: "Líneas de cuenta cargadas",
      detail: "Cada línea del export se cargó en su fila y centro de costo.",
      value: String(lineas.length), ok: true,
    },
    {
      name: "Cuentas ya conocidas",
      detail: "Estaban en el mapeo y se actualizaron en su lugar.",
      value: String(resumen.conocidas), ok: true,
    },
    {
      name: "Cuentas nuevas insertadas",
      detail: "Se les creó la fila y se reacomodaron las fórmulas dependientes.",
      value: String(resumen.nuevas), ok: true,
    },
  ];
  if (resumen.clasificadas) {
    items.push({
      name: "Cuentas recién clasificadas",
      detail: "Ya existían en la hoja pero no tenían categoría asignada.",
      value: String(resumen.clasificadas), ok: true,
    });
  }
  if (resumen.sinCc.length) {
    items.push({
      name: "Centros de costo no identificados",
      detail: "Sus líneas NO se cargaron: " + resumen.sinCc.join(", "),
      value: String(resumen.sinCc.length), ok: false,
    });
  }

  const body = $("checksBody");
  body.innerHTML = "";
  items.forEach(c => {
    const div = document.createElement("div");
    div.className = "check " + (c.ok ? "ok" : "bad");
    div.innerHTML = `
      <div><div class="name">${c.name}</div><div class="detail">${c.detail}</div></div>
      <div style="display:flex; align-items:center; gap:10px;">
        <span class="value">${c.value}</span>
        <span class="badge ${c.ok ? "ok" : "bad"}">${c.ok ? "OK" : "Revisar"}</span>
      </div>`;
    body.appendChild(div);
  });
}

// --------------------------------------------------------------------
// Descarga y guardado del mapeo
// --------------------------------------------------------------------

$("btnDescargar").addEventListener("click", async () => {
  if (!wbResultado) return;
  const periodo = $("periodo").value.trim() || "sin-periodo";
  const buffer = await wbResultado.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `BALANCE_DE_COMPROBACION_USD_${periodo}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// El mapeo solo se guarda cuando hubo cuentas nuevas, y avisando: al insertarse una
// fila cambian las filas de las cuentas siguientes, así que a partir de ahí hay que
// seguir trabajando sobre el archivo que generó esta corrida, no sobre el original.
$("btnGuardarMapeo").addEventListener("click", async () => {
  const status = $("descargaStatus");
  if (!hasGhSettings()) {
    status.innerHTML = '<div class="status-msg bad">Configurá GitHub (⚙, arriba a la derecha) antes de guardar.</div>';
    openSettings();
    return;
  }
  const ok = confirm(
    "Se va a guardar el mapeo con las cuentas nuevas ya incorporadas.\n\n" +
    "Importante: a partir de ahora el mapeo corresponde al archivo que generó ESTA " +
    "corrida. Si el mes que viene volvés a partir del balance viejo, el sistema te va " +
    "a avisar que están desincronizados.\n\n¿Guardar igual?"
  );
  if (!ok) return;

  $("btnGuardarMapeo").disabled = true;
  status.innerHTML = '<div class="status-msg">Guardando en GitHub…</div>';
  try {
    // Se relee justo antes de escribir para tomar el sha más reciente.
    const { sha } = await ghGetMapeo();
    const codigos = pendientes.map(p => p.codigo).join(", ");
    await ghPutMapeo(mapeoResultante, sha,
      `Balance USD: incorpora ${pendientes.length} cuenta(s) nueva(s): ${codigos}`);
    mapeoBase = mapeoResultante;
    mapeoSha = null;
    status.innerHTML = '<div class="status-msg ok">Mapeo guardado en GitHub.</div>';
    $("btnGuardarMapeo").classList.add("hidden");
  } catch (e) {
    status.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    $("btnGuardarMapeo").disabled = false;
  }
});

cargarMapeo();
