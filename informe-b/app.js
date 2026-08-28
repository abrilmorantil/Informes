// Pega todo junto: lee el archivo de SISE, corre la lógica de core.js,
// muestra los resultados, y guarda cuentas nuevas en GitHub.

let lastResult = null;   // { cuentas, control, categoryTotals, mapping, mappingSha }
let currentMapping = null;
let currentMappingSha = null;

// --------------------------------------------------------------------
// Carga inicial del mapping.json (por GitHub si ya está configurado, si
// no, del propio sitio estático como fallback de solo lectura)
// --------------------------------------------------------------------

let estadoB = { ultimo_periodo: null, saldos: {}, historial: [] };

async function loadMapping() {
  const statusEl = document.getElementById("mappingStatus");
  if (hasGhSettings()) {
    try {
      const { mapping, sha } = await ghGetFile();
      currentMapping = mapping;
      currentMappingSha = sha;
      statusEl.textContent = `Mapeo cargado desde GitHub (${mapping.length} cuentas).`;
      statusEl.style.color = "";
      return;
    } catch (e) {
      statusEl.textContent = "No pude leer el mapeo desde GitHub: " + e.message;
      statusEl.style.color = "var(--warn)";
      // sigue al fallback de abajo
    }
  }
  try {
    // El mapping vive en la RAIZ del repo, que es donde github.js lo lee y lo escribe
    // (`s.path || "mapping.json"`). Pedirlo sin la subida de carpeta traia
    // `informe-b/mapping.json`, que es una copia vieja de julio: la pagina mostraba una
    // configuracion que no era la que se guarda. Un solo archivo, una sola verdad.
    const res = await fetch("../mapping.json");
    currentMapping = await res.json();
    currentMappingSha = null;
    if (!hasGhSettings()) {
      statusEl.textContent = `Mapeo cargado del sitio (${currentMapping.length} cuentas). Configurá GitHub (⚙) para poder guardar cuentas nuevas.`;
    }
  } catch (e) {
    statusEl.textContent = "No encontré mapping.json. Configurá GitHub o subilo junto al sitio.";
    statusEl.style.color = "var(--warn)";
  }
}

// --------------------------------------------------------------------
// Configuración de GitHub
// --------------------------------------------------------------------

function openSettings() {
  const s = loadGhSettings();
  document.getElementById("cfgToken").value = s.token || "";
  document.getElementById("cfgRepo").value = s.repo || "";
  document.getElementById("cfgBranch").value = s.branch || "main";
  document.getElementById("cfgPath").value = s.path || "mapping.json";
  document.getElementById("settingsStatus").innerHTML = "";
  document.getElementById("settingsModal").classList.remove("hidden");
}

function closeSettings() {
  document.getElementById("settingsModal").classList.add("hidden");
}

async function saveSettings() {
  const settings = {
    token: document.getElementById("cfgToken").value.trim(),
    repo: document.getElementById("cfgRepo").value.trim(),
    branch: document.getElementById("cfgBranch").value.trim() || "main",
    path: document.getElementById("cfgPath").value.trim() || "mapping.json",
  };
  saveGhSettings(settings);
  const statusEl = document.getElementById("settingsStatus");
  statusEl.innerHTML = '<div class="status-msg">Probando conexión…</div>';
  try {
    await ghGetFile();
    statusEl.innerHTML = '<div class="status-msg ok">Listo, conectado a GitHub correctamente.</div>';
    await loadMapping();
    setTimeout(closeSettings, 900);
  } catch (e) {
    statusEl.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

document.getElementById("howToToken").addEventListener("click", () => {
  alert(
    "Cómo generar el token:\n\n" +
    "1. Andá a github.com → foto de perfil → Settings.\n" +
    "2. Developer settings → Personal access tokens → Tokens (classic).\n" +
    "3. Generate new token (classic).\n" +
    "4. Marcá el scope 'repo' (acceso al repositorio).\n" +
    "5. Generá y copiá el token (empieza con 'ghp_') — no lo vas a poder ver de nuevo, pegalo acá.\n\n" +
    "Queda guardado solo en este navegador, en tu computadora."
  );
});

// --------------------------------------------------------------------
// 1. Procesar el archivo de SISE
// --------------------------------------------------------------------

const fileInput = document.getElementById("fileInput");
const dropText = document.getElementById("dropText");
const btnProcesar = document.getElementById("btnProcesar");

fileInput.addEventListener("change", () => {
  if (fileInput.files.length) {
    dropText.textContent = fileInput.files[0].name;
    btnProcesar.disabled = false;
  }
});

btnProcesar.addEventListener("click", async () => {
  document.getElementById("spinner").classList.remove("hidden");
  btnProcesar.disabled = true;
  try {
    if (!currentMapping) await loadMapping();
    if (!currentMapping) throw new Error("No hay mapeo cargado todavía.");

    const file = fileInput.files[0];
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    const { cuentas, control, categoryTotals } = parseSiseExport(rows);
    const unmapped = findUnmapped(cuentas, currentMapping);
    const duplicates = findDuplicateCodes(currentMapping);
    const { lineas, totalDebe, totalHaber } = buildBalance(cuentas, currentMapping, {});
    const { checks, allOk, categoryDiffs } = runValidation(
      lineas, totalDebe, totalHaber, control, unmapped, categoryTotals, duplicates);

    // Se guardan las filas crudas del export para poder meterlas como una hoja mas del
    // informe: asi el archivo se explica solo sin depender de encontrar el .xls despues.
    lastResult = { cuentas, control, categoryTotals, checks, allOk, unmapped, duplicates, categoryDiffs, lineas, filasExport: rows };
    render(lastResult);
  } catch (e) {
    alert("Error: " + e.message);
  } finally {
    document.getElementById("spinner").classList.add("hidden");
    btnProcesar.disabled = false;
  }
});

// --------------------------------------------------------------------
// Render
// --------------------------------------------------------------------

function render(data) {
  renderUnmapped(data.unmapped, CATEGORIES);

  // La validacion solo se muestra cuando algo NO cierra. Cuando todo da bien son tres
  // "OK 0,00" verdes que no dicen nada y ocupan media pantalla todos los meses; y cuando
  // algo falla, al no haber nada mas en la tarjeta, salta a la vista.
  //
  // No se saca del todo porque es lo unico que avisa que el balance no ata con Onvio, y
  // ademas es lo que traba el boton de descarga: sin la tarjeta, el boton quedaria
  // deshabilitado sin explicacion.
  const hayProblema = !data.allOk;
  if (hayProblema) {
    renderChecks(data.checks);
    renderCategoryDiffs(data.categoryDiffs || []);
    renderDuplicates(data.duplicates || {});
  }
  document.getElementById("cardChecks").classList.toggle("hidden", !hayProblema);
  const btnFinalizar = document.getElementById("btnFinalizar");
  if (data.unmapped.length === 0 && data.allOk) {
    document.getElementById("cardUnmapped").classList.add("hidden");
    document.getElementById("cardFinal").classList.remove("hidden");
    btnFinalizar.disabled = false;
    document.getElementById("finalWarning").classList.add("hidden");
  } else if (data.unmapped.length === 0) {
    document.getElementById("cardUnmapped").classList.add("hidden");
    document.getElementById("cardFinal").classList.remove("hidden");
    btnFinalizar.disabled = true;
    document.getElementById("finalWarning").classList.remove("hidden");
  } else {
    document.getElementById("cardUnmapped").classList.remove("hidden");
    document.getElementById("cardFinal").classList.add("hidden");
  }
}

function renderUnmapped(unmapped, categorias) {
  const body = document.getElementById("unmappedBody");
  body.innerHTML = "";
  window._candidatasMadre = (currentMapping || [])
    .filter(e => e.category === "RESULTADOS")
    .map(e => ({ code: e.code, description: e.description }))
    .sort((a, b) => a.description.toLowerCase().localeCompare(b.description.toLowerCase()));

  unmapped.forEach((u, idx) => {
    const tr = document.createElement("tr");
    tr.className = "pending-row";
    tr.innerHTML = `
      <td>${u.code}</td>
      <td>${u.description}</td>
      <td>${u.debe.toFixed(2)}</td>
      <td>${u.haber.toFixed(2)}</td>
      <td>
        <select data-idx="${idx}" class="catSelect">
          ${categorias.map(c => `<option value="${c}" ${c === u.suggested_category ? "selected" : ""}>${c}</option>`).join("")}
        </select>
      </td>
      <td>
        <select data-idx="${idx}" class="parentCode">
          <option value="">— sin cuenta madre —</option>
          ${window._candidatasMadre.map(c =>
            `<option value="${c.code}">${c.code} — ${c.description}</option>`).join("")}
        </select>
      </td>
    `;
    body.appendChild(tr);
  });

  // El mismo desplegable con búsqueda adentro que usan los otros informes. Antes acá había
  // un typeahead propio: había que escribir para que apareciera algo, no se podía abrir la
  // lista y mirarla, y cortaba en 25 resultados sin decirlo.
  conBuscadorTodos(".parentCode", "Buscar cuenta madre…");
}

function renderChecks(checks) {
  const body = document.getElementById("checksBody");
  body.innerHTML = "";
  checks.forEach(c => {
    const div = document.createElement("div");
    div.className = "check " + (c.ok ? "ok" : "bad");
    div.innerHTML = `
      <div><div class="name">${c.name}</div><div class="detail">${c.detail}</div></div>
      <div style="display:flex; align-items:center; gap:10px;">
        <span class="value">${c.value.toFixed(2)}</span>
        <span class="badge ${c.ok ? "ok" : "bad"}">${c.ok ? "OK" : "Revisar"}</span>
      </div>`;
    body.appendChild(div);
  });
}

function renderCategoryDiffs(diffs) {
  const body = document.getElementById("categoryBody");
  body.innerHTML = "";
  if (!diffs.length || diffs.every(d => d.ok)) return;
  const title = document.createElement("div");
  title.style.cssText = "font-weight:600; font-size:13.5px; margin-bottom:8px;";
  title.textContent = "Dónde está la diferencia, por categoría";
  body.appendChild(title);
  diffs.forEach(d => {
    const div = document.createElement("div");
    div.className = "check " + (d.ok ? "ok" : "bad");
    div.innerHTML = `
      <div><div class="name">${d.categoria}</div></div>
      <div style="display:flex; align-items:center; gap:14px;">
        <span class="value">Debe: ${d.diff_debe.toFixed(2)}</span>
        <span class="value">Haber: ${d.diff_haber.toFixed(2)}</span>
        <span class="badge ${d.ok ? "ok" : "bad"}">${d.ok ? "OK" : "Acá"}</span>
      </div>`;
    body.appendChild(div);
  });
}

function renderDuplicates(duplicates) {
  const body = document.getElementById("duplicatesBody");
  body.innerHTML = "";
  const codes = Object.keys(duplicates);
  if (!codes.length) return;
  const div = document.createElement("div");
  div.className = "check bad";
  div.style.flexDirection = "column";
  div.style.alignItems = "flex-start";
  let html = '<div class="name">Códigos duplicados en el mapeo</div>';
  html += '<div class="detail">Estas cuentas están mapeadas en más de un lugar, así que su Debe/Haber se está sumando doble:</div>';
  codes.forEach(code => {
    html += `<div style="margin-top:6px; font-size:13px;"><b>${code}</b>: ${duplicates[code].join(" · ")}</div>`;
  });
  div.innerHTML = html;
  body.appendChild(div);
}

// --------------------------------------------------------------------
// 2. Clasificar cuentas nuevas (guarda en GitHub)
// --------------------------------------------------------------------

document.getElementById("btnClasificar").addEventListener("click", async () => {
  const statusEl = document.getElementById("clasificarStatus");
  if (!hasGhSettings()) {
    statusEl.innerHTML = '<div class="status-msg bad">Configurá GitHub (⚙, arriba a la derecha) antes de guardar.</div>';
    openSettings();
    return;
  }

  const cats = document.querySelectorAll(".catSelect");
  const parents = document.querySelectorAll(".parentCode");
  const nuevas = lastResult.unmapped.map((u, idx) => ({
    code: u.code, description: u.description,
    category: cats[idx].value, parent_code: (parents[idx].value || "").trim() || null,
  }));

  document.getElementById("spinnerClasificar").classList.remove("hidden");
  document.getElementById("btnClasificar").disabled = true;
  statusEl.innerHTML = "";
  try {
    // Releer el mapeo justo antes de escribir, para tener el sha más
    // reciente (por si alguien más lo tocó mientras tanto).
    const { mapping, sha } = await ghGetFile();
    for (const n of nuevas) {
      if (n.category === "RESULTADOS" && n.parent_code) {
        const madre = mapping.find(e => e.code === n.parent_code);
        if (madre) {
          madre.children = madre.children || [];
          madre.children.push({ code: n.code, description: n.description });
          madre.type = "parent";
          continue;
        }
      }
      mapping.push({ code: n.code, description: n.description, category: n.category, type: "simple", aliases: [] });
    }

    const mensaje = `Agrega ${nuevas.length} cuenta(s) nueva(s): ${nuevas.map(n => n.code).join(", ")}`;
    await ghPutFile(mapping, sha, mensaje);
    currentMapping = mapping;
    statusEl.innerHTML = '<div class="status-msg ok">Guardado en GitHub. Recalculando…</div>';

    // Recalcular con el mapeo actualizado
    const { cuentas, control, categoryTotals } = lastResult;
    const unmapped2 = findUnmapped(cuentas, currentMapping);
    const duplicates2 = findDuplicateCodes(currentMapping);
    const { lineas, totalDebe, totalHaber } = buildBalance(cuentas, currentMapping, {});
    const { checks, allOk, categoryDiffs } = runValidation(
      lineas, totalDebe, totalHaber, control, unmapped2, categoryTotals, duplicates2);
    lastResult = { ...lastResult, checks, allOk, unmapped: unmapped2, duplicates: duplicates2, categoryDiffs, lineas };
    render(lastResult);
  } catch (e) {
    statusEl.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    document.getElementById("spinnerClasificar").classList.add("hidden");
    document.getElementById("btnClasificar").disabled = false;
  }
});

// --------------------------------------------------------------------
// 4. Finalizar y descargar
// --------------------------------------------------------------------

document.getElementById("btnFinalizar").addEventListener("click", async () => {
  const periodo = document.getElementById("periodo").value || "sin-fecha";
  if (!lastResult.allOk) {
    alert("Todavía queda algo en rojo en la validación. Resolvelo antes de generar el archivo.");
    return;
  }
  const wb = await writeOutputXlsx(lastResult.lineas, periodo, estadoB.saldos, lastResult.filasExport);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `BALCOMPROBDOLARES_${periodo}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// --------------------------------------------------------------------
// Saldos del mes anterior
// --------------------------------------------------------------------

function renderEstadoB() {
  const caja = document.getElementById("estadoBBox");
  if (!caja) return;
  const n = Object.keys(estadoB.saldos || {}).length;
  if (!estadoB.ultimo_periodo) {
    caja.innerHTML =
      '<div class="status-msg">Todavía no hay ningún mes registrado, así que la columna ' +
      '"Saldo anterior" va a salir en amarillo para completarla a mano. Cuando registres ' +
      'este mes, el próximo ya la trae cargada.</div>';
    return;
  }
  caja.innerHTML =
    '<div class="status-msg ok">Último mes registrado: <b>' + estadoB.ultimo_periodo +
    '</b> — ' + n + ' cuentas con saldo guardado. La columna "Saldo anterior" del archivo ' +
    'sale con esos importes.</div>';
}

async function cargarEstadoB() {
  if (!hasGhSettings()) return;
  estadoB = await leerEstadoB();
  renderEstadoB();
}

// --- Traer los saldos de un informe del mes pasado ya terminado ---------------

let saldosImportados = null;

const inputSaldos = document.getElementById("fileSaldos");
if (inputSaldos) {
  inputSaldos.addEventListener("change", async (ev) => {
    const archivo = ev.target.files[0];
    const texto = document.getElementById("textoSaldos");
    const st = document.getElementById("importarStatus");
    const btn = document.getElementById("btnImportarSaldos");
    saldosImportados = null;
    btn.disabled = true;
    if (!archivo) { texto.textContent = "Elegí el informe terminado (.xlsx)"; st.innerHTML = ""; return; }

    texto.textContent = archivo.name;
    st.innerHTML = '<div class="status-msg">Leyendo el archivo…</div>';
    try {
      const wb = await abrirLibroDeSaldos(await archivo.arrayBuffer(), archivo.name);
      const r = leerSaldosDeBalcomp(wb);
      if (!r.cuentas) {
        st.innerHTML = '<div class="status-msg bad">' + r.avisos.join(" ") + '</div>';
        return;
      }
      saldosImportados = r.saldos;
      const total = Object.values(r.saldos).reduce((a, v) => a + v, 0);
      // el residuo de un peso sale de redondear cada saldo a dos decimales. Ojo: esto es
      // una condición necesaria y NO suficiente — la columna de movimientos también netea
      // a cero, así que el control de verdad es el nombre de la columna, que se avisa aparte.
      const cierra = Math.abs(total) <= 1;
      const fmt = (v) => v.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      st.innerHTML =
        '<div class="status-msg ' + (cierra ? "ok" : "") + '">Encontré <b>' + r.cuentas +
        ' cuentas</b> en la hoja "' + r.hoja + '"' +
        (r.filas !== r.cuentas ? ', leídas de ' + r.filas + ' filas' : '') + '. ' +
        (cierra
          ? "La suma de los saldos cierra en cero, como tiene que dar un balance."
          : "<b>Ojo:</b> la suma de los saldos da " + fmt(total) + ", y en un balance " +
            "tendría que cerrar en cero. Revisá que sea la columna correcta.") +
        '</div>' +
        r.avisos.map(a => '<div class="status-msg">' + a + '</div>').join("");
      btn.disabled = false;
    } catch (e) {
      st.innerHTML = '<div class="status-msg bad">No pude leer el archivo: ' + e.message + '</div>';
    }
  });
}

const btnImportarB = document.getElementById("btnImportarSaldos");
if (btnImportarB) {
  btnImportarB.addEventListener("click", async () => {
    const st = document.getElementById("importarStatus");
    const periodo = document.getElementById("periodoSaldos").value.trim();
    if (!hasGhSettings()) {
      st.innerHTML = '<div class="status-msg bad">Configurá GitHub primero: los saldos se guardan ahí.</div>';
      return;
    }
    if (!periodo) {
      st.innerHTML = '<div class="status-msg bad">Poné de qué mes es el informe que subiste.</div>';
      return;
    }
    if (!saldosImportados) return;
    if (estadoB.ultimo_periodo && !confirm(
        `Ya hay saldos registrados de ${estadoB.ultimo_periodo}. Importar los de ${periodo} los reemplaza. ¿Seguimos?`)) {
      return;
    }
    btnImportarB.disabled = true;
    st.innerHTML = '<div class="status-msg">Guardando los saldos…</div>';
    try {
      const archivo = document.getElementById("fileSaldos").files[0];
      estadoB = await sembrarEstadoB(estadoB, periodo, saldosImportados, archivo ? archivo.name : null);
      renderEstadoB();
      st.innerHTML = '<div class="status-msg ok">Listo: quedaron guardados los saldos de ' + periodo +
        '. La próxima corrida trae la columna "Saldo anterior" ya cargada.</div>';
    } catch (e) {
      st.innerHTML = '<div class="status-msg bad">' + e.message + '</div>';
    } finally {
      btnImportarB.disabled = false;
    }
  });
}

const btnRegistrarB = document.getElementById("btnRegistrarMes");
if (btnRegistrarB) {
  btnRegistrarB.addEventListener("click", async () => {
    const periodo = document.getElementById("periodo").value;
    const st = document.getElementById("registrarStatus");
    if (!periodo) {
      st.innerHTML = '<div class="status-msg bad">Poné el período antes de registrarlo.</div>';
      return;
    }
    if (!lastResult || !lastResult.lineas || !lastResult.allOk) {
      st.innerHTML = '<div class="status-msg bad">Primero procesá el export y dejá la validación en verde.</div>';
      return;
    }
    btnRegistrarB.disabled = true;
    st.innerHTML = '<div class="status-msg">Guardando los saldos…</div>';
    try {
      estadoB = await guardarEstadoB(estadoB, periodo, lastResult.lineas);
      renderEstadoB();
      st.innerHTML = '<div class="status-msg ok">Listo: ' + periodo + ' quedó registrado. ' +
        'El mes que viene la columna "Saldo anterior" se completa sola.</div>';
    } catch (e) {
      st.innerHTML = '<div class="status-msg bad">' + e.message + '</div>';
    } finally {
      btnRegistrarB.disabled = false;
    }
  });
}

// --------------------------------------------------------------------
// Arranque
// --------------------------------------------------------------------

loadMapping();
cargarEstadoB();
