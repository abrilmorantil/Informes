// Une todo para el Ajuste por Conversión: subir el SyS, cargar los tipos de cambio del
// cierre, revisar lo que la salvaguarda marca, y recién ahí generar el importador.
//
// El flujo es el de la sección 12.3 de la especificación, en dos fases: primero se
// calcula y se separa lo dudoso SIN generar nada, y sólo después de que la usuaria
// decide se arma el asiento. Ninguna cuenta sospechosa entra sin confirmación.

const $ = (id) => document.getElementById(id);
const mostrar = (id, v) => $(id).classList.toggle("hidden", !v);

let config = null;        // la config persistente (config_difcambio.json)
let estado = { historial: [] };
let cuentasSys = null;    // lo parseado del SyS
let params = null;        // los parámetros del período
let lineasOk = [];
let lineasARevisar = [];
let asiento = null;       // el asiento final, ya con la línea de balanceo

const nf = (x, d = 2) => (typeof x === "number" ? x : 0)
  .toLocaleString("es-AR", { minimumFractionDigits: d, maximumFractionDigits: d });

// --------------------------------------------------------------- arranque

async function arrancar() {
  mostrar("cargando", true);
  try {
    if (!hasGhdSettings()) {
      mostrar("cargando", false);
      mostrar("cardSinConfig", true);
      return;
    }
    const guardada = await ghdLeerConfig();
    config = configDifCambio(guardada || {});
    if (!guardada) {
      // primera vez: se deja escrita la config de fábrica para que quede editable
      await ghdGuardarConfig(config, "Diferencia de cambio: configuración inicial");
    }
    estado = await ghdLeerEstado();

    const h = estado.historial || [];
    $("txtUltimo").textContent = h.length
      ? `Último asiento generado: ${h[h.length - 1].periodo} (${h[h.length - 1].lineas} líneas).`
      : "Todavía no se generó ningún asiento.";

    mostrar("cargando", false);
    mostrar("cardSys", true);
    renderHistorial();
  } catch (err) {
    mostrar("cargando", false);
    mostrar("cardSinConfig", true);
    $("cardSinConfig").innerHTML =
      `<h2>No pude leer la configuración</h2>
       <p class="footer-note">${err.message}</p>
       <div style="margin-top:14px;"><button onclick="abrirConfig()">Revisar configuración</button></div>`;
  }
}

function renderHistorial() {
  const h = (estado && estado.historial) || [];
  if (!h.length) return;
  $("historialBody").innerHTML = h.slice().reverse().map(x => `
    <tr>
      <td>${x.periodo}</td>
      <td class="num">${x.lineas}</td>
      <td class="num">${nf(x.balanceo)}</td>
      <td>${x.fecha ? new Date(x.fecha).toLocaleDateString("es-AR") : "—"}</td>
    </tr>`).join("");
  mostrar("cardHistorial", true);
}

// --------------------------------------------------------------- configuración

function abrirConfig() {
  const s = loadGhdSettings();
  $("cfgToken").value = s.token || "";
  $("cfgRepo").value = s.repo || "";
  $("cfgRama").value = s.rama || "main";
  $("cfgCarpeta").value = s.carpeta || "informe-d";
  $("configStatus").innerHTML = "";
  mostrar("modalConfig", true);
}
function cerrarConfig() { mostrar("modalConfig", false); }

async function guardarConfig() {
  saveGhdSettings({
    token: $("cfgToken").value.trim(),
    repo: $("cfgRepo").value.trim(),
    rama: $("cfgRama").value.trim() || "main",
    carpeta: $("cfgCarpeta").value.trim() || "informe-d",
  });
  const st = $("configStatus");
  st.innerHTML = '<div class="status-msg">Probando la conexión…</div>';
  try {
    await ghdLeerEstado();
    st.innerHTML = '<div class="status-msg ok">Conectado correctamente.</div>';
    setTimeout(() => { cerrarConfig(); location.reload(); }, 800);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

$("comoToken").addEventListener("click", (ev) => {
  ev.preventDefault();
  alert(
    "Cómo generar el token:\n\n" +
    "1. Andá a github.com → foto de perfil → Settings.\n" +
    "2. Developer settings → Personal access tokens → Tokens (classic).\n" +
    "3. Generate new token (classic).\n" +
    "4. Marcá el scope 'repo'.\n" +
    "5. Copiá el token (empieza con 'ghp_') — GitHub lo muestra una sola vez.\n\n" +
    "Queda guardado solo en este navegador."
  );
});

// --------------------------------------------------------------- 1. el SyS

$("fileSys").addEventListener("change", async () => {
  const f = $("fileSys").files[0];
  if (!f) return;
  $("txtSys").textContent = f.name;
  $("sysStatus").innerHTML = "";
  mostrar("cardParams", false);
  ["cardRevision", "cardResultado", "cardDescarga"].forEach(id => mostrar(id, false));
  try {
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", raw: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const r = parseSysBimonetario(filas, config.columnasSys);
    cuentasSys = r.cuentas;

    if (!cuentasSys.length) throw new Error("No reconocí ninguna cuenta. ¿Es el Balance de Sumas y Saldos?");
    const avisos = revisarParseoSys(cuentasSys);
    const porSeccion = {};
    cuentasSys.forEach(c => { porSeccion[c.seccion] = (porSeccion[c.seccion] || 0) + 1; });

    $("sysStatus").innerHTML =
      `<div class="status-msg ok">Leí ${cuentasSys.length} cuentas: ` +
      Object.entries(porSeccion).map(([s, n]) => `${n} de ${s}`).join(", ") + `.</div>` +
      avisos.map(a => `<div class="status-msg bad">${a}</div>`).join("");

    mostrar("cardParams", true);
  } catch (e) {
    cuentasSys = null;
    $("sysStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
});

// --------------------------------------------------------------- 2. cálculo (fase 1)

$("btnCalcular").addEventListener("click", async () => {
  const st = $("paramsStatus");
  st.innerHTML = "";
  const fecha = $("pFecha").value.trim();
  const compra = parseFloat($("pCompra").value.replace(",", "."));
  const venta = parseFloat($("pVenta").value.replace(",", "."));
  const asientoNro = parseInt($("pAsiento").value, 10);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
    st.innerHTML = '<div class="status-msg bad">La fecha de cierre va como 2026-06-30.</div>'; return;
  }
  if (!compra || compra <= 0 || !venta || venta <= 0) {
    st.innerHTML = '<div class="status-msg bad">Revisá los dos tipos de cambio.</div>'; return;
  }
  if (!cuentasSys) {
    st.innerHTML = '<div class="status-msg bad">Primero subí el Balance de Sumas y Saldos.</div>'; return;
  }

  mostrar("spinner", true);
  try {
    params = {
      periodoFin: fecha, tcCompra: compra, tcVenta: venta,
      numeroAsiento: isNaN(asientoNro) ? 1 : asientoNro, concepto: config.concepto,
    };
    const r = calcularConRevision(cuentasSys, params, config);
    lineasOk = r.lineasOk;
    lineasARevisar = r.lineasARevisar;

    if (lineasARevisar.length) {
      renderRevision();
      mostrar("cardRevision", true);
      mostrar("cardResultado", false);
      mostrar("cardDescarga", false);
    } else {
      mostrar("cardRevision", false);
      armarYMostrar(lineasOk);
    }
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    mostrar("spinner", false);
  }
});

// --------------------------------------------------------------- 3. revisión (fase 2)

function renderRevision() {
  $("revisionBody").innerHTML = lineasARevisar.map((l, i) => `
    <tr class="pending-row">
      <td>${l.codigo}</td>
      <td>${l.denominacion}</td>
      <td class="num">${nf(l.saldoPesos)}</td>
      <td class="num">${nf(l.usdLibros)}</td>
      <td class="num">${nf(l.usdTeorico)}</td>
      <td class="num">${nf(l.ajusteUsd)}</td>
      <td>
        <select class="selRevision" data-idx="${i}">
          <option value="">— elegí —</option>
          <option value="excluir">Excluir (es no monetaria)</option>
          <option value="incluir">Incluir (sí es monetaria)</option>
        </select>
      </td>
    </tr>`).join("");
}

$("btnConfirmarRevision").addEventListener("click", async () => {
  const selects = document.querySelectorAll(".selRevision");
  const decisiones = {};
  for (let i = 0; i < selects.length; i++) {
    if (!selects[i].value) {
      $("revisionStatus").innerHTML =
        '<div class="status-msg bad">Falta decidir al menos una cuenta.</div>';
      return;
    }
    decisiones[lineasARevisar[i].codigo] = selects[i].value;
  }
  $("revisionStatus").innerHTML = "";
  mostrar("spinnerRevision", true);
  $("btnConfirmarRevision").disabled = true;
  try {
    const { incluidas, cambiosConfig } = aplicarDecisiones(lineasARevisar, decisiones);
    config = aplicarCambiosConfig(config, cambiosConfig);
    await ghdGuardarConfig(config, "Diferencia de cambio: clasificación de cuentas confirmada");
    mostrar("cardRevision", false);
    armarYMostrar(lineasOk.concat(incluidas));
    $("descargaStatus").innerHTML =
      '<div class="status-msg ok">Clasificación guardada. No se vuelve a preguntar por esas cuentas.</div>';
  } catch (e) {
    $("revisionStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    mostrar("spinnerRevision", false);
    $("btnConfirmarRevision").disabled = false;
  }
});

// --------------------------------------------------------------- 4. resultado

function armarYMostrar(lineas) {
  asiento = armarAsiento(lineas, config);
  const { cierra, suma } = verificarCierre(asiento);
  const balanceo = asiento.find(l => l.esBalanceo);
  const publicadas = asiento.filter(l => !l.esBalanceo);

  const items = [
    { name: "Líneas del asiento", value: String(asiento.length), ok: true,
      detail: `${publicadas.length} cuentas ajustadas más la de balanceo.` },
    { name: `Cuenta de balanceo (${config.cuentaBalanceo})`, value: nf(balanceo.ajusteUsd), ok: true,
      detail: `${config.denomBalanceo}. Absorbe el neto y el residuo de redondeo.` },
    { name: "El asiento cierra en 0", value: nf(suma), ok: cierra,
      detail: cierra
        ? "La suma de todas las líneas da exactamente 0, como tiene que ser."
        : "NO cierra. No se habilita la descarga: revisá los tipos de cambio y el SyS." },
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
  const excluidasHoy = (config.noMonetariasExactas || []).length;
  avisos.push(`
    <div class="check ok" style="display:block;">
      <div class="name">Cuentas que no se ajustan</div>
      <div class="detail">
        Quedan afuera por criterio contable las de los prefijos
        ${(config.prefijosNoMonetarios || []).map(p => `<b>${p}</b>`).join(", ")}
        (bienes de uso y cargos diferidos) y ${excluidasHoy} cuenta(s) puntuales
        marcadas como no monetarias. El Patrimonio Neto y los Resultados no entran.
      </div>
    </div>`);
  $("avisosBody").innerHTML = avisos.join("");

  $("calculoBody").innerHTML = asiento.map(l => `
    <tr${l.esBalanceo ? ' style="font-weight:600; background:#FCF8F1;"' : ""}>
      <td>${l.codigo}</td>
      <td>${l.denominacion}</td>
      <td>${l.seccion}</td>
      <td class="num">${l.tcAplicado ? nf(l.tcAplicado, 0) : "—"}</td>
      <td class="num">${l.esBalanceo ? "—" : nf(l.saldoPesos)}</td>
      <td class="num">${l.esBalanceo ? "—" : nf(l.usdTeorico)}</td>
      <td class="num">${l.esBalanceo ? "—" : nf(l.usdLibros)}</td>
      <td class="num">${nf(l.ajusteUsd)}</td>
    </tr>`).join("");

  mostrar("cardResultado", true);
  mostrar("cardDescarga", true);
  // La descarga sólo se habilita si el asiento cierra: es la red de contención final.
  $("btnDescargar").disabled = !cierra;
  $("txtHabilitado").textContent = cierra
    ? ""
    : "Deshabilitado porque el asiento no cierra en 0.";
}

// --------------------------------------------------------------- 5. descarga

$("btnDescargar").addEventListener("click", () => {
  if (!asiento) return;
  const datos = escribirImportador(asiento, params, "xls");
  const blob = new Blob([datos], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Importador_Ajuste_Conversion_${params.periodoFin}.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

$("btnRegistrar").addEventListener("click", async () => {
  if (!asiento) return;
  mostrar("spinnerRegistrar", true);
  $("btnRegistrar").disabled = true;
  try {
    const balanceo = asiento.find(l => l.esBalanceo);
    estado.historial = [...(estado.historial || []), {
      periodo: params.periodoFin,
      fecha: new Date().toISOString(),
      lineas: asiento.length,
      balanceo: balanceo ? balanceo.ajusteUsd : 0,
      tcCompra: params.tcCompra,
      tcVenta: params.tcVenta,
    }];
    await ghdGuardarEstado(estado, `Diferencia de cambio: asiento de ${params.periodoFin}`);
    $("descargaStatus").innerHTML = '<div class="status-msg ok">Registrado en el historial.</div>';
    renderHistorial();
  } catch (e) {
    $("descargaStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    mostrar("spinnerRegistrar", false);
    $("btnRegistrar").disabled = false;
  }
});

arrancar();
