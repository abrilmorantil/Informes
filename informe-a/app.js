// Une todo: estado guardado en GitHub, alta inicial, y el ciclo mensual
// (cargar → borrador → revisar en Excel → cerrar el mes).

const $ = (id) => document.getElementById(id);

let estado = null;        // { ultimo_mes_cerrado, historial }
let mapeoGuardado = null; // mapeo del archivo base
let bufferBase = null;    // el .xlsx base tal como está guardado
let periodoActual = null; // el que se está cargando ahora

let lineas = null;
let pendientes = [];
let wbBorrador = null;    // resultado de la corrida (todavía sin cerrar)
let mapeoBorrador = null;
let resumenBorrador = null;
let logLineas = [];

// archivo elegido en la alta inicial
let altaWb = null;
let altaBuffer = null;
let altaDeteccion = null;

const mostrar = (id, v) => $(id).classList.toggle("hidden", !v);
const log = (m) => { logLineas.push(String(m)); $("logBody").textContent = logLineas.join("\n"); };

// --------------------------------------------------------------- arranque

async function arrancar() {
  mostrar("cargando", true);
  try {
    if (!hasGhSettings()) {
      mostrar("cargando", false);
      mostrar("cardSinConfig", true);
      return;
    }
    const e = await leerEstado();
    if (!e) {
      mostrar("cargando", false);
      mostrar("cardAlta", true);
      return;
    }
    estado = e.estado;

    const m = await leerMapeo();
    const b = await leerBase();
    if (!m || !b) {
      throw new Error(
        "Hay un estado guardado pero falta el archivo base o el mapeo en el repositorio. " +
        "Revisá la carpeta configurada en ⚙."
      );
    }
    mapeoGuardado = m.mapeo;
    bufferBase = b.buffer;

    periodoActual = periodoSiguiente(estado.ultimo_mes_cerrado);
    $("txtUltimoCerrado").textContent = etiquetaPeriodo(estado.ultimo_mes_cerrado);
    $("txtPeriodoActual").textContent = etiquetaPeriodo(periodoActual);
    $("txtMesCerrar").textContent = etiquetaPeriodo(periodoActual);

    mostrar("cargando", false);
    mostrar("cardPeriodo", true);
    mostrar("cardAcumulados", true);
    renderAcumulados();
    if ($("txtReabrir") && estado.ultimo_mes_cerrado) {
      $("txtReabrir").textContent = etiquetaPeriodo(estado.ultimo_mes_cerrado);
    }
    mostrar("cardOnvio", true);
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
      <td>${etiquetaPeriodo(x.periodo)}</td>
      <td class="num">${x.lineas ?? "—"}</td>
      <td class="num">${x.nuevas ?? 0}</td>
      <td>${x.fecha ? new Date(x.fecha).toLocaleDateString("es-AR") : "—"}</td>
    </tr>`).join("");
  mostrar("cardHistorial", true);
}

// --------------------------------------------------------------- configuración

function abrirConfig() {
  const s = loadGhSettings();
  $("cfgToken").value = s.token || "";
  $("cfgRepo").value = s.repo || "";
  $("cfgRama").value = s.rama || "main";
  $("cfgCarpeta").value = s.carpeta || "informe-a";
  $("configStatus").innerHTML = "";
  mostrar("modalConfig", true);
}

function cerrarConfig() { mostrar("modalConfig", false); }

async function guardarConfig() {
  saveGhSettings({
    token: $("cfgToken").value.trim(),
    repo: $("cfgRepo").value.trim(),
    rama: $("cfgRama").value.trim() || "main",
    carpeta: $("cfgCarpeta").value.trim() || "informe-a",
  });
  const st = $("configStatus");
  st.innerHTML = '<div class="status-msg">Probando la conexión…</div>';
  try {
    await leerEstado();   // no importa si no existe todavía; importa que el repo conteste
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

// --------------------------------------------------------------- alta inicial

$("fileBase").addEventListener("change", async () => {
  const f = $("fileBase").files[0];
  if (!f) return;
  $("txtBase").textContent = f.name;
  $("altaStatus").innerHTML = "";
  try {
    altaBuffer = await f.arrayBuffer();
    altaWb = await abrirWorkbook(altaBuffer);

    const dist = altaWb.getWorksheet("Dist.de gastos");
    if (!dist) throw new Error("El archivo no tiene la hoja 'Dist.de gastos'.");
    const vivos = mesesVivos(dist);
    altaDeteccion = { vivos };
    renderDeteccionAlta(vivos);
    $("btnGuardarAlta").disabled = false;
  } catch (e) {
    $("altaStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnGuardarAlta").disabled = true;
  }
});

function renderDeteccionAlta(vivos) {
  const cont = $("altaDeteccion");
  const anio = new Date().getFullYear();
  if (vivos.length === 1) {
    const v = vivos[0];
    cont.innerHTML = `
      <div class="check ok" style="display:block;">
        <div class="name">Detecté que el mes en curso del archivo es ${v.nombre}</div>
        <div class="detail">Es la única columna que sigue al movimiento del mes.</div>
      </div>
      <div style="margin-top:14px; display:flex; gap:16px; align-items:flex-end; flex-wrap:wrap;">
        <div>
          <label style="font-size:13px; color:var(--ink-soft); display:block; margin-bottom:6px;">Año</label>
          <input type="text" id="altaAnio" value="${anio}" style="max-width:110px;">
        </div>
        <div style="flex:1; min-width:260px;">
          <label style="font-size:13px; color:var(--ink-soft); display:block; margin-bottom:6px;">
            ¿${v.nombre} ya está cerrado?
          </label>
          <select id="altaCerrado">
            <option value="si">Sí, ${v.nombre} ya está cerrado — el próximo que voy a cargar es el siguiente</option>
            <option value="no">No, ${v.nombre} todavía lo estoy trabajando — quiero cargarlo ahora</option>
          </select>
        </div>
      </div>`;
  } else if (vivos.length === 0) {
    cont.innerHTML = `
      <div class="check ok" style="display:block;">
        <div class="name">El archivo no tiene ningún mes abierto</div>
        <div class="detail">Todos los meses están cerrados con importes fijos. Decime cuál fue el último.</div>
      </div>
      <div style="margin-top:14px; display:flex; gap:16px; align-items:flex-end; flex-wrap:wrap;">
        <div>
          <label style="font-size:13px; color:var(--ink-soft); display:block; margin-bottom:6px;">Año</label>
          <input type="text" id="altaAnio" value="${anio}" style="max-width:110px;">
        </div>
        <div style="flex:1; min-width:220px;">
          <label style="font-size:13px; color:var(--ink-soft); display:block; margin-bottom:6px;">Último mes cerrado</label>
          <select id="altaUltimoMes">
            ${NOMBRES_MES.map((n, i) => `<option value="${i + 1}">${n}</option>`).join("")}
          </select>
        </div>
      </div>`;
  } else {
    cont.innerHTML = `
      <div class="check bad" style="display:block;">
        <div class="name">El archivo tiene más de un mes abierto: ${vivos.map(v => v.nombre).join(", ")}</div>
        <div class="detail">
          Solo puede haber uno. Cerrá en Excel los que correspondan (dejando el importe fijo en
          lugar de la fórmula) y volvé a subir el archivo.
        </div>
      </div>`;
  }
  mostrar("altaDeteccion", true);
}

$("btnGuardarAlta").addEventListener("click", async () => {
  const st = $("altaStatus");
  const vivos = altaDeteccion.vivos;
  if (vivos.length > 1) {
    st.innerHTML = '<div class="status-msg bad">Primero hay que dejar un solo mes abierto en el archivo.</div>';
    return;
  }
  const anio = parseInt(($("altaAnio") || {}).value, 10);
  if (!anio || anio < 2000 || anio > 2100) {
    st.innerHTML = '<div class="status-msg bad">Revisá el año.</div>';
    return;
  }

  mostrar("spinnerAlta", true);
  $("btnGuardarAlta").disabled = true;
  st.innerHTML = "";
  try {
    let ultimoCerrado;
    if (vivos.length === 1) {
      const mesVivo = vivos[0].mes;
      if ($("altaCerrado").value === "si") {
        // se cierra ahora, con los importes que ya calculó Excel
        aprobarMes({ wb: altaWb, periodo: formatearPeriodo({ anio, mes: mesVivo }), log: () => {} });
        ultimoCerrado = formatearPeriodo({ anio, mes: mesVivo });
      } else {
        // queda abierto: el próximo a cargar es ese mismo mes
        ultimoCerrado = mesVivo === 1
          ? formatearPeriodo({ anio: anio - 1, mes: 12 })
          : formatearPeriodo({ anio, mes: mesVivo - 1 });
      }
    } else {
      ultimoCerrado = formatearPeriodo({ anio, mes: parseInt($("altaUltimoMes").value, 10) });
    }

    const mapeo = derivarMapeo(altaWb);
    const nuevoEstado = { ultimo_mes_cerrado: ultimoCerrado, historial: [] };
    const buffer = await altaWb.xlsx.writeBuffer();

    await guardarTodo({
      bufferBase: buffer, mapeo, estado: nuevoEstado,
      mensaje: `Balance USD: carga inicial (último mes cerrado: ${etiquetaPeriodo(ultimoCerrado)})`,
    });

    st.innerHTML = '<div class="status-msg ok">Guardado. Ya podés cargar el mes.</div>';
    setTimeout(() => location.reload(), 900);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnGuardarAlta").disabled = false;
  } finally {
    mostrar("spinnerAlta", false);
  }
});

// --------------------------------------------------------------- ciclo mensual

$("fileOnvio").addEventListener("change", () => {
  const f = $("fileOnvio").files[0];
  if (f) $("txtOnvio").textContent = f.name;
  $("btnProcesar").disabled = !f;
});

async function leerLineasOnvio() {
  const buf = await $("fileOnvio").files[0].arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return parseOnvioExport(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }));
}

$("btnProcesar").addEventListener("click", async () => {
  ["cardPendientes", "cardResultado", "cardCierre"].forEach(id => mostrar(id, false));
  $("onvioStatus").innerHTML = "";
  mostrar("spinner", true);
  $("btnProcesar").disabled = true;
  try {
    lineas = await leerLineasOnvio();
    const det = detectarPendientes(lineas, mapeoGuardado);
    pendientes = det.pendientes;

    if (det.sinCc.length) {
      $("onvioStatus").innerHTML =
        `<div class="status-msg bad">Estos centros de costo del export no se pudieron identificar
         y sus líneas no se van a cargar: ${det.sinCc.join(", ")}</div>`;
    }

    if (pendientes.length) {
      renderPendientes();
      mostrar("cardPendientes", true);
    } else {
      await correrMotor({});
    }
  } catch (e) {
    $("onvioStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    mostrar("spinner", false);
    $("btnProcesar").disabled = false;
  }
});

function renderPendientes() {
  const cats = categoriasDisponibles(mapeoGuardado);
  $("pendientesBody").innerHTML = pendientes.map((p, i) => `
    <tr class="pending-row">
      <td>${p.codigo}</td>
      <td>${p.label}</td>
      <td>${p.cc_nombre}</td>
      <td class="num">${p.saldo.toFixed(2)}</td>
      <td>
        <select class="catSelect" data-idx="${i}">
          <option value="">— elegí una categoría —</option>
          ${cats.map(c => `<option value="${c}">${c}</option>`).join("")}
        </select>
      </td>
    </tr>`).join("");
  // son 95 categorías: sin buscador hay que recorrerlas a ojo
  conBuscadorTodos(".catSelect", "Buscar categoría…");
}

$("btnConfirmar").addEventListener("click", async () => {
  const selects = document.querySelectorAll(".catSelect");
  const elegidas = {};
  for (let i = 0; i < selects.length; i++) {
    if (!selects[i].value) {
      $("pendientesStatus").innerHTML =
        '<div class="status-msg bad">Falta elegir la categoría de al menos una cuenta.</div>';
      return;
    }
    elegidas[pendientes[i].codigo] = selects[i].value;
  }
  $("pendientesStatus").innerHTML = "";
  $("btnConfirmar").disabled = true;
  try {
    await correrMotor(elegidas);
    mostrar("cardPendientes", false);
  } catch (e) {
    $("pendientesStatus").innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  } finally {
    $("btnConfirmar").disabled = false;
  }
});

async function correrMotor(categoriasElegidas) {
  logLineas = [];
  // Se parte siempre de una copia limpia del archivo guardado: el workbook se
  // modifica en el lugar, así que reusarlo tras un intento fallido arrastraría
  // cambios a medio aplicar.
  const wb = await abrirWorkbook(bufferBase.slice(0));

  const { mapeo, resumen } = procesar({
    wb, lineas, mapeo: mapeoGuardado, categoriasElegidas, periodo: periodoActual, log,
  });

  wbBorrador = wb;
  mapeoBorrador = mapeo;
  resumenBorrador = resumen;

  renderResultado(resumen);
  mostrar("cardResultado", true);
  mostrar("cardCierre", true);
}

function renderResultado(resumen) {
  const items = [
    { name: "Total en USD de las cuentas de gasto", value: resumen.totalSaldo.toFixed(2), ok: true,
      detail: "Es la cifra que tiene que dar el balance. Excel la recalcula al abrir el archivo." },
    { name: "Líneas cargadas", value: String(resumen.lineas), ok: true,
      detail: "Cada línea del export quedó en la hoja SyS, con su centro de costo." },
    { name: "Cuentas ya conocidas", value: String(resumen.conocidas), ok: true,
      detail: "Se actualizaron en su lugar." },
    { name: "Cuentas nuevas insertadas", value: String(resumen.nuevas), ok: true,
      detail: "Se les creó la fila con sus fórmulas y se las referenció en Dist.de gastos." },
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
  if (resumen.fueraDelBalance && resumen.fueraDelBalance.length) {
    avisos.push(`
      <div class="check ok" style="display:block;">
        <div class="name">Quedaron afuera ${resumen.fueraDelBalance.length} cuenta(s) que no son de resultado</div>
        <div class="detail">
          El export trae cuentas patrimoniales que no van a este balance, que es solo de
          gastos (las cuentas que empiezan con 4):
          ${resumen.fueraDelBalance.map(c =>
            `<b>${c.codigo}</b> ${c.label} (${c.saldo.toFixed(2)})`).join(", ")}.
          No se cargaron ni se sumaron al total. Es lo esperado.
        </div>
      </div>`);
  }
  if (resumen.ccSinColumna && resumen.ccSinColumna.length) {
    avisos.push(`
      <div class="check bad" style="display:block;">
        <div class="name">Hay gastos que no van a aparecer en Dist.de gastos</div>
        <div class="detail">
          Estos centros de costo tienen movimiento este mes pero no tienen columna en esa hoja:
          ${resumen.ccSinColumna.map(c => `<b>${c.nombre}</b> (${c.saldo.toFixed(2)})`).join(", ")}.
          El importe igual queda en Sumas y Saldos.
        </div>
      </div>`);
  }
  $("avisosBody").innerHTML = avisos.join("");
}

// --------------------------------------------------------------- descarga y cierre

$("btnDescargar").addEventListener("click", async () => {
  if (!wbBorrador) return;
  const buffer = await wbBorrador.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `BALANCE_DE_COMPROBACION_USD_${periodoActual}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

$("fileAprobar").addEventListener("change", () => {
  const f = $("fileAprobar").files[0];
  if (f) $("txtAprobar").textContent = f.name;
  $("btnCerrarMes").disabled = !f;
});

$("btnCerrarMes").addEventListener("click", async () => {
  const st = $("cierreStatus");
  const f = $("fileAprobar").files[0];
  if (!f) return;

  const ok = confirm(
    `Vas a cerrar ${etiquetaPeriodo(periodoActual)}.\n\n` +
    `Los importes de ese mes quedan fijos y el archivo se guarda como definitivo. ` +
    `A partir de ahí el sistema pasa al mes siguiente.\n\n¿Confirmás?`
  );
  if (!ok) return;

  mostrar("spinnerCierre", true);
  $("btnCerrarMes").disabled = true;
  st.innerHTML = "";
  try {
    const buf = await f.arrayBuffer();
    const wb = await abrirWorkbook(buf);

    // el archivo que sube tiene que ser el de este mes, no otro
    const dist = wb.getWorksheet("Dist.de gastos");
    if (!dist) throw new Error("El archivo no tiene la hoja 'Dist.de gastos'.");
    const { mes } = parsearPeriodo(periodoActual);
    const vivos = mesesVivos(dist);
    if (!vivos.some(v => v.mes === mes)) {
      throw new Error(
        `Este archivo no tiene abierto ${nombreMes(mes)}` +
        (vivos.length ? `, tiene ${vivos.map(v => v.nombre).join(", ")}` : "") +
        `. Subí el borrador que descargaste recién. NO se guardó nada.`
      );
    }

    // el mapeo y las líneas hacen falta para pasar el mes al acumulado de "Gastos Acumulados"
    aprobarMes({ wb, periodo: periodoActual, mapeo: mapeoBorrador || mapeoGuardado, lineas, log });

    const nuevoEstado = {
      ultimo_mes_cerrado: periodoActual,
      historial: [...(estado.historial || []), {
        periodo: periodoActual,
        fecha: new Date().toISOString(),
        lineas: resumenBorrador ? resumenBorrador.lineas : null,
        nuevas: resumenBorrador ? resumenBorrador.nuevas : 0,
      }],
    };

    const buffer = await wb.xlsx.writeBuffer();
    await guardarTodo({
      bufferBase: buffer,
      mapeo: mapeoBorrador || mapeoGuardado,
      estado: nuevoEstado,
      mensaje: `Balance USD: cierra ${etiquetaPeriodo(periodoActual)}`,
    });

    st.innerHTML =
      `<div class="status-msg ok">${etiquetaPeriodo(periodoActual)} quedó cerrado y guardado.</div>`;
    setTimeout(() => location.reload(), 1200);
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    $("btnCerrarMes").disabled = false;
  } finally {
    mostrar("spinnerCierre", false);
  }
});

// ------------------------------------------------------------ reabrir un mes

// Deshace el último cierre para poder volver a cargar ese mes. Es una operación destructiva
// sobre el maestro, así que se confirma y se dice exactamente qué queda y qué no.
if ($("btnReabrir")) {
  $("btnReabrir").addEventListener("click", async () => {
    const st = $("reabrirStatus");
    const ultimo = estado && estado.ultimo_mes_cerrado;
    if (!ultimo) {
      st.innerHTML = '<div class="status-msg bad">No hay ningún mes cerrado para reabrir.</div>';
      return;
    }
    if (!confirm(
        `Reabrir ${etiquetaPeriodo(ultimo)}?\n\n` +
        `Su columna vuelve a seguir el movimiento y vas a poder cargar ese mes de nuevo.\n` +
        `Las cuentas dadas de alta en ese mes quedan como están.`)) return;

    $("btnReabrir").disabled = true;
    st.innerHTML = '<div class="status-msg">Reabriendo…</div>';
    try {
      const wb = await abrirWorkbook(bufferBase.slice(0));
      const salida = [];
      reabrirMes({ wb, periodo: ultimo, log: (t) => salida.push(t) });
      const buffer = await wb.xlsx.writeBuffer();

      const historial = (estado.historial || []).filter(h => h.periodo !== ultimo);
      const nuevoEstado = {
        ultimo_mes_cerrado: historial.length ? historial[historial.length - 1].periodo : null,
        historial,
      };
      await guardarTodo({
        bufferBase: buffer, mapeo: mapeoGuardado, estado: nuevoEstado,
        mensaje: `Se reabre ${ultimo} para volver a cargarlo`,
      });
      bufferBase = buffer;
      estado = nuevoEstado;
      st.innerHTML = `<div class="status-msg ok">${etiquetaPeriodo(ultimo)} quedó abierto de ` +
        `nuevo. Actualizá la página (Ctrl+Shift+R) y cargá su export.</div>`;
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    } finally {
      $("btnReabrir").disabled = false;
    }
  });
}

// ------------------------------------------------------------ Gastos Acumulados

// Esa hoja quedó fuera del sistema mucho tiempo, así que el maestro la tiene parada en un mes
// viejo. Se siembra una sola vez con el último informe terminado a mano y de ahí en adelante
// la mantiene `avanzarGastosAcumulados` en cada cierre.
let acumImportado = null;

async function renderAcumulados() {
  const caja = $("acumuladosBox");
  if (!caja || !bufferBase) return;
  try {
    const wb = await abrirWorkbook(bufferBase.slice(0));
    const d = leerGastosAcumulados(wb);
    if (d.error) { caja.innerHTML = `<div class="status-msg bad">${d.error}</div>`; return; }
    // El mes en curso de la hoja tiene que ser el que sigue al último cerrado. Puede estar
    // atrasado (le faltan meses) o adelantado (ya contó un mes que el balance no cerró);
    // antes las dos cosas se anunciaban como "atrasado", que es la mitad de las veces al revés.
    const esperado = estado && estado.ultimo_mes_cerrado
      ? parsearPeriodo(estado.ultimo_mes_cerrado).mes + 1 : null;
    const desfase = esperado === null ? 0 : d.mes - esperado;
    const detalle = desfase === 0 ? "" :
      desfase > 0
        ? ` Ojo: el balance tiene cerrado hasta ${etiquetaPeriodo(estado.ultimo_mes_cerrado)}, ` +
          `así que esta hoja va <b>${desfase} mes adelantada</b> y ya contó un mes que el ` +
          `balance todavía no cerró. Registrá ese informe acá abajo antes de seguir.`
        : ` Ojo: el balance tiene cerrado hasta ${etiquetaPeriodo(estado.ultimo_mes_cerrado)}, ` +
          `así que a esta hoja le faltan <b>${-desfase} mes(es)</b>.`;
    caja.innerHTML = `<div class="status-msg ${desfase === 0 ? "ok" : "bad"}">La hoja está en ` +
      `<b>${d.rotuloAcum}</b> con el mes en curso en <b>${d.rotuloMes}</b>.${detalle}</div>`;
  } catch (e) {
    caja.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
  }
}

if ($("fileAcum")) {
  $("fileAcum").addEventListener("change", async (ev) => {
    const archivo = ev.target.files[0];
    const st = $("importarAcumStatus");
    acumImportado = null;
    $("btnImportarAcum").disabled = true;
    if (!archivo) { $("txtAcum").textContent = "Elegí el informe terminado (.xlsx)"; st.innerHTML = ""; return; }
    $("txtAcum").textContent = archivo.name;
    st.innerHTML = '<div class="status-msg">Leyendo el archivo…</div>';
    try {
      const d = leerGastosAcumulados(await abrirWorkbook(await archivo.arrayBuffer()));
      if (d.error) { st.innerHTML = `<div class="status-msg bad">${d.error}</div>`; return; }
      // El archivo terminado no sólo trae el acumulado: es el mes ya hecho. Se registra
      // entero —pasa a ser el archivo base y ese mes queda cerrado— para que las tres cosas
      // (estado, Dist.de gastos y Gastos Acumulados) queden en el mismo mes. Sembrar sólo el
      // acumulado dejaba la hoja un mes adelantada y el cierre siguiente contaba doble.
      const wbSubido = await abrirWorkbook(await archivo.arrayBuffer());
      const vivos = mesesVivos(wbSubido.getWorksheet("Dist.de gastos"));
      if (vivos.length !== 1) {
        st.innerHTML = `<div class="status-msg bad">Este archivo tiene ` +
          `${vivos.length === 0 ? "todos los meses cerrados" : "más de un mes sin cerrar (" +
          vivos.map(m => m.nombre).join(", ") + ")"}, así que no puedo saber cuál registrar.</div>`;
        return;
      }
      if (vivos[0].mes !== d.mes) {
        st.innerHTML = `<div class="status-msg bad">El archivo no es coherente consigo mismo: ` +
          `"Gastos Acumulados" dice que el mes es <b>${d.rotuloMes}</b> pero en Dist.de gastos ` +
          `el mes sin cerrar es <b>${vivos[0].nombre}</b>.</div>`;
        return;
      }
      acumImportado = { datos: d, wb: wbSubido, mes: d.mes };
      const sig = MESES_IMP[d.mes] || "(cambio de ejercicio)";
      st.innerHTML = `<div class="status-msg ok">Es el informe de <b>${d.rotuloMes}</b> ` +
        `(${d.proyectos.length} proyectos). Al registrarlo: pasa a ser el archivo base, ` +
        `<b>${d.rotuloMes}</b> queda cerrado, el acumulado queda en ` +
        `<b>ENERO - ${d.rotuloMes.toUpperCase()}</b> y el próximo a cargar es <b>${sig}</b>, ` +
        `ya sólo con el export de Onvio.</div>`;
      $("btnImportarAcum").disabled = false;
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">No pude leer el archivo: ${e.message}</div>`;
    }
  });
}

if ($("btnImportarAcum")) {
  $("btnImportarAcum").addEventListener("click", async () => {
    const st = $("importarAcumStatus");
    if (!acumImportado) return;
    $("btnImportarAcum").disabled = true;
    st.innerHTML = '<div class="status-msg">Guardando…</div>';
    try {
      // el archivo terminado pasa a ser el maestro, con su mes ya cerrado: así el estado,
      // Dist.de gastos y Gastos Acumulados quedan los tres en el mismo mes
      const wb = acumImportado.wb;
      const anio = parsearPeriodo(estado.ultimo_mes_cerrado || "2026-01").anio;
      const periodo = formatearPeriodo({ anio, mes: acumImportado.mes });

      const rep = sembrarGastosAcumulados(wb, acumImportado.datos, anio);
      // sin mapeo ni líneas a propósito: el acumulado ya lo dejó puesto `sembrar`, acá sólo
      // hay que congelar la columna del mes con los importes que calculó Excel
      aprobarMes({ wb, periodo, log: () => {} });

      const mapeo = derivarMapeo(wb);
      const nuevoEstado = {
        ultimo_mes_cerrado: periodo,
        historial: [...(estado.historial || []),
                    { periodo, fecha: new Date().toISOString(), registrado: true }],
      };
      const buffer = await wb.xlsx.writeBuffer();
      await guardarTodo({
        bufferBase: buffer, mapeo, estado: nuevoEstado,
        mensaje: `Balance USD: se registra el informe terminado de ${etiquetaPeriodo(periodo)}`,
      });
      bufferBase = buffer;
      estado = nuevoEstado;
      mapeoGuardado = mapeo;
      await renderAcumulados();
      st.innerHTML = `<div class="status-msg ok">Registrado <b>${etiquetaPeriodo(periodo)}</b>. ` +
        `El acumulado quedó en <b>${rep.rotuloAcum}</b> y el próximo a cargar es ` +
        `<b>${rep.rotuloMes}</b>, ya sólo con el export.</div>` +
        (rep.sinCorrespondencia.length
          ? `<div class="status-msg">Sin correspondencia en el archivo que subiste: ` +
            `${rep.sinCorrespondencia.join(", ")}. Quedaron como estaban.</div>` : "") +
        (rep.anterioresCambiados.length
          ? `<div class="status-msg">Se actualizó además "años anteriores" en ` +
            `${rep.anterioresCambiados.length} proyecto(s), porque el archivo traía otro valor.</div>` : "") +
        (rep.total
          ? `<div class="status-msg">El total del acumulado (fila ${rep.total.fila}) tenía un ` +
            `número fijo${rep.total.previo !== null ? ` de ${rep.total.previo.toFixed(2)}` : ""} y ` +
            `ahora es <b>=${rep.total.formula}</b>` +
            `${rep.total.nuevo !== null ? `, que da <b>${rep.total.nuevo.toFixed(2)}</b>` : ""}.</div>`
          : `<div class="status-msg bad">No encontré la fila de totales de la hoja: revisá a mano ` +
            `que el total del acumulado sume la columna entera.</div>`);
    } catch (e) {
      st.innerHTML = `<div class="status-msg bad">${e.message}</div>`;
    } finally {
      $("btnImportarAcum").disabled = false;
    }
  });
}

arrancar();
