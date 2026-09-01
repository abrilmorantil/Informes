// La pantalla de los centros de costo que el balance no reconoce.
//
// El motor NO adivina a propósito: antes lo hacía y mandaba "Tanque Blanco" a TANQUE NEGRO,
// "Cerro Amarillo" a CERRO ABANICO y "La Voluntad" a LA HOYADA. El total del balance daba
// bien igual —la plata entra en algún lado— así que nada lo delataba. Por eso ahora se
// pregunta, y hay exactamente dos respuestas posibles:
//
//   1) Es uno de los que YA están, escrito distinto. Pasa seguido: el export trae
//      "Proyecto Lonco Vaca- Palenque" y el bloque se llama "LONCO VACA - PELENQUE" (se movió
//      el guion y dice PALENQUE con A). Se elige cuál es de un desplegable y queda declarado:
//      de ahí en más ese nombre resuelve solo.
//
//   2) Es un proyecto NUEVO. Se le arma el espacio igual que a los demás —su bloque
//      DEBE/HABER/SALDO en "Sumas y Saldos" y su columna en "Dist.de gastos"— con el nombre
//      TAL CUAL lo escribe Onvio, así el mes que viene no hay que declarar nada.
//
// Los nombres llevan prefijo `cc` por el ámbito global único que comparten los scripts.

let ccPendientes = [];        // [{nombre, cc_codigo, lineas, saldo}]
let ccElegido = {};           // nombre -> "existente" | "nuevo"

const ccFmt = (n) => (typeof n === "number" ? n : 0)
  .toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function ccEsc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Un id que se pueda meter en el HTML aunque el nombre tenga espacios, guiones o acentos.
function ccId(nombre) {
  let h = 0;
  for (const ch of String(nombre)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return "cc" + h.toString(36);
}

function ccElegir(id, modo) {
  const p = ccPendientes.find(x => ccId(x.nombre) === id);
  if (!p) return;
  ccElegido[p.nombre] = modo;
  ccRender();
}

function ccRender() {
  const caja = document.getElementById("centrosBody");
  if (!caja) return;
  if (!ccPendientes.length) {
    caja.innerHTML = "";
    mostrar("cardCentros", false);
    return;
  }

  const bloques = (mapeoGuardado.cc_blocks || [])
    .map(b => b.nombre_balance)
    .sort((a, b) => a.localeCompare(b));

  caja.innerHTML = ccPendientes.map(p => {
    const id = ccId(p.nombre);
    const modo = ccElegido[p.nombre] || "";
    const form = modo === "existente" ? `
      <div class="cc-form">
        <div class="campo">
          <span>¿Cuál de los del balance es?</span>
          <select id="sel_${id}">
            <option value="">— elegí el proyecto —</option>
            ${bloques.map(n => `<option value="${ccEsc(n)}">${ccEsc(n)}</option>`).join("")}
          </select>
        </div>
        <button onclick="ccDeclarar('${id}')">Es éste</button>
      </div>`
      : modo === "nuevo" ? `
      <div class="cc-form">
        <div class="campo">
          <span>Nombre en el balance (como lo escribe Onvio)</span>
          <input type="text" id="nom_${id}" value="${ccEsc(p.nombre)}">
        </div>
        <div class="campo">
          <span>Código de c. de costo</span>
          <input type="text" class="corto" id="cod_${id}" value="${ccEsc(p.cc_codigo || "")}">
        </div>
        <button onclick="ccAgregar('${id}')">Darlo de alta</button>
      </div>
      <p class="footer-note" style="margin-top:10px;">
        Se le arma el bloque de DEBE/HABER/SALDO en <b>Sumas y Saldos</b> y su columna en
        <b>Dist.de gastos</b>, entre los otros proyectos, y entra en la columna de totales.
      </p>`
      : "";

    return `
      <div class="cc-item">
        <div class="cc-nombre">Onvio lo escribe <code>${ccEsc(p.nombre)}</code></div>
        <div class="cc-datos">
          ${p.lineas} línea(s) · ${ccFmt(p.saldo)} USD sin cargar${
            p.cc_codigo ? ` · centro de costo ${ccEsc(p.cc_codigo)} en Onvio` : ""}
        </div>
        <div class="cc-opciones">
          <label><input type="radio" name="${id}" ${modo === "existente" ? "checked" : ""}
                 onchange="ccElegir('${id}','existente')"> Ya está en el balance, escrito distinto</label>
          <label><input type="radio" name="${id}" ${modo === "nuevo" ? "checked" : ""}
                 onchange="ccElegir('${id}','nuevo')"> Es un proyecto nuevo</label>
        </div>
        ${form}
      </div>`;
  }).join("");

  // El desplegable de proyectos usa el mismo buscador que el resto del sitio.
  for (const p of ccPendientes) {
    if ((ccElegido[p.nombre] || "") !== "existente") continue;
    const sel = document.getElementById(`sel_${ccId(p.nombre)}`);
    if (sel && typeof conBuscador === "function") conBuscador(sel, "Buscar el proyecto…");
  }
  mostrar("cardCentros", true);
}

function ccMostrar(sinCc) {
  ccPendientes = (sinCc || []).slice();
  ccElegido = {};
  const st = document.getElementById("centrosStatus");
  if (st) st.innerHTML = "";
  ccRender();
}

// ---------------------------------------------------------------- 1) es uno que ya está
async function ccDeclarar(id) {
  const p = ccPendientes.find(x => ccId(x.nombre) === id);
  const st = document.getElementById("centrosStatus");
  if (!p) return;
  const sel = document.getElementById(`sel_${id}`);
  const destino = sel ? sel.value : "";
  if (!destino) {
    st.innerHTML = `<div class="status-msg bad">Elegí cuál de los proyectos del balance es.</div>`;
    return;
  }
  try {
    declararEquivalenciaCc({
      mapeo: mapeoGuardado, nombreOnvio: p.nombre, nombreBalance: destino, log: () => {},
    });
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${ccEsc(e.message)}</div>`;
    return;
  }
  await ccGuardarYReprocesar(st,
    `Declarado: Onvio escribe "${p.nombre}" al proyecto "${destino}"`,
    `Balance USD: "${p.nombre}" es el centro de costo "${destino}"`);
}

// ---------------------------------------------------------------- 2) es uno nuevo
async function ccAgregar(id) {
  const p = ccPendientes.find(x => ccId(x.nombre) === id);
  const st = document.getElementById("centrosStatus");
  if (!p) return;
  const nombre = (document.getElementById(`nom_${id}`) || {}).value || "";
  const codigo = (document.getElementById(`cod_${id}`) || {}).value || "";
  if (!String(nombre).trim()) {
    st.innerHTML = `<div class="status-msg bad">Falta el nombre del proyecto.</div>`;
    return;
  }
  if (!/^\d+$/.test(String(codigo).trim())) {
    st.innerHTML = `<div class="status-msg bad">El código de centro de costo de Onvio tiene que
      ser un número. Es el que Onvio usa para ese proyecto, y es lo que el balance busca en la
      hoja SyS.</div>`;
    return;
  }
  if (!confirm(`Se va a dar de alta el proyecto "${String(nombre).trim()}" en el balance:\n\n` +
               `· su bloque DEBE/HABER/SALDO en "Sumas y Saldos"\n` +
               `· su columna en "Dist.de gastos", entre los otros proyectos\n` +
               `· y entra en la columna de TOTALES\n\n` +
               `Queda para siempre. ¿Sigo?`)) return;

  st.innerHTML = `<div class="status-msg">Armando el espacio del proyecto…</div>`;
  const salida = [];
  let wb;
  try {
    wb = await abrirWorkbook(bufferBase.slice(0));
    agregarCentroDeCosto({
      wb, mapeo: mapeoGuardado, nombreOnvio: nombre, ccCodigo: codigo,
      log: (m) => salida.push(String(m)),
    });
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">${ccEsc(e.message)}</div>`;
    return;
  }

  // El maestro cambió de estructura, así que hay que subirlo junto con el mapeo: si sólo se
  // guardara el mapeo, quedaría hablando de columnas que el archivo no tiene.
  try {
    const buf = await wb.xlsx.writeBuffer();
    await guardarTodo({
      bufferBase: buf, mapeo: mapeoGuardado, estado,
      mensaje: `Balance USD: centro de costo nuevo "${String(nombre).trim()}"`,
    });
    bufferBase = buf;
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">Se armó el bloque pero no se pudo guardar en
      GitHub: ${ccEsc(e.message)}. No se cargó nada.</div>`;
    return;
  }

  ccPendientes = ccPendientes.filter(x => x.nombre !== p.nombre);
  st.innerHTML = `<div class="status-msg ok">Listo. ${salida.join(" ").replace(/\s+/g, " ").trim()}</div>`;
  ccRender();
  await ccReprocesar();
}

// ---------------------------------------------------------------- guardar y volver a mirar
async function ccGuardarYReprocesar(st, aviso, mensajeCommit) {
  st.innerHTML = `<div class="status-msg">Guardando…</div>`;
  try {
    // El maestro no cambió: se manda el mismo buffer que ya estaba, para no tener dos caminos
    // de guardado que puedan quedar desincronizados.
    await guardarTodo({ bufferBase, mapeo: mapeoGuardado, estado, mensaje: mensajeCommit });
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">No se pudo guardar: ${ccEsc(e.message)}</div>`;
    return;
  }
  st.innerHTML = `<div class="status-msg ok">${ccEsc(aviso)}.</div>`;
  await ccReprocesar();
}

// Se vuelve a mirar el export con el mapeo nuevo: lo que se resolvió desaparece de la lista.
async function ccReprocesar() {
  if (typeof lineas === "undefined" || !lineas || !lineas.length) return;
  const det = detectarPendientes(lineas, mapeoGuardado);
  ccMostrar(det.sinCc);
  if (typeof onDetectarPendientes === "function") onDetectarPendientes(det);
}

if (typeof module !== "undefined") {
  module.exports = { ccId, ccFmt, ccEsc };
}
