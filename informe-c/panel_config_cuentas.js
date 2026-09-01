// Panel de configuración de cuentas — Pesos.
//
// Dos cosas distintas, una arriba de la otra:
//   1. Pendientes (config_balances.js): lo que el maestro tiene ambiguo o sin resolver —
//      dos cuentas escritas en la misma fila, líneas de la Nota 4 sin ninguna cuenta. Es
//      diagnóstico, no se edita acá.
//   2. Categorías (gestion_categorias.js): qué cuentas de Onvio arman cada categoría de
//      RESULTADOS (las que Anexo II reparte en Administración/Comercialización/Exploración/
//      Financieros). Ahí sí se edita: sacar una cuenta, agregar una nueva, corregir código o
//      nombre. Los cambios quedan en memoria hasta que se aprietan "Guardar cambios en GitHub".
//
// Lee y escribe el maestro que ya está en memoria (bufferBase, cargado en app.js).

const PCC_CAP_ABREV = {
  "ACTIVO": "Activo", "PASIVO": "Pasivo",
  "PATRIMONIO NETO": "Patrimonio neto", "RESULTADOS": "Resultados",
};

// ------------------------------------------------------------------ pendientes (lectura)

function pccResumenHtml(cfg, pendientesTotal, categoriasN) {
  const r = cfg.resumen;
  const items = [
    { name: "Líneas del balance", value: r.lineas },
    { name: "Categorías de gastos", value: categoriasN },
    { name: "Código distinto para el cliente", value: r.conCliente },
    { name: "Pendientes de revisar", value: pendientesTotal, bad: pendientesTotal > 0 },
  ];
  return `<div class="checks" style="grid-template-columns:repeat(4,1fr); margin-top:16px;">` +
    items.map(it => `
      <div class="check ${it.bad ? "bad" : "ok"}">
        <div class="name">${it.name}</div>
        <div class="value">${it.value}</div>
      </div>`).join("") +
    `</div>`;
}

function pccPendientesHtml(cfg) {
  const bloques = [];
  for (const a of cfg.avisos) {
    if (a.tipo === "dos_cuentas_en_la_misma_fila") {
      bloques.push(`
        <div class="status-msg bad">
          <b>Fila ${a.fila} de SALDOS tiene dos cuentas escritas</b> — no está claro cuál
          alimenta esta línea del balance:
          <br>· <b>${a.usa.code}</b> — ${a.usa.description}
          <br>· <b>${a.tambien.code}</b> — ${a.tambien.description} (columna ${a.tambien.columna})
        </div>`);
    } else if (a.tipo === "cuenta_en_la_columna_equivocada") {
      bloques.push(`
        <div class="status-msg bad">
          <b>Nota 4, fila ${a.nota4}: "${a.texto}"</b> nunca puede traer un importe: en la fila
          ${a.saldos} de SALDOS la cuenta está escrita en la columna
          ${a.cuentas.map(c => `<b>${c.columna}</b> (${c.code} — ${c.description})`).join(" y ")},
          pero la fórmula lee otra columna, que está vacía. Se arregla moviendo el texto de
          columna en el Excel.
        </div>`);
    } else if (a.tipo === "linea_sin_cuenta") {
      bloques.push(`
        <div class="status-msg bad">
          <b>Nota 4, fila ${a.nota4}: "${a.texto}"</b> no tiene ninguna cuenta que la
          alimente hoy — siempre sale en 0.
        </div>`);
    } else if (a.tipo === "codigo_repetido") {
      bloques.push(`
        <div class="status-msg bad">
          <b>Código repetido ${a.code}</b> en SALDOS — filas ${a.filaPrevia} y ${a.fila}.
        </div>`);
    }
  }
  return bloques.length ? `<div style="margin-top:14px;">${bloques.join("")}</div>` : "";
}

// -------------------------------------------------------------- estado del panel, en memoria

let pccWb = null;              // el maestro con los cambios ya aplicados, en memoria
let pccMapeo = null;
let pccCfg = null;             // pendientes (config_balances.js), sólo diagnóstico
let pccCategorias = [];
let pccSueltas = [];
let pccCambios = [];           // log de lo que se hizo esta sesión, para el commit
let pccAbiertas = new Set();   // qué categorías quedan desplegadas al re-renderizar
let pccEditando = null;        // {fila} de la cuenta en edición inline
let pccAgregando = null;       // filaMadre de la categoría con el form de "agregar" abierto
let pccAnexo = null;           // {lineas, conceptos, columnas, verificacion} del Anexo II

function pccLog(msg) {
  pccCambios.push(msg);
  const el = $("gcLog");
  el.textContent = pccCambios.join("\n");
  mostrar("gcLog", true);
  mostrar("gcGuardarBar", true);
  $("gcGuardarEstado").textContent = `${pccCambios.length} cambio(s) sin guardar.`;
}

function pccRecalcular() {
  pccMapeo = derivarMapeoMaestro(pccWb, "pesos", clasificacion);
  const lineasNota4 = lineasDeNota4(pccWb);
  pccCfg = derivarConfigBalance(pccWb, "pesos", pccMapeo, lineasNota4, PARAMS.pesos, { filasQueAgrega });
  const g = categoriasPesos(pccWb, pccMapeo);
  pccCategorias = g.categorias;
  pccSueltas = g.sueltas;

  const madres = madresResultados(pccWb, "pesos");
  pccAnexo = {
    lineas: a2Mapa(pccWb, madres).lineas,
    conceptos: a2Conceptos(pccWb),
    columnas: a2Columnas(pccWb),
    verificacion: a2Verificar(pccWb, madres),
    madres,
  };
}

async function verConfigCuentasPesos() {
  mostrar("cardConfigCuentas", true);
  $("cardConfigCuentas").scrollIntoView({ behavior: "smooth", block: "start" });
  if (pccWb) { gcRender(); return; }

  if (!bufferBase) {
    $("configCuentasResumen").innerHTML =
      `<div class="status-msg bad">Todavía no se cargó el maestro de pesos. Esperá a que termine de cargar la página.</div>`;
    return;
  }

  $("configCuentasResumen").innerHTML = `<p class="footer-note">Leyendo el maestro…</p>`;
  $("configCuentasPendientes").innerHTML = "";
  try {
    pccWb = await abrirWorkbook(bufferBase.slice(0));
    pccRecalcular();
    gcRender();
  } catch (err) {
    pccWb = null;
    $("configCuentasResumen").innerHTML = `<div class="status-msg bad">No pude armar la configuración: ${err.message}</div>`;
  }
}

function cerrarConfigCuentas() {
  mostrar("cardConfigCuentas", false);
}

// -------------------------------------------------------------------------- render

function gcEsc(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function gcMiembroHtml(cat, m) {
  if (pccEditando && pccEditando.fila === m.fila) {
    return `
      <div class="gc-miembro">
        <div class="gc-form" style="margin:0; flex:1;">
          <input type="text" class="gc-cod-input" id="gcEditCod" value="${gcEsc(m.codigo)}" placeholder="Código">
          <input type="text" id="gcEditNom" value="${gcEsc(m.nombre)}" placeholder="Nombre">
          <button data-accion="editar-guardar" data-fila="${m.fila}" data-col="${m.col}">Guardar</button>
          <button class="secundario" data-accion="editar-cancelar">Cancelar</button>
        </div>
      </div>`;
  }
  const puedeQuitar = !m.esFilaCompartida && cat.miembros.length > 1;
  const acciones = puedeQuitar
    ? `<button class="secundario" data-accion="editar-abrir" data-fila="${m.fila}">Editar</button>
       <button class="secundario" data-accion="quitar" data-madre="${cat.filaMadre}" data-fila="${m.fila}">Quitar</button>`
    : `<button class="secundario" data-accion="editar-abrir" data-fila="${m.fila}">Editar</button>
       <span class="gc-compartida">${m.esFilaCompartida ? "comparte fila con la categoría" : "única cuenta"}</span>`;
  return `
    <div class="gc-miembro">
      <div class="gc-nom"><span class="cod">${m.codigo}</span>${gcEsc(m.nombre)}</div>
      <div class="gc-acciones">${acciones}</div>
    </div>`;
}

function gcAgregarFormHtml(cat) {
  if (pccAgregando !== cat.filaMadre) {
    return `<button class="secundario" style="margin-top:10px;" data-accion="agregar-abrir" data-madre="${cat.filaMadre}">+ Agregar cuenta</button>`;
  }
  return `
    <div class="gc-form">
      <input type="text" class="gc-cod-input" id="gcNuevoCod" placeholder="Código Onvio (ej. 425940000)">
      <input type="text" id="gcNuevoNom" placeholder="Nombre de la cuenta">
      <button data-accion="agregar-guardar" data-madre="${cat.filaMadre}">Agregar</button>
      <button class="secundario" data-accion="agregar-cancelar">Cancelar</button>
      <span class="footer-note">Entra dentro de "${gcEsc(cat.nombre)}", como si se insertara a mano en Excel.</span>
    </div>`;
}

function gcCategoriaHtml(cat) {
  const abierta = pccAbiertas.has(cat.filaMadre);
  return `
    <details class="gc-cat" data-fila-madre="${cat.filaMadre}" ${abierta ? "open" : ""}>
      <summary>
        <span><span class="gc-tit">${gcEsc(cat.nombre)}</span><span class="gc-cod">${cat.codigo}</span></span>
        <span class="gc-n">${cat.miembros.length} cuenta${cat.miembros.length === 1 ? "" : "s"}</span>
      </summary>
      <div class="gc-body">
        ${cat.miembros.map(m => gcMiembroHtml(cat, m)).join("")}
        ${gcAgregarFormHtml(cat)}
      </div>
    </details>`;
}

function gcSueltasHtml(sueltas) {
  if (!sueltas.length) return "";
  return `
    <details class="gc-cat">
      <summary>
        <span><span class="gc-tit">Sin categoría</span><span class="gc-cod">cuentas de Resultados que no arman ninguna categoría</span></span>
        <span class="gc-n">${sueltas.length}</span>
      </summary>
      <div class="gc-body">
        ${sueltas.map(s => `
          <div class="gc-miembro">
            <div class="gc-nom"><span class="cod">${s.codigo}</span>${gcEsc(s.nombre)}</div>
          </div>`).join("")}
      </div>
    </details>`;
}

function gcRender() {
  const pendientesTotal = pccCfg.avisos.length;
  $("configCuentasResumen").innerHTML = pccResumenHtml(pccCfg, pendientesTotal, pccCategorias.length);
  $("configCuentasPendientes").innerHTML = pccPendientesHtml(pccCfg);

  const q = ($("gcBuscar").value || "").trim().toLowerCase();
  const filtrar = (cat) => !q ||
    cat.nombre.toLowerCase().includes(q) || cat.codigo.includes(q) ||
    cat.miembros.some(m => m.nombre.toLowerCase().includes(q) || m.codigo.includes(q));
  const categoriasFiltradas = pccCategorias.filter(filtrar);

  $("gcCategorias").innerHTML = categoriasFiltradas.length
    ? categoriasFiltradas.map(gcCategoriaHtml).join("")
    : `<p class="footer-note">No hay categorías que coincidan con la búsqueda.</p>`;
  $("gcSueltas").innerHTML = q ? "" : gcSueltasHtml(pccSueltas);
  $("gcAnexo2").innerHTML = a2BloqueHtml();
  a2SincronizarSelects();
}

// ------------------------------------------------------------------ Anexo II
//
// La otra mitad de la configuración, y la que decide de verdad: el Anexo II lee el SUBTOTAL de
// cada cuenta madre y lo pone en una de sus cuatro columnas (Administración, Comercialización,
// Exploración, Financieros). Ahí se define si un gasto es de administración o de exploración.
//
// El invariante: cada madre entra EXACTAMENTE UNA VEZ. Dos veces cuenta el gasto doble; cero lo
// hace desaparecer del estado de resultados. Y no lo detecta ningún control de suma —el balance
// sigue cerrando igual— así que se muestra siempre, no sólo cuando falla.
function a2EstadoHtml(v) {
  if (v.ok) {
    return `<div class="status-msg ok">Las ${v.total} cuentas madre entran al Anexo II
      <b>exactamente una vez</b>. Ninguna cuenta doble ni queda afuera.</div>`;
  }
  const partes = [];
  if (v.dobles.length) {
    partes.push(`<b>${v.dobles.length} cuenta(s) madre entran más de una vez</b>, así que su
      gasto se cuenta doble: ` + v.dobles.map(d =>
        `${gcEsc(d.nombre)} (${d.donde.map(x => gcEsc(x.columna)).join(" y ")})`).join("; "));
  }
  if (v.sinAnexo.length) {
    partes.push(`<b>${v.sinAnexo.length} cuenta(s) madre no entran a ninguna columna</b>, así que
      su gasto no aparece en el estado de resultados: ` +
      v.sinAnexo.map(d => gcEsc(d.nombre)).join(", "));
  }
  return `<div class="status-msg bad">${partes.join("<br>")}</div>`;
}

function a2BloqueHtml() {
  if (!pccAnexo) return "";
  const { lineas, conceptos, columnas, verificacion } = pccAnexo;

  const q = (($("gcBuscar") && $("gcBuscar").value) || "").trim().toLowerCase();
  const visibles = lineas.filter(l => !q ||
    l.nombre.toLowerCase().includes(q) || String(l.codigo).includes(q) ||
    l.donde.some(d => d.concepto.toLowerCase().includes(q)));

  // agrupado por columna, que es como se lee el Anexo II
  const porColumna = {};
  for (const l of visibles) {
    const col = l.donde.length ? l.donde[0].columna : "(sin columna)";
    (porColumna[col] = porColumna[col] || []).push(l);
  }

  const opcConcepto = conceptos.map(c =>
    `<option value="${c.fila}">${gcEsc(c.texto)}</option>`).join("");
  const opcColumna = columnas.map(c =>
    `<option value="${c.col}">${gcEsc(c.nombre)}</option>`).join("");

  const grupos = Object.entries(porColumna).map(([col, ls]) => `
    <details class="gc-cat" ${q ? "open" : ""}>
      <summary>
        <span><span class="gc-tit">${gcEsc(col)}</span></span>
        <span class="gc-n">${ls.length} cuenta${ls.length === 1 ? "" : "s"} madre</span>
      </summary>
      <div class="gc-body">
        ${ls.map(l => {
          const d = l.donde[0];
          return `
          <div class="gc-miembro">
            <div class="gc-nom">
              <span class="cod">${gcEsc(l.codigo)}</span>${gcEsc(l.nombre)}
              ${d ? `<div class="gc-compartida">en "${gcEsc(d.concepto)}"</div>` : ""}
            </div>
            <div class="gc-acciones">
              <select class="a2-concepto" data-fila="${l.filaSaldos}">${opcConcepto}</select>
              <select class="a2-columna" data-fila="${l.filaSaldos}">${opcColumna}</select>
              <button data-accion="a2-mover" data-fila="${l.filaSaldos}">Mover</button>
            </div>
          </div>`;
        }).join("")}
      </div>
    </details>`).join("");

  return `
    <h3 style="font-family:'Instrument Serif',serif; font-weight:400; font-size:20px; margin:26px 0 4px;">
      Anexo II — en qué columna cae cada gasto</h3>
    <p class="footer-note" style="margin-top:0;">
      El Anexo II lee el <b>subtotal de la cuenta madre</b>, nunca sus subcuentas: por eso una
      subcuenta nueva entra sola y acá no hay que tocar nada. Lo que se decide acá es
      <b>en qué columna</b> cae ese gasto.
    </p>
    ${a2EstadoHtml(verificacion)}
    <div style="margin-top:12px;">${grupos || '<p class="footer-note">Nada coincide con la búsqueda.</p>'}</div>`;
}

// Deja los desplegables mostrando dónde está HOY cada madre. Se hace después de pintar,
// porque poner `selected` en el HTML se pierde al re-renderizar.
function a2SincronizarSelects() {
  if (!pccAnexo) return;
  for (const l of pccAnexo.lineas) {
    const d = l.donde[0];
    if (!d) continue;
    const sc = document.querySelector(`.a2-concepto[data-fila="${l.filaSaldos}"]`);
    const sl = document.querySelector(`.a2-columna[data-fila="${l.filaSaldos}"]`);
    if (sc) sc.value = String(d.anexoFila);
    if (sl) sl.value = String(d.col);
    // 94 conceptos no se pueden recorrer a ojo: el mismo desplegable con buscador que usa el
    // resto del sitio. El de columnas son 4 y no lo necesita.
    if (sc && typeof conBuscador === "function") conBuscador(sc, "Buscar el concepto…");
  }
}

// -------------------------------------------------------------------------- interacción

// Guardia: en los tests de Node este archivo se carga como módulo, sin `document`.
if (typeof document !== "undefined") {
document.addEventListener("DOMContentLoaded", () => {
  const buscador = $("gcBuscar");
  if (buscador) buscador.addEventListener("input", () => { if (pccWb) gcRender(); });

  // Se escucha en la tarjeta entera y no sólo en la lista de categorías: el bloque del
  // Anexo II vive en otro contenedor y sus botones tienen que funcionar igual.
  const cont = $("cardConfigCuentas");
  if (!cont) return;

  // Guarda qué categorías estaban abiertas cuando el usuario las toggle a mano (no cuando
  // gcRender() las vuelve a abrir por estar en pccAbiertas).
  cont.addEventListener("toggle", (e) => {
    if (!e.target.classList.contains("gc-cat")) return;
    const fila = +e.target.dataset.filaMadre;
    if (!fila) return;
    if (e.target.open) pccAbiertas.add(fila); else pccAbiertas.delete(fila);
  }, true);

  cont.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-accion]");
    if (!btn) return;
    e.preventDefault();
    const accion = btn.dataset.accion;
    const filaMadre = btn.dataset.madre ? +btn.dataset.madre : null;
    const fila = btn.dataset.fila ? +btn.dataset.fila : null;

    try {
      if (accion === "quitar") {
        const cat = pccCategorias.find(c => c.filaMadre === filaMadre);
        quitarCuentaDeCategoria(pccWb, cat, fila, pccLog);
        pccRecalcular();
        gcRender();
      } else if (accion === "editar-abrir") {
        pccEditando = { fila };
        gcRender();
      } else if (accion === "editar-cancelar") {
        pccEditando = null;
        gcRender();
      } else if (accion === "editar-guardar") {
        const col = +btn.dataset.col;
        const cod = $("gcEditCod").value.trim();
        const nom = $("gcEditNom").value.trim();
        if (!nom) throw new Error("El nombre no puede quedar vacío.");
        editarCuenta(pccWb, fila, col, cod, nom, pccLog);
        pccEditando = null;
        pccRecalcular();
        gcRender();
      } else if (accion === "agregar-abrir") {
        pccAgregando = filaMadre;
        pccAbiertas.add(filaMadre);
        gcRender();
        $("gcNuevoCod").focus();
      } else if (accion === "agregar-cancelar") {
        pccAgregando = null;
        gcRender();
      } else if (accion === "a2-mover") {
        const f = +btn.dataset.fila;
        const sc = document.querySelector(`.a2-concepto[data-fila="${f}"]`);
        const sl = document.querySelector(`.a2-columna[data-fila="${f}"]`);
        a2Mover({
          wb: pccWb, madres: pccAnexo.madres, filaSaldos: f,
          anexoFilaDestino: +sc.value, colDestino: +sl.value, log: pccLog,
        });
        pccRecalcular();
        gcRender();
      } else if (accion === "agregar-guardar") {
        const cod = $("gcNuevoCod").value.trim();
        const nom = $("gcNuevoNom").value.trim();
        if (!/^\d{6,}$/.test(cod)) throw new Error(`"${cod}" no parece un código de cuenta válido (mínimo 6 dígitos).`);
        if (!nom) throw new Error("El nombre no puede quedar vacío.");
        const cat = pccCategorias.find(c => c.filaMadre === filaMadre);
        agregarCuentaACategoria(pccWb, pccMapeo, { codigo: cod, nombre: nom }, cat, pccLog);
        pccAgregando = null;
        pccAbiertas.add(filaMadre);
        pccRecalcular();
        gcRender();
      }
    } catch (err) {
      $("gcGuardarEstado").innerHTML = "";
      alert(err.message);
    }
  });
});
}

async function guardarCambiosCategorias() {
  if (!pccWb || !pccCambios.length) return;
  const boton = $("btnGcGuardar");
  boton.disabled = true;
  $("gcGuardarEstado").innerHTML = `<span class="spinner">Guardando en GitHub…</span>`;
  try {
    const buffer = await pccWb.xlsx.writeBuffer();
    const mensaje = "Configuración de cuentas (pesos): " + pccCambios.length + " cambio(s)";
    await ghcGuardarBase(buffer, mensaje);
    bufferBase = buffer;
    $("gcGuardarEstado").innerHTML = `<div class="status-msg ok" style="margin:0;">Guardado. La próxima corrida ya usa este maestro.</div>`;
    pccCambios = [];
    mostrar("gcGuardarBar", true);
  } catch (err) {
    $("gcGuardarEstado").innerHTML = `<div class="status-msg bad" style="margin:0;">No pude guardar: ${err.message}</div>`;
  } finally {
    boton.disabled = false;
  }
}

if (typeof module !== "undefined") {
  module.exports = { pccResumenHtml, pccPendientesHtml, gcMiembroHtml, gcCategoriaHtml, gcEsc };
}
