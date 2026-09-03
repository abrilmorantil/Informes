// Panel de configuración de cuentas — Pesos.
//
// Dos cosas distintas, una arriba de la otra:
//   1. Pendientes (config_balances.js): lo que el maestro tiene ambiguo o mal puesto — dos
//      cuentas escritas en la misma fila, una cuenta en una columna que la fórmula no lee.
//      Aparte van las líneas SIN cuenta asignada, que no son un error sino un estado válido
//      y traen un campo para asignarles una cuenta de Onvio.
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

// Sólo lo que pide una acción.
//
// Antes había cuatro contadores —líneas, categorías, códigos distintos, pendientes— y tres de
// ellos no le decían a nadie qué hacer: son números que están bien cualesquiera sean. Al
// ponerlos al lado del único que importa, lo diluían. Queda el de pendientes, que es el que
// dice si hay algo para mirar; el detalle de cada uno va abajo.
function pccResumenHtml(cfg, pendientesTotal, categoriasN) {
  const hay = pendientesTotal > 0;
  return `
    <div class="checks" style="margin-top:16px;">
      <div class="check ${hay ? "bad" : "ok"}">
        <div>
          <div class="name">Pendientes de revisar</div>
          <div class="detail">${hay
            ? "Líneas del balance que no está claro de qué cuenta salen."
            : "Cada línea del balance sabe de qué cuenta de Onvio sale."}</div>
        </div>
        <div class="value">${pendientesTotal}</div>
      </div>
    </div>`;
}

function pccPendientesHtml(cfg) {
  const bloques = [];
  for (const a of cfg.avisos) {
    if (a.tipo === "dos_cuentas_en_la_misma_fila") {
      bloques.push(`
        <div class="status-msg bad">
          <b>La fila ${a.fila} de Distribución por línea tiene dos cuentas escritas</b> — no está claro cuál
          alimenta esta línea del balance:
          <br>· <b>${a.usa.code}</b> — ${a.usa.description}
          <br>· <b>${a.tambien.code}</b> — ${a.tambien.description} (columna ${a.tambien.columna})
        </div>`);
    } else if (a.tipo === "cuenta_en_la_columna_equivocada") {
      // Se ofrece el arreglo con los datos ya puestos, pero EDITABLES. No se copia y listo
      // porque el código de la otra columna puede ser el viejo del cliente y no existir en
      // Onvio: es el caso de "12301000 - Titulos Públicos", cuyo código real es 123010000.
      // Copiarlo tal cual dejaría la cuenta igual de muda, y encima pareciendo arreglada.
      const c0 = a.cuentas[0];
      bloques.push(`
        <div class="status-msg bad">
          <b>Nota 4, fila ${a.nota4}: "${a.texto}"</b> nunca puede traer un importe: en la fila
          ${a.saldos} de Distribución por línea la cuenta está escrita en la columna
          ${a.cuentas.map(c => `<b>${c.columna}</b> (${c.code} — ${c.description})`).join(" y ")},
          pero la fórmula lee otra columna, que está vacía.
          ${a.cuentas.length === 1 ? `
          <div class="gc-form" style="margin-top:10px;">
            <span style="font-size:12px;">Ponerla en la columna que la fórmula lee:</span>
            <input type="text" class="gc-cod-input" id="cc_cod_${a.saldos}" value="${gcEsc(c0.code)}">
            <input type="text" id="cc_nom_${a.saldos}" value="${gcEsc(c0.description)}">
            <button data-accion="cc-corregir" data-fila="${a.saldos}">Corregir</button>
            <p class="footer-note">Revisá el código: si el que figura es el viejo del plan del
              cliente, la cuenta va a seguir sin traer importe. Tiene que ser el de Onvio.</p>
          </div>` : ""}
        </div>`);
    } else if (a.tipo === "linea_sin_cuenta") {
      continue;   // no es un pendiente: ver pccSinCuentaHtml
    } else if (a.tipo === "codigo_repetido") {
      bloques.push(`
        <div class="status-msg bad">
          <b>Código repetido ${a.code}</b> en Distribución por línea — filas ${a.filaPrevia} y ${a.fila}.
        </div>`);
    }
  }
  return bloques.length ? `<div style="margin-top:14px;">${bloques.join("")}</div>` : "";
}

// Las líneas del balance que hoy no tienen ninguna cuenta que las alimente.
//
// NO son un pendiente: que una línea esté prevista y todavía sin cuenta es un estado válido
// —"- Socios" de la Nota 4 es así— y marcarla en rojo todos los meses enseña a ignorar el
// panel. Lo que sí hace falta es que se vea y que se pueda resolver sin abrir el Excel, así
// que cada una viene con la lista de cuentas de Onvio para elegirle la suya.
//
// El código se ESCRIBE, con sugerencias. La tentación era un desplegable cerrado con las
// cuentas de `Hoja1`, pero Hoja1 no es el plan de Onvio: la corrida sólo le agrega las cuentas
// que el balance ya tiene una fila para recibir. Medido en el maestro de pesos, Hoja1 tiene 243
// cuentas contra las ~2.900 del export, y de esas 243 quedan 10 sin usar, 9 de ellas
// proveedores. Un desplegable con eso adentro no deja elegir casi nada y encima sugiere que
// eso es todo lo que Onvio manda.
//
// Entonces: campo libre para el código, con las libres de Hoja1 como sugerencia (esas son las
// seguras, porque el VLOOKUP busca contra el texto de Hoja1 y ahí ya está), y el nombre se
// completa solo cuando el código es una de ellas. Es la misma forma que ya tiene "corregir"
// más arriba, así que el panel no estrena una manera distinta de hacer lo mismo.
function pccSinCuentaHtml(cfg, cuentasOnvio) {
  const lineas = (cfg.avisos || []).filter(a => a.tipo === "linea_sin_cuenta");
  if (!lineas.length) return "";
  const sug = (cuentasOnvio || []).map(c =>
    `<option value="${gcEsc(c.codigo)}">${gcEsc(c.nombre)}</option>`).join("");
  return `
    <div style="margin-top:14px;">
      <div class="gc-tit" style="margin-bottom:6px;">Líneas sin cuenta asignada</div>
      <p class="footer-note" style="margin-top:0;">Están en el balance y hoy salen en cero
        porque no hay ninguna cuenta de Onvio que las alimente. Se les puede asignar una.</p>
      <datalist id="scSugeridas">${sug}</datalist>
      ${lineas.map(a => `
        <div class="status-msg">
          <b>Nota 4, fila ${a.nota4}: "${gcEsc(a.texto)}"</b> — sale en cero.
          <div class="gc-form" style="margin-top:10px;">
            <span style="font-size:12px;">Asignarle una cuenta de Onvio:</span>
            <input type="text" class="gc-cod-input sc-cod" id="cc_cod_${a.saldos}"
                   list="scSugeridas" placeholder="código" data-fila="${a.saldos}">
            <input type="text" id="cc_nom_${a.saldos}" placeholder="nombre, igual que en Onvio">
            <button data-accion="cc-corregir" data-fila="${a.saldos}">Asignar</button>
          </div>
          <p class="footer-note">El código y el nombre tienen que ser los de Onvio, letra por
            letra: el balance busca la cuenta por ese texto. Si el código es uno de los
            sugeridos, el nombre se completa solo y no hay con qué errarle.</p>
        </div>`).join("")}
    </div>`;
}

// Las cuentas de Onvio, leídas de Hoja1 del propio maestro. La clave de Hoja1 es el texto
// completo "código - nombre", que es exactamente lo que hay que escribir en SALDOS.
//
// Se dejan afuera las que YA están puestas en otra línea del balance. Sin ese filtro la lista
// ofrecía "111010001 - Caja", que ya alimenta su propia línea: asignarla a una segunda la
// contaba dos veces y el balance se abría. Probado: pasaba de 0 pendientes a 2.
function pccCuentasDeHoja1(wb, mapeo) {
  const ws = wb && hojaSumas(wb);
  if (!ws) return [];
  const yaPuestas = new Set(Object.keys((mapeo && mapeo.cuentas) || {}).map(String));
  const out = [];
  const vistas = new Set();
  const texto = (v) => {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
      if (v.richText) return v.richText.map(t => t.text).join("");
      if (v.result !== undefined) return String(v.result);
      return "";
    }
    return String(v);
  };
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r < 2) return;
    const t = texto(row.getCell(1).value).trim();
    const m = /^\s*(\d{5,})\s*-\s*(.+?)\s*$/.exec(t);
    if (!m || vistas.has(m[1])) return;
    vistas.add(m[1]);
    if (yaPuestas.has(m[1])) return;      // ya alimenta otra línea del balance
    out.push({ codigo: m[1], nombre: m[2] });
  });
  out.sort((a, b) => a.codigo.localeCompare(b.codigo));
  return out;
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
let pccEditandoCat = null;     // filaMadre de la categoría cuyo rótulo se está editando
let pccMoviendo = null;        // fila de la cuenta suelta que se está mandando a una categoría
let pccAnexo = null;           // {lineas, conceptos, columnas, verificacion} del Anexo II
let pccOnvio = [];             // las cuentas de Onvio (Hoja1), para asignarlas a una línea

function pccLog(msg) {
  pccCambios.push(msg);
  const el = $("gcLog");
  el.textContent = pccCambios.join("\n");
  mostrar("gcLog", true);
  mostrar("gcGuardarBar", true);
  $("gcGuardarEstado").textContent = `${pccCambios.length} cambio(s) sin guardar.`;
}

function pccRecalcular() {
  pccConceptos = null;          // el mapa de "dónde se imprime" se rehace con el archivo
  pccMapeo = derivarMapeoMaestro(pccWb, "pesos", clasificacion);
  const lineasNota4 = lineasDeNota4(pccWb);
  pccCfg = derivarConfigBalance(pccWb, "pesos", pccMapeo, lineasNota4, PARAMS.pesos, { filasQueAgrega });
  const g = categoriasPesos(pccWb, pccMapeo);
  pccCategorias = g.categorias;
  pccSueltas = g.sueltas;

  pccOnvio = pccCuentasDeHoja1(pccWb, pccMapeo);

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

function gcCategoriaHtml(cat, cpt) {
  const abierta = pccAbiertas.has(cat.filaMadre);
  return `
    <details class="gc-cat" data-fila-madre="${cat.filaMadre}" ${abierta ? "open" : ""}>
      <summary>
        <span><span class="gc-tit">${gcEsc(cat.nombre)}</span><span class="gc-cod">${cat.codigo}</span></span>
        <span class="gc-n">${cat.miembros.length} cuenta${cat.miembros.length === 1 ? "" : "s"}</span>
      </summary>
      <div class="gc-body">
        ${gcRotuloHtml(cat, cpt)}
        ${cat.miembros.map(m => gcMiembroHtml(cat, m)).join("")}
        ${gcAgregarFormHtml(cat)}
      </div>
    </details>`;
}

// El rótulo de la categoría: el código del cliente y el nombre con el que se imprime.
// Cambiarlo NO toca las cuentas que la integran ni el subtotal — sólo cómo se llama.
function gcRotuloHtml(cat, cptDado) {
  // Con qué nombre sale impresa. NO es el rótulo de la hoja de trabajo: es el concepto del
  // Anexo II, que es el único lugar donde este nombre aparece en un estado. Mostrarlo acá es
  // lo que hace entendible que editar la categoría cambie las dos cosas — o que no pueda
  // cambiar el concepto, cuando lo comparte con otras categorías.
  const cpt = cptDado !== undefined ? cptDado : pccConceptoDe(cat.filaMadre);
  const impreso = cpt
    ? (cpt.compartidoCon.length
        ? `sale impresa dentro de <b>${gcEsc(cpt.concepto)}</b>, junto con ` +
          `${gcEsc(cpt.compartidoCon.join(" y "))}`
        : `sale impresa como <b>${gcEsc(cpt.concepto)}</b>`)
    : "no aparece en el Anexo II";

  if (pccEditandoCat !== cat.filaMadre) {
    return `
      <div class="gc-miembro gc-rotulo">
        <div class="gc-nom"><span class="cod">${cat.codigo}</span>${gcEsc(cat.nombre)}
          <span class="gc-compartida">código del cliente · ${impreso}</span></div>
        <div class="gc-acciones">
          <button class="secundario" data-accion="cat-editar-abrir"
                  data-madre="${cat.filaMadre}">Editar categoría</button>
        </div>
      </div>`;
  }
  const nota = cpt && cpt.compartidoCon.length
    ? `El concepto <b>${gcEsc(cpt.concepto)}</b> del Anexo II no se va a tocar: también lo arman
       ${gcEsc(cpt.compartidoCon.join(" y "))}, así que no es sólo de esta categoría. El nombre
       nuevo queda para identificarla acá.`
    : cpt
      ? `El nombre nuevo también se escribe en el Anexo II, que es donde sale impreso. Las
         cuentas de Onvio que la integran y el subtotal no se tocan.`
      : `OJO: esta categoría no la referencia ningún concepto del Anexo II, así que el nombre
         nuevo no se va a imprimir en ningún lado.`;
  return `
    <div class="gc-form">
      <input type="text" id="gcCatCod" class="gc-cod-input" value="${cat.codigo}"
             placeholder="código" size="12">
      <input type="text" id="gcCatNom" value="${gcEsc(cat.nombre)}" placeholder="nombre">
      <button data-accion="cat-editar-guardar" data-madre="${cat.filaMadre}">Guardar</button>
      <button class="secundario" data-accion="cat-editar-cancelar">Cancelar</button>
      <p class="footer-note" style="flex-basis:100%; margin:6px 0 0;">${nota}</p>
    </div>`;
}

// Dónde se imprime cada categoría: fila del Anexo II, concepto, y con qué otras categorías lo
// comparte. Es un Map por fila de la cuenta madre.
function pccMapaConceptos(wb) {
  const mapa = new Map();
  if (typeof a2Mapa !== "function" || !wb) return mapa;
  const { lineas } = a2Mapa(wb, madresResultados(wb, "pesos"));
  const porAnexo = new Map();
  for (const l of lineas) for (const d of (l.donde || [])) {
    if (!porAnexo.has(d.anexoFila)) porAnexo.set(d.anexoFila, []);
    porAnexo.get(d.anexoFila).push(l);
  }
  for (const l of lineas) {
    const d = (l.donde || [])[0];
    if (!d) continue;
    mapa.set(l.filaSaldos, {
      fila: d.anexoFila, concepto: d.concepto,
      compartidoCon: (porAnexo.get(d.anexoFila) || [])
        .filter(x => x.filaSaldos !== l.filaSaldos).map(x => x.nombre),
    });
  }
  return mapa;
}

// Se calcula una vez por render y no una vez por categoría: a2Mapa recorre el Anexo II entero
// y hacerlo 58 veces cuesta.
let pccConceptos = null;
function pccConceptoDe(filaMadre) {
  if (!pccConceptos) pccConceptos = pccMapaConceptos(pccWb);
  return pccConceptos.get(filaMadre) || null;
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
        ${sueltas.map(s => gcSueltaHtml(s)).join("")}
      </div>
    </details>`;
}

// Una cuenta suelta: la que no integra ninguna categoría. Se le puede dar una.
//
// El selector se arma sólo para la fila que se está moviendo, no para las 67 a la vez: con
// 58 categorías serían casi 4.000 opciones dibujadas para usar una.
function gcSueltaHtml(s) {
  if (pccMoviendo !== s.fila) {
    return `
      <div class="gc-miembro">
        <div class="gc-nom"><span class="cod">${s.codigo}</span>${gcEsc(s.nombre)}</div>
        <div class="gc-acciones">
          <button class="secundario" data-accion="mover-abrir" data-fila="${s.fila}">
            Poner en una categoría</button>
        </div>
      </div>`;
  }
  return `
    <div class="gc-form">
      <span class="gc-nom" style="flex-basis:100%;"><span class="cod">${s.codigo}</span>${gcEsc(s.nombre)}</span>
      <select id="gcMoverDestino">
        ${pccCategorias.map(c =>
          `<option value="${c.filaMadre}">${gcEsc(c.nombre)} (${c.codigo})</option>`).join("")}
      </select>
      <button data-accion="mover-guardar" data-fila="${s.fila}">Poner acá</button>
      <button class="secundario" data-accion="mover-cancelar">Cancelar</button>
      <p class="footer-note" style="flex-basis:100%; margin:6px 0 0;">
        La cuenta se muda al bloque de esa categoría y su importe pasa a sumar en el subtotal
        que lee el Anexo II. Si alguna fórmula lee su fila directamente, no se puede mover y se
        avisa sin tocar el archivo.
      </p>
    </div>`;
}

function gcRender() {
  // Una línea sin cuenta asignada no cuenta como pendiente: es un estado válido y tiene su
  // propio bloque, con la lista para asignarle una.
  const pendientesTotal = pccCfg.avisos.filter(a => a.tipo !== "linea_sin_cuenta").length;
  $("configCuentasResumen").innerHTML = pccResumenHtml(pccCfg, pendientesTotal, pccCategorias.length);
  $("configCuentasPendientes").innerHTML =
    pccPendientesHtml(pccCfg) + pccSinCuentaHtml(pccCfg, pccOnvio);

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
        <div class="gc-miembro a2-encabezado">
          <div class="gc-nom">Cuenta madre</div>
          <div class="gc-acciones">
            <span class="a2-rotulo a2-rot-concepto">Concepto del Anexo II</span>
            <span class="a2-rotulo a2-rot-tipo">Tipo de gasto</span>
            <span class="a2-rotulo a2-rot-boton"></span>
          </div>
        </div>
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
    <h3 style="font-family:'Newsreader',Georgia,serif; font-weight:500; font-size:20px; margin:26px 0 4px;">
      Anexo II — en qué columna cae cada gasto</h3>
    <p class="footer-note" style="margin-top:0;">
      El Anexo II lee el <b>subtotal de la cuenta madre</b>, nunca sus subcuentas: por eso una
      subcuenta nueva entra sola y acá no hay que tocar nada. Lo que se decide acá son dos
      cosas: el <b>concepto</b> —el renglón con el que sale impreso en el Anexo II— y el
      <b>tipo de gasto</b>, que es la columna donde cae: Administración, Comercialización,
      Exploración o Financieros.
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

  // Al escribir (o elegir) un código sugerido, se completa el nombre solo. El VLOOKUP del
  // balance busca la cuenta por el texto entero "código - nombre", así que un nombre tipeado
  // con una letra distinta deja la línea igual de muda, pero pareciendo arreglada.
  cont.addEventListener("input", (e) => {
    const campo = e.target.closest("input.sc-cod");
    if (!campo) return;
    const cuenta = pccOnvio.find(c => c.codigo === campo.value.trim());
    if (cuenta) $(`cc_nom_${campo.dataset.fila}`).value = cuenta.nombre;
  });

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
      } else if (accion === "cat-editar-abrir") {
        pccEditandoCat = filaMadre;
        pccAbiertas.add(filaMadre);
        gcRender();
        $("gcCatNom").focus();
      } else if (accion === "cat-editar-cancelar") {
        pccEditandoCat = null;
        gcRender();
      } else if (accion === "cat-editar-guardar") {
        const cat = pccCategorias.find(c => c.filaMadre === filaMadre);
        renombrarCategoria(pccWb, cat, $("gcCatCod").value.trim(), $("gcCatNom").value.trim(), pccLog);
        pccEditandoCat = null;
        pccRecalcular();
        gcRender();
      } else if (accion === "mover-abrir") {
        pccMoviendo = fila;
        gcRender();
      } else if (accion === "mover-cancelar") {
        pccMoviendo = null;
        gcRender();
      } else if (accion === "mover-guardar") {
        const destino = pccCategorias.find(c => c.filaMadre === +$("gcMoverDestino").value);
        if (!destino) throw new Error("Elegí una categoría.");
        moverCuentaACategoria(pccWb, pccMapeo, fila, destino, pccLog);
        pccMoviendo = null;
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
      } else if (accion === "cc-corregir") {
        const f = +btn.dataset.fila;
        const cod = $(`cc_cod_${f}`).value.trim();
        const nom = $(`cc_nom_${f}`).value.trim();
        if (!nom) throw new Error("El nombre no puede quedar vacío.");
        corregirColumnaDeCuenta(pccWb, f, cod, nom, pccLog);
        pccRecalcular();
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
  module.exports = { pccResumenHtml, pccPendientesHtml, pccSinCuentaHtml, pccCuentasDeHoja1,
                     gcMiembroHtml, gcCategoriaHtml, gcEsc,
                     gcSueltasHtml, gcSueltaHtml, gcRotuloHtml, pccMapaConceptos };
}
