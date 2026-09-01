// Panel "Configurar cuentas" — la vista maestra de qué fila ve el cliente y con qué
// cuentas reales se llena.
//
// Por qué existe: este informe habla DOS numeraciones a la vez. El cliente lee el plan de
// cuentas viejo y nosotras trabajamos con el que manda Onvio, y no hay ninguna regla que
// relacione uno con el otro — los códigos fueron reasignados, no reformateados (medido:
// la regla de "sacarle un dígito al medio" acierta 1 de 38). Así que la equivalencia se
// declara a mano, y hasta ahora vivía escrita en la columna A del Excel, a mano también.
//
// Por qué se muestra AGRUPADO y no como lista: el informe es jerárquico —54 cuentas madre
// con 216 subcuentas— y en una lista plana esa estructura desaparece. Los tres errores que
// aparecieron al armar todo esto eran de jerarquía (una cuenta en dólares suelta al lado
// de siete hermanas anidadas, y dos madres declaradas como cuenta simple). Agrupado se ven
// de una; en la lista plana hubo que salir a buscarlos.
//
// Los cambios se acumulan en memoria y recién se suben al apretar "Guardar cambios".

let cbMapping = null;       // copia de trabajo del mapping
let cbCambios = [];         // qué se hizo, para el mensaje del commit
let cbBusqueda = "";
let cbEditando = null;      // código de la fila que se está editando
let cbGruposAbiertos = {};  // capítulo -> abierto
let cbHijasAbiertas = {};   // código de la madre -> desplegada
let cbAgregandoHija = null; // código de la madre a la que se le está agregando una subcuenta
let cbAgregandoFila = false;
let cbSinAsignarAbierto = false;  // el bloque de lo que quedó sin asignar, arriba de todo
let cbAsignando = null;           // código de la cuenta de Onvio que se está ubicando

const CB_CAPITULOS = ["ACTIVO", "PASIVO", "CAPITAL Y PATRIMONIO", "RESULTADOS"];

function cbEsc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// El código y el nombre que ve el cliente. Si la fila no declara `cliente`, las dos
// numeraciones coinciden y se muestra la propia.
function cbCliente(e) {
  return (e && e.cliente) ? e.cliente : { code: e.code, description: e.description };
}

// Con qué cuentas reales de Onvio se llena la fila.
function cbFuentes(e) {
  if (e.type === "parent") return (e.children || []).map(h => ({ code: h.code, nom: h.description }));
  if (e.type === "range") return [{ code: "", nom: `todas las cuentas que empiezan con ${e.prefix}` }];
  if (e.sin_cuentas) return [];
  return [{ code: e.code, nom: e.description }];
}

// Las cuentas reales que se pueden asignar. Salen del export que se haya procesado en
// esta sesion; si todavia no se proceso ninguno, de las que ya estan en el mapping. Se
// marca cuales ya tienen dueño, que es la unica forma de ver cuales faltan ubicar.
function cbCuentasDisponibles() {
  const dueño = {};
  for (const e of cbMapping) {
    if (e.type === "parent") for (const h of (e.children || [])) dueño[h.code] = cbCliente(e).code;
    else if (e.type !== "range" && !e.sin_cuentas) dueño[e.code] = cbCliente(e).code;
  }
  const cuentas = {};
  // Del export, si hay uno procesado: son las cuentas que Onvio manda de verdad.
  if (typeof lastResult === "object" && lastResult && lastResult.cuentas) {
    for (const [cod, c] of Object.entries(lastResult.cuentas)) cuentas[cod] = c.descripcion;
  }
  // Y las que ya estan declaradas, aunque este mes no hayan venido.
  //
  // Las filas "sin cuentas asignadas" NO cuentan: su codigo no es una cuenta de Onvio. Si se
  // las metiera acá, "421170000 Gastos Legales" aparecería como una cuenta de Onvio a ubicar
  // y encima con el nombre del renglón, cuando en Onvio ese número es "Alojamiento Rel.
  // Comunitarias Catamarca". Si esa cuenta existe de verdad, ya viene en el export con su
  // nombre bueno.
  for (const e of cbMapping) {
    if (e.type === "parent") for (const h of (e.children || [])) cuentas[h.code] = cuentas[h.code] || h.description;
    else if (e.type !== "range" && !e.sin_cuentas) cuentas[e.code] = cuentas[e.code] || e.description;
  }
  return Object.entries(cuentas)
    .map(([code, nom]) => ({ code, nom, dueño: dueño[code] || null }))
    .sort((a, b) => (a.dueño ? 1 : 0) - (b.dueño ? 1 : 0) || a.code.localeCompare(b.code));
}

// ------------------------------------------------------------------ lo que quedó sin asignar
//
// Las dos puntas sueltas entre los dos planes de cuentas, que son problemas distintos:
//   - una cuenta de Onvio que no está en ninguna fila: su plata no entra al informe;
//   - una fila que el cliente ve y que no se llena con nada: se imprime siempre en cero.
//
// Lo importante es que muestra las cuentas QUIETAS. Las que tuvieron movimiento ya las reclama
// la pantalla de "cuentas nuevas" y trancan la descarga, así que se notan solas; las dormidas
// no rompen nada hasta el mes que se despiertan y hasta ahora no las mostraba nada. Pasó con
// "422130000 Gastos en trámites" y "424140000 Canon", dormidas todo el ejercicio.
function cbSinAsignar() {
  const mov = (typeof lastResult === "object" && lastResult && lastResult.cuentas) || {};
  const cuentas = cbCuentasDisponibles().filter(c => !c.dueño).map(c => {
    const m = mov[c.code];
    return { code: c.code, nom: c.nom, movio: !!(m && (m.debe || m.haber)),
             debe: m ? m.debe : null, haber: m ? m.haber : null };
  });
  const filas = cbMapping.filter(e => !cbFuentes(e).length).map(e => ({
    code: e.code, cliente: cbCliente(e),
    motivo: e.sin_cuentas ? "declarada sin cuentas asignadas"
                          : "es cuenta madre y no tiene ninguna subcuenta",
  }));
  return { cuentas, filas };
}

function cbToggleSinAsignar() { cbSinAsignarAbierto = !cbSinAsignarAbierto; cbRender(); }
function cbAsignar(code) { cbAsignando = code; cbRender(); }

// Ubica una cuenta de Onvio en la fila que se elija. Si la fila era simple, pasa a ser madre
// y se lleva su propia cuenta como primera subcuenta: es la misma operación que ya se hacía
// a mano con el botón "convertir en cuenta madre".
function cbConfirmarAsignacion(code) {
  const sel = document.getElementById("cbSelDestino");
  const destino = sel ? sel.value : "";
  if (!destino) { alert("Elegí en qué fila va."); return; }
  const e = cbMapping.find(x => x.code === destino);
  if (!e) return;
  const disp = cbCuentasDisponibles().find(c => c.code === code);
  const nombre = disp ? disp.nom : code;
  const cli = cbCliente(e);

  if (e.type === "range") { alert("Esa fila junta cuentas por prefijo; no se le agregan una por una."); return; }
  if (e.type !== "parent") {
    if (!confirm(`"${cli.code} ${cli.description}" es una fila simple.\n\n` +
                 `Para poder meterle "${code} ${nombre}" tiene que pasar a ser cuenta madre` +
                 (e.sin_cuentas ? "." : `, y su cuenta "${e.code} ${e.description}" queda como primera subcuenta.`) +
                 `\n\n¿Sigo?`)) return;
    const hijas = e.sin_cuentas ? [] : [{ code: e.code, description: e.description }];
    const i = cbMapping.indexOf(e);
    cbMapping[i] = { code: cli.code, aliases: [], description: cli.description,
                     category: e.category, type: "parent", orden: e.orden, children: hijas };
    if (e.ocultar_si_cero) cbMapping[i].ocultar_si_cero = true;
    cbCambios.push(`"${cli.code} ${cli.description}" pasa a ser cuenta madre`);
  }
  const madre = cbMapping.find(x => x.code === cli.code);
  madre.children = (madre.children || []).concat([{ code, description: nombre }]);
  cbCambios.push(`"${code} ${nombre}" entra en "${cli.code} ${cli.description}"`);
  cbHijasAbiertas[cli.code] = true;
  cbAsignando = null;
  cbRender();
}

function cbSinAsignarHtml() {
  const { cuentas, filas } = cbSinAsignar();
  if (!cuentas.length && !filas.length) {
    return `<div class="status-msg ok" style="margin-bottom:16px;">Todas las cuentas de Onvio
      están en alguna fila, y todas las filas se llenan con alguna cuenta. No quedó nada suelto.</div>`;
  }
  const cabecera = `
    <button class="cb-grupo-tit" onclick="cbToggleSinAsignar()">
      <span class="cb-chev">${cbSinAsignarAbierto ? "▾" : "▸"}</span>
      <span>Sin asignar</span>
      <span class="cb-cant">${cuentas.length} cuenta${cuentas.length === 1 ? "" : "s"} de Onvio ·
        ${filas.length} fila${filas.length === 1 ? "" : "s"} del cliente</span>
    </button>`;
  if (!cbSinAsignarAbierto) return `<div class="cb-grupo cb-sueltas">${cabecera}</div>`;

  const imp = (n) => n == null ? "—"
    : n.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // El desplegable de destino se arma una sola vez, con todas las filas del informe.
  const opciones = cbMapping
    .filter(e => e.type !== "range")
    .map(e => ({ v: e.code, t: `${cbCliente(e).code} — ${cbCliente(e).description}` }))
    .sort((a, b) => a.t.localeCompare(b.t));

  let html = cabecera + '<div class="cb-sueltas-cuerpo">';

  if (cuentas.length) {
    html += `<p class="cb-leyenda" style="margin:0 0 10px;">
      <strong>Cuentas de Onvio que no están en ninguna fila.</strong> Su importe no entra al
      informe. Las que dicen “movió” lo están dejando afuera <em>este mes</em>; las quietas no
      rompen nada todavía, pero el mes que tengan movimiento van a trancar la corrida.</p>`;
    html += cuentas.map(c => `
      <div class="cb-suelta">
        <div class="cb-suelta-id">
          <strong>${cbEsc(c.code)}</strong>
          <span>${cbEsc(c.nom)}</span>
        </div>
        <span class="cb-estado ${c.movio ? "cb-movio" : ""}">${
          c.movio ? `movió · ${imp(c.debe)} / ${imp(c.haber)}` : "sin movimiento este mes"}</span>
        ${cbAsignando === c.code ? `
          <div class="cb-suelta-form">
            <select id="cbSelDestino">
              <option value="">— ¿en qué fila va? —</option>
              ${opciones.map(o => `<option value="${cbEsc(o.v)}">${cbEsc(o.t)}</option>`).join("")}
            </select>
            <button onclick="cbConfirmarAsignacion('${cbEsc(c.code)}')">Asignar</button>
            <button class="secundario" onclick="cbAsignando=null; cbRender()">Cancelar</button>
          </div>`
        : `<button class="secundario" onclick="cbAsignar('${cbEsc(c.code)}')">Asignar a una fila</button>`}
      </div>`).join("");
  }

  if (filas.length) {
    html += `<p class="cb-leyenda" style="margin:16px 0 10px;">
      <strong>Filas que ve el cliente y no se llenan con ninguna cuenta de Onvio.</strong>
      Se imprimen siempre en cero. Para darles una cuenta, buscalas abajo y usá
      “+ Agregar subcuenta”.</p>`;
    html += filas.map(f => `
      <div class="cb-suelta">
        <div class="cb-suelta-id">
          <strong>${cbEsc(f.cliente.code)}</strong>
          <span>${cbEsc(f.cliente.description)}</span>
        </div>
        <span class="cb-estado">${cbEsc(f.motivo)}</span>
        <button class="secundario" onclick="cbBuscar('${cbEsc(f.cliente.code)}'); document.getElementById('cbBuscar').value='${cbEsc(f.cliente.code)}'">Ver la fila</button>
      </div>`).join("");
  }

  return `<div class="cb-grupo cb-sueltas">${html}</div></div>`;
}

function cbAbrirPanel() {
  if (!currentMapping) {
    document.getElementById("cbStatus").innerHTML =
      '<div class="status-msg bad">Todavía no se cargó el mapeo. Esperá a que termine de cargar la página.</div>';
    return;
  }
  cbMapping = JSON.parse(JSON.stringify(currentMapping));
  cbCambios = [];
  cbBusqueda = "";
  cbEditando = null;
  cbHijasAbiertas = {};
  cbAgregandoHija = null;
  cbAgregandoFila = false;
  cbAsignando = null;
  // Arranca abierto sólo si hay algo suelto: es lo primero que hay que ver, pero cuando no
  // hay nada no tiene por qué ocupar lugar.
  cbSinAsignarAbierto = false;
  // Arrancan cerrados: 206 filas desplegadas de golpe no se pueden leer.
  cbGruposAbiertos = {};
  document.getElementById("cardCuentas").classList.remove("hidden");
  cbRender();
  document.getElementById("cardCuentas").scrollIntoView({ behavior: "smooth", block: "start" });
}

function cbCerrar() {
  if (cbCambios.length && !confirm(`Hay ${cbCambios.length} cambio(s) sin guardar. ¿Cerrar igual y perderlos?`)) return;
  document.getElementById("cardCuentas").classList.add("hidden");
}

function cbBuscar(v) {
  cbBusqueda = String(v || "").trim().toLowerCase();
  cbRender();
}

function cbToggleGrupo(cap) { cbGruposAbiertos[cap] = !cbGruposAbiertos[cap]; cbRender(); }
function cbToggleHijas(code) { cbHijasAbiertas[code] = !cbHijasAbiertas[code]; cbRender(); }
function cbEditar(code) { cbEditando = code; cbRender(); }

function cbTextoDe(e) {
  const cli = cbCliente(e);
  return [cli.code, cli.description, e.code, e.description, e.category]
    .concat(cbFuentes(e).map(f => f.code + " " + f.nom)).join(" ").toLowerCase();
}

function cbRender() {
  const filtradas = cbBusqueda
    ? cbMapping.filter(e => cbTextoDe(e).includes(cbBusqueda))
    : cbMapping;

  const porCap = {};
  for (const cap of CB_CAPITULOS) porCap[cap] = [];
  for (const e of filtradas) (porCap[e.category] = porCap[e.category] || []).push(e);

  let html = "";
  for (const cap of Object.keys(porCap)) {
    const filas = porCap[cap];
    if (!filas.length) continue;
    // Buscando, los grupos se abren solos: si no, hay que abrirlos a mano para ver el
    // resultado de la búsqueda, que es exactamente lo que no se quiere.
    const abierto = cbBusqueda ? true : !!cbGruposAbiertos[cap];
    html += `
      <div class="cb-grupo">
        <button class="cb-grupo-tit" onclick="cbToggleGrupo('${cbEsc(cap)}')">
          <span class="cb-chev">${abierto ? "▾" : "▸"}</span>
          <span>${cbEsc(cap)}</span>
          <span class="cb-cant">${filas.length} fila${filas.length === 1 ? "" : "s"}</span>
        </button>
        ${abierto ? filas.map(cbFilaHtml).join("") : ""}
      </div>`;
  }
  // Pedir el codigo y el nombre con dos prompt() seguidos era incomodo y no dejaba ver
  // lo que se estaba escribiendo. El formulario va arriba, a la vista.
  const formNueva = cbAgregandoFila ? `
    <div class="cb-nueva">
      <input type="text" id="cbNuevaCod" placeholder="Código que ve el cliente" style="max-width:200px">
      <input type="text" id="cbNuevaNom" placeholder="Nombre" style="max-width:300px">
      <button onclick="cbConfirmarFila()">Agregar</button>
      <button class="secundario" onclick="cbAgregandoFila=false; cbRender()">Cancelar</button>
    </div>` : "";
  // El bloque de lo suelto va SIEMPRE arriba, y no lo filtra la búsqueda: si buscando una
  // fila desapareciera, dejaría de servir para lo que está.
  document.getElementById("cbSinAsignar").innerHTML = cbSinAsignarHtml();
  document.getElementById("cbGrupos").innerHTML = formNueva + (html ||
    '<p class="footer-note">No hay filas que coincidan con la búsqueda.</p>');

  const conCliente = cbMapping.filter(e => e.cliente).length;
  const sinCuentas = cbMapping.filter(e => !cbFuentes(e).length).length;
  const onvioSueltas = cbSinAsignar().cuentas.length;
  const madres = cbMapping.filter(e => e.type === "parent").length;
  document.getElementById("cbPie").innerHTML =
    `${cbMapping.length} filas · ${madres} cuentas madre · ${conCliente} con un código distinto ` +
    `para el cliente · ${sinCuentas} sin cuentas asignadas · ${onvioSueltas} cuenta(s) de Onvio ` +
    `sin ubicar` +
    (cbBusqueda ? ` · mostrando ${filtradas.length} que coinciden con "${cbEsc(cbBusqueda)}"` : "") +
    (cbCambios.length
      ? `<br><strong>${cbCambios.length} cambio(s) sin guardar:</strong> ${cbCambios.map(cbEsc).join(" · ")}`
      : "");
  document.getElementById("btnCbGuardar").disabled = cbCambios.length === 0;

  // El desplegable de subcuentas tiene 163 cuentas: sin buscador no se encuentra ninguna.
  // conBuscador() es el mismo helper que ya usa el resto del sitio.
  if (cbAgregandoHija && typeof conBuscador === "function") {
    conBuscador(document.getElementById(`cbSelHija_${cbAgregandoHija}`), "Buscar cuenta de Onvio…");
  }
  // Son 206 filas: el desplegable de destino tampoco sirve sin buscador.
  if (cbAsignando && typeof conBuscador === "function") {
    conBuscador(document.getElementById("cbSelDestino"), "Buscar la fila del cliente…");
  }
}

function cbFilaHtml(e) {
  const cli = cbCliente(e);
  const fuentes = cbFuentes(e);
  const editando = cbEditando === e.code;
  const esMadre = e.type === "parent";
  const abierta = !!cbHijasAbiertas[e.code] || (cbBusqueda && esMadre);
  const sinCuentas = !fuentes.length;

  // Los chips son para la EXCEPCIÓN. Poner uno en cada fila —"1 cuenta" en las 150 filas
  // normales— es ruido que tapa las pocas que piden atención.
  const chips = [];
  if (sinCuentas) chips.push('<span class="cb-chip aviso">sin cuentas</span>');
  // "0 subcuentas" al lado de "sin cuentas" es decir dos veces lo mismo.
  else if (esMadre) chips.push(`<span class="cb-chip dato">${fuentes.length} subcuenta${fuentes.length === 1 ? "" : "s"}</span>`);
  if (e.type === "range") chips.push('<span class="cb-chip dato">rango</span>');
  if (e.ocultar_si_cero) chips.push('<span class="cb-chip">se oculta en cero</span>');

  // La cuenta real solo se muestra cuando es DISTINTA de la que ve el cliente. En 168 de
  // las 206 filas los dos códigos coinciden, y repetirlos era la mitad del ruido.
  let detalle = "";
  if (esMadre) {
    // Las subcuentas se editan acá: cada una con su cruz para sacarla, y un boton para
    // sumar otra. Es lo que mas se toca —los tres errores que aparecieron al configurar
    // esto eran cuentas mal ubicadas— y era justamente lo unico que no se podia cambiar.
    detalle = abierta ? `
      <div class="cb-hijas">
        <div class="cb-onvio-tit">cuentas de Onvio que la llenan</div>
        ${fuentes.map(f => `
          <div class="cb-hija">
            <span>${cbEsc(f.code)} ${cbEsc(f.nom)}</span>
            <button class="cb-x" title="Sacar esta subcuenta"
                    onclick="cbQuitarSubcuenta('${cbEsc(e.code)}','${cbEsc(f.code)}')">×</button>
          </div>`).join("")}
        ${cbAgregandoHija === e.code
          ? `<div class="cb-hija cb-sumar">
               <select id="cbSelHija_${cbEsc(e.code)}" style="max-width:340px;">
                 <option value="">Elegí la cuenta de Onvio…</option>
                 ${cbCuentasDisponibles().map(c =>
                   `<option value="${cbEsc(c.code)}">${cbEsc(c.code)} ${cbEsc(c.nom)}` +
                   `${c.dueño ? " — ya está en " + cbEsc(c.dueño) : ""}</option>`).join("")}
               </select>
               <button onclick="cbConfirmarSubcuenta('${cbEsc(e.code)}')">Agregar</button>
               <button class="secundario" onclick="cbAgregandoHija=null; cbRender()">Cancelar</button>
             </div>`
          : `<button class="cb-sumar-btn" onclick="cbAgregandoHija='${cbEsc(e.code)}'; cbRender()">
               + Agregar subcuenta</button>`}
      </div>` : "";
  } else if (e.type === "range") {
    detalle = `<div class="cb-real">${cbEsc(fuentes[0].nom)}</div>`;
  } else if (e.cliente) {
    detalle = `<div class="cb-real">en Onvio: ${cbEsc(e.code)} ${cbEsc(e.description)}</div>`;
  }

  const acciones = editando ? "" : `
    <div class="cb-acc">
      <button onclick="cbEditar('${cbEsc(e.code)}')" title="Cambiar el código o el nombre que ve el cliente">Editar</button>
      ${(!esMadre && e.type !== "range")
        ? `<button onclick="cbConvertirEnMadre('${cbEsc(e.code)}')"
                   title="Convertirla en cuenta madre para poder colgarle subcuentas">Hacer madre</button>`
        : ""}
      <button onclick="cbToggleOcultar('${cbEsc(e.code)}')" title="Mostrarla siempre o solo cuando tenga movimiento">
        ${e.ocultar_si_cero ? "Mostrar" : "Ocultar en cero"}</button>
      <button class="cb-peligro" onclick="cbQuitar('${cbEsc(e.code)}')">Quitar</button>
    </div>`;

  return `
    <div class="cb-fila${sinCuentas ? " cb-atencion" : ""}">
      ${esMadre
        ? `<button class="cb-tog" onclick="cbToggleHijas('${cbEsc(e.code)}')"
                   title="Ver las cuentas que la alimentan">${abierta ? "▾" : "▸"}</button>`
        : '<span class="cb-nada"></span>'}
      <div class="cb-txt">
        <div><span class="cb-cod">${cbEsc(cli.code)}</span> <span class="cb-nom">${cbEsc(cli.description)}</span></div>
        ${detalle}
        ${editando ? `
          <div class="cb-edit">
            <input type="text" id="cbCod_${cbEsc(e.code)}" value="${cbEsc(cli.code)}" style="max-width:130px" placeholder="código">
            <input type="text" id="cbNom_${cbEsc(e.code)}" value="${cbEsc(cli.description)}" style="max-width:260px" placeholder="nombre">
            <button onclick="cbGuardarFila('${cbEsc(e.code)}')">Aceptar</button>
            <button class="secundario" onclick="cbEditando=null; cbRender()">Cancelar</button>
          </div>` : ""}
      </div>
      <div class="cb-chips">${chips.join("")}</div>
      ${acciones}
    </div>`;
}

function cbGuardarFila(code) {
  const e = cbMapping.find(x => x.code === code);
  if (!e) return;
  const cod = document.getElementById(`cbCod_${code}`).value.trim();
  const nom = document.getElementById(`cbNom_${code}`).value.trim();
  if (!cod || !nom) { alert("El código y el nombre no pueden quedar vacíos."); return; }

  // Dos filas con el mismo código de cliente serían dos renglones iguales en el informe.
  const choca = cbMapping.find(x => x.code !== code && cbCliente(x).code === cod);
  if (choca) {
    alert(`Ese código ya lo usa otra fila: "${cbCliente(choca).code} ${cbCliente(choca).description}". ` +
          `Dos filas con el mismo código salen repetidas en el informe.`);
    return;
  }

  const antes = cbCliente(e);
  if (antes.code === cod && antes.description === nom) { cbEditando = null; cbRender(); return; }
  if (cod === e.code && nom === e.description) delete e.cliente;   // vuelven a coincidir
  else e.cliente = { code: cod, description: nom };
  cbCambios.push(`"${antes.code} ${antes.description}" pasa a "${cod} ${nom}"`);
  cbEditando = null;
  cbRender();
}

function cbToggleOcultar(code) {
  const e = cbMapping.find(x => x.code === code);
  if (!e) return;
  const cli = cbCliente(e);
  if (e.ocultar_si_cero) {
    delete e.ocultar_si_cero;
    cbCambios.push(`"${cli.code} ${cli.description}" se muestra siempre`);
  } else {
    e.ocultar_si_cero = true;
    cbCambios.push(`"${cli.code} ${cli.description}" se oculta si está en cero`);
  }
  cbRender();
}

function cbQuitar(code) {
  const e = cbMapping.find(x => x.code === code);
  if (!e) return;
  const cli = cbCliente(e);
  const fuentes = cbFuentes(e);

  // Sacar una fila que se llena con cuentas de Onvio es peligroso: esa plata deja de tener
  // dónde caer. Y si la fila tiene saldo guardado del mes pasado, ese saldo se pierde —
  // pasó de verdad al reorganizar las cuentas en dólares.
  const saldo = (typeof estadoB === "object" && estadoB && estadoB.saldos)
    ? estadoB.saldos[e.code] : undefined;
  let aviso = `¿Sacar del informe la fila "${cli.code} ${cli.description}"?`;
  if (fuentes.length) {
    aviso += `\n\nOJO: esta fila se llena con ${fuentes.length === 1 ? "una cuenta" : fuentes.length + " cuentas"} ` +
             `de Onvio. Si esas cuentas traen importe y la fila no está, ese importe queda sin destino.`;
  }
  if (typeof saldo === "number" && Math.abs(saldo) > 0.005) {
    aviso += `\n\nAdemás tiene un saldo guardado del mes pasado de ${saldo.toFixed(2)}, que se va a perder.`;
  }
  if (!confirm(aviso)) return;

  cbMapping = cbMapping.filter(x => x.code !== code);
  cbCambios.push(`se saca "${cli.code} ${cli.description}"`);
  cbRender();
}

function cbAgregar() {
  cbAgregandoFila = true;
  cbRender();
  const c = document.getElementById("cbNuevaCod");
  if (c) c.focus();
}

function cbConfirmarFila() {
  const codigo = (document.getElementById("cbNuevaCod").value || "").trim();
  const nom = (document.getElementById("cbNuevaNom").value || "").trim();
  if (!/^\d{5,}$/.test(codigo)) { alert("El código tiene que ser un número de 5 dígitos o más."); return; }
  if (!nom) { alert("Poné un nombre."); return; }
  if (cbMapping.some(x => cbCliente(x).code === codigo || x.code === codigo)) {
    alert("Ya hay una fila con ese código."); return;
  }
  const PORDIGITO = { "1": "ACTIVO", "2": "PASIVO", "3": "CAPITAL Y PATRIMONIO", "4": "RESULTADOS" };
  const categoria = PORDIGITO[codigo[0]] || "RESULTADOS";
  // Arranca SIN cuentas asignadas: es una fila que el cliente espera ver y que sale en
  // cero. Para llenarla con cuentas de Onvio se la convierte en madre y se le agregan.
  const orden = Math.max(...cbMapping.filter(x => x.category === categoria).map(x => x.orden || 0), 0) + 0.5;
  cbMapping.push({
    code: codigo, aliases: [], description: nom,
    category: categoria, type: "simple", orden, sin_cuentas: true,
  });
  cbCambios.push(`se agrega "${codigo} ${nom}" en ${categoria}`);
  cbGruposAbiertos[categoria] = true;
  cbAgregandoFila = false;
  cbRender();
}

function cbQuitarSubcuenta(padre, hija) {
  const e = cbMapping.find(x => x.code === padre);
  if (!e) return;
  const h = (e.children || []).find(x => x.code === hija);
  if (!h) return;
  // Sacarla de acá la deja sin ninguna fila donde caer: si esa cuenta trae importe, ese
  // importe queda afuera del informe. La app avisa al procesar, pero mejor decirlo ahora.
  if (!confirm(`¿Sacar "${hija} ${h.description}" de "${cbCliente(e).code} ${cbCliente(e).description}"?` +
               `

Si no la ponés en otra fila, su importe queda sin destino y la app te lo va a avisar ` +
               `al procesar el mes.`)) return;
  e.children = (e.children || []).filter(x => x.code !== hija);
  cbCambios.push(`"${hija}" sale de "${cbCliente(e).code}"`);
  cbRender();
}

function cbConfirmarSubcuenta(padre) {
  const e = cbMapping.find(x => x.code === padre);
  if (!e) return;
  const sel = document.getElementById(`cbSelHija_${padre}`);
  const cod = sel ? sel.value : "";
  if (!cod) { alert("Elegí una cuenta."); return; }
  if ((e.children || []).some(h => h.code === cod)) { alert("Esa cuenta ya está en esta madre."); return; }

  // Una cuenta en dos filas se suma dos veces. Si ya tiene dueño, se la muda.
  const anterior = cbMapping.find(x => x !== e &&
    ((x.children || []).some(h => h.code === cod) || (x.type !== "parent" && x.type !== "range" && x.code === cod)));
  const disponible = cbCuentasDisponibles().find(c => c.code === cod);
  const nombre = disponible ? disponible.nom : cod;
  if (anterior) {
    if (!confirm(`"${cod} ${nombre}" hoy está en "${cbCliente(anterior).code} ${cbCliente(anterior).description}".` +
                 `

¿Mudarla acá? Si quedara en las dos, su importe se sumaría dos veces.`)) return;
    if (anterior.type === "parent") anterior.children = anterior.children.filter(h => h.code !== cod);
    else cbMapping = cbMapping.filter(x => x !== anterior);   // era una fila propia
    cbCambios.push(`"${cod}" se muda de "${cbCliente(anterior).code}" a "${cbCliente(e).code}"`);
  } else {
    cbCambios.push(`"${cod}" entra en "${cbCliente(e).code}"`);
  }
  e.children = (e.children || []).concat([{ code: cod, description: nombre }]);
  cbAgregandoHija = null;
  cbRender();
}

// Convertir una fila simple en cuenta madre es exactamente lo que pidio el cliente para
// Seguridad y Alquileres: el codigo viejo pasa a ser la madre y la cuenta real de Onvio,
// su primera subcuenta.
function cbConvertirEnMadre(code) {
  const e = cbMapping.find(x => x.code === code);
  if (!e || e.type === "parent") return;
  const cli = cbCliente(e);
  if (!confirm(`"${cli.code} ${cli.description}" pasa a ser cuenta madre.` +
               (e.sin_cuentas ? `

Arranca sin subcuentas; se las agregás vos.` :
                `

Su cuenta de Onvio "${e.code} ${e.description}" pasa a ser la primera subcuenta.`))) return;
  const hijas = e.sin_cuentas ? [] : [{ code: e.code, description: e.description }];
  const i = cbMapping.indexOf(e);
  cbMapping[i] = {
    code: cli.code, aliases: [], description: cli.description,
    category: e.category, type: "parent", orden: e.orden, children: hijas,
  };
  if (e.ocultar_si_cero) cbMapping[i].ocultar_si_cero = true;
  cbCambios.push(`"${cli.code} ${cli.description}" pasa a ser cuenta madre`);
  cbHijasAbiertas[cli.code] = true;
  cbRender();
}

async function cbGuardar() {
  const st = document.getElementById("cbStatus");
  // El mismo control que usa el resto de la app, no una copia. Acá se pedía además un campo
  // `owner`, que la configuración NUNCA tuvo —se guarda {token, repo, branch, path}, con el
  // repo entero en `repo` ("usuaria/Informes")— así que la condición fallaba siempre y el
  // botón Guardar del panel no pudo guardar nunca, con GitHub configurado o sin configurar.
  if (!hasGhSettings()) {
    st.innerHTML = '<div class="status-msg bad">Configurá GitHub (⚙, arriba a la derecha) antes de guardar.</div>';
    return;
  }
  document.getElementById("btnCbGuardar").disabled = true;
  st.innerHTML = '<div class="status-msg">Guardando…</div>';
  try {
    // Se relee el sha justo antes de escribir: entre que se abrió el panel y ahora, la
    // corrida mensual pudo haber guardado cuentas nuevas.
    const { sha } = await ghGetFile();
    await ghPutFile(cbMapping, sha, `Informe B: ${cbCambios.length} cambio(s) en la configuración de cuentas`);
    currentMapping = cbMapping;
    st.innerHTML = `<div class="status-msg ok">Guardado. ${cbCambios.length} cambio(s) subidos a GitHub.</div>`;
    cbCambios = [];
    cbRender();
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">No se pudo guardar: ${cbEsc(e.message)}</div>`;
    document.getElementById("btnCbGuardar").disabled = false;
  }
}
