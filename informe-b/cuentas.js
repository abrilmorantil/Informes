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
  document.getElementById("cbGrupos").innerHTML = html ||
    '<p class="footer-note">No hay filas que coincidan con la búsqueda.</p>';

  const conCliente = cbMapping.filter(e => e.cliente).length;
  const sinCuentas = cbMapping.filter(e => !cbFuentes(e).length).length;
  const madres = cbMapping.filter(e => e.type === "parent").length;
  document.getElementById("cbPie").innerHTML =
    `${cbMapping.length} filas · ${madres} cuentas madre · ${conCliente} con un código distinto ` +
    `para el cliente · ${sinCuentas} sin cuentas asignadas` +
    (cbBusqueda ? ` · mostrando ${filtradas.length} que coinciden con "${cbEsc(cbBusqueda)}"` : "") +
    (cbCambios.length
      ? `<br><strong>${cbCambios.length} cambio(s) sin guardar:</strong> ${cbCambios.map(cbEsc).join(" · ")}`
      : "");
  document.getElementById("btnCbGuardar").disabled = cbCambios.length === 0;
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
  if (esMadre) chips.push(`<span class="cb-chip dato">${fuentes.length} subcuenta${fuentes.length === 1 ? "" : "s"}</span>`);
  if (e.type === "range") chips.push('<span class="cb-chip dato">rango</span>');
  if (e.ocultar_si_cero) chips.push('<span class="cb-chip">se oculta en cero</span>');

  // La cuenta real solo se muestra cuando es DISTINTA de la que ve el cliente. En 168 de
  // las 206 filas los dos códigos coinciden, y repetirlos era la mitad del ruido.
  let detalle = "";
  if (esMadre) {
    detalle = abierta
      ? `<div class="cb-hijas">${fuentes.map(f =>
          `<div class="cb-hija">${cbEsc(f.code)} ${cbEsc(f.nom)}</div>`).join("")}</div>`
      : "";
  } else if (e.type === "range") {
    detalle = `<div class="cb-real">${cbEsc(fuentes[0].nom)}</div>`;
  } else if (e.cliente) {
    detalle = `<div class="cb-real">en Onvio: ${cbEsc(e.code)} ${cbEsc(e.description)}</div>`;
  }

  const acciones = editando ? "" : `
    <div class="cb-acc">
      <button onclick="cbEditar('${cbEsc(e.code)}')" title="Cambiar el código o el nombre que ve el cliente">Editar</button>
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
  const cod = prompt("Código que va a ver el cliente (ej: 42433000)");
  if (!cod) return;
  const codigo = cod.trim();
  if (!/^\d{5,}$/.test(codigo)) { alert("El código tiene que ser un número de 5 dígitos o más."); return; }
  if (cbMapping.some(x => cbCliente(x).code === codigo || x.code === codigo)) {
    alert("Ya hay una fila con ese código."); return;
  }
  const nom = prompt("Nombre que va a ver el cliente");
  if (!nom || !nom.trim()) return;

  const PORDIGITO = { "1": "ACTIVO", "2": "PASIVO", "3": "CAPITAL Y PATRIMONIO", "4": "RESULTADOS" };
  const categoria = PORDIGITO[codigo[0]] || "RESULTADOS";
  // Se agrega SIN cuentas asignadas: es una fila que el cliente espera ver y que siempre
  // sale en cero. Si más adelante tiene que llenarse con una cuenta de Onvio, se cambia acá.
  const orden = Math.max(...cbMapping.filter(x => x.category === categoria).map(x => x.orden || 0), 0) + 0.5;
  cbMapping.push({
    code: codigo, aliases: [], description: nom.trim(),
    category: categoria, type: "simple", orden, sin_cuentas: true,
  });
  cbCambios.push(`se agrega "${codigo} ${nom.trim()}" en ${categoria}`);
  cbGruposAbiertos[categoria] = true;
  cbRender();
}

async function cbGuardar() {
  const st = document.getElementById("cbStatus");
  const s = loadGhSettings();
  if (!s.token || !s.owner || !s.repo) {
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
