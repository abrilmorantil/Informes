// Panel "Configurar cuentas" — la vista maestra de qué fila ve el cliente y con qué
// cuentas reales se llena.
//
// Por qué existe: este informe habla DOS numeraciones a la vez. El cliente lee el plan de
// cuentas viejo y nosotras trabajamos con el que manda Onvio, y no hay ninguna regla que
// relacione uno con el otro — los códigos fueron reasignados, no reformateados (medido:
// la regla de "sacarle un dígito al medio" acierta 1 de 38). Así que la equivalencia se
// declara a mano, y hasta ahora vivía escrita en la columna A del Excel, a mano también.
// Acá se ve y se cambia.
//
// Los cambios se acumulan en memoria y recién se suben al apretar "Guardar cambios".

let cbMapping = null;       // copia de trabajo del mapping
let cbCambios = [];         // qué se hizo, para el mensaje del commit
let cbBusqueda = "";
let cbAbierto = false;      // el detalle arranca cerrado: son 206 filas
let cbEditando = null;      // código de la fila que se está editando

function cbEsc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// El código y el nombre que ve el cliente. Si la fila no declara `cliente`, es porque las
// dos numeraciones coinciden y se muestra la propia.
function cbCliente(e) {
  return (e && e.cliente) ? e.cliente : { code: e.code, description: e.description };
}

// Con qué cuentas reales de Onvio se llena la fila. Una cuenta madre se llena con sus
// subcuentas; el rango de proveedores, con todo lo que empiece con su prefijo; una fila
// simple, con su propio código.
function cbFuentes(e) {
  if (e.type === "parent") return (e.children || []).map(h => `${h.code} ${h.description}`);
  if (e.type === "range") return [`todas las cuentas que empiezan con ${e.prefix}`];
  if (e.sin_cuentas) return [];
  return [`${e.code} ${e.description}`];
}

function cbEstado(e) {
  const f = cbFuentes(e);
  if (!f.length) return { texto: "Sin cuentas asignadas", clase: "bad" };
  if (e.type === "parent") return { texto: `${f.length} subcuenta(s)`, clase: "ok" };
  if (e.type === "range") return { texto: "Rango de proveedores", clase: "ok" };
  return { texto: "1 cuenta", clase: "ok" };
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
  cbAbierto = false;
  document.getElementById("cardCuentas").classList.remove("hidden");
  cbRender();
  document.getElementById("cardCuentas").scrollIntoView({ behavior: "smooth", block: "start" });
}

function cbCerrar() {
  if (cbCambios.length && !confirm(`Hay ${cbCambios.length} cambio(s) sin guardar. ¿Cerrar igual y perderlos?`)) return;
  document.getElementById("cardCuentas").classList.add("hidden");
}

function cbToggle() {
  cbAbierto = !cbAbierto;
  cbRender();
}

function cbBuscar(v) {
  cbBusqueda = String(v || "").trim().toLowerCase();
  if (cbBusqueda) cbAbierto = true;
  cbRender();
}

function cbFilas() {
  return cbMapping.map(e => {
    const cli = cbCliente(e);
    const est = cbEstado(e);
    return {
      e, cli, est,
      texto: [cli.code, cli.description, e.code, e.description, e.category, est.texto]
        .concat(cbFuentes(e)).join(" ").toLowerCase(),
    };
  });
}

function cbRender() {
  const filas = cbFilas();
  const visibles = cbBusqueda ? filas.filter(f => f.texto.includes(cbBusqueda)) : filas;

  document.getElementById("cbCuerpo").innerHTML = visibles.length ? visibles.map(f => {
    const e = f.e;
    const fuentes = cbFuentes(e);
    const editando = cbEditando === e.code;
    return `
      <tr>
        <td>
          <div><strong>${cbEsc(f.cli.code)}</strong></div>
          <div class="footer-note">${cbEsc(f.cli.description)}</div>
          ${e.ocultar_si_cero ? '<div class="footer-note">se oculta si está en cero</div>' : ""}
        </td>
        <td>
          ${fuentes.length
            ? `<div class="footer-note">${fuentes.map(cbEsc).join("<br>")}</div>`
            : '<span class="badge bad">ninguna</span>'}
        </td>
        <td><span class="badge ${f.est.clase}">${cbEsc(f.est.texto)}</span></td>
        <td>
          ${editando ? `
            <input type="text" id="cbCod_${cbEsc(e.code)}" value="${cbEsc(f.cli.code)}"
                   style="max-width:120px" placeholder="código">
            <input type="text" id="cbNom_${cbEsc(e.code)}" value="${cbEsc(f.cli.description)}"
                   style="max-width:220px" placeholder="nombre">
            <button onclick="cbGuardarFila('${cbEsc(e.code)}')">Aceptar</button>
            <button class="secundario" onclick="cbEditando=null; cbRender()">Cancelar</button>
          ` : `
            <button class="secundario" onclick="cbEditar('${cbEsc(e.code)}')">Lo que ve el cliente</button>
            <button class="secundario" onclick="cbToggleOcultar('${cbEsc(e.code)}')">
              ${e.ocultar_si_cero ? "Mostrar siempre" : "Ocultar si está en cero"}
            </button>
            <button class="secundario" onclick="cbQuitar('${cbEsc(e.code)}')">Quitar</button>
          `}
        </td>
      </tr>`;
  }).join("") : '<tr><td colspan="4" class="footer-note">No hay filas que coincidan con la búsqueda.</td></tr>';

  document.getElementById("cbLista").classList.toggle("hidden", !cbAbierto);
  document.getElementById("cbFlecha").textContent = cbAbierto ? "▾" : "▸";
  document.getElementById("cbVerTexto").textContent = cbAbierto
    ? `Ocultar el detalle (${visibles.length} de ${filas.length} filas)`
    : (visibles.length === filas.length
        ? `Ver el detalle fila por fila (${filas.length})`
        : `Ver el detalle (${visibles.length} de ${filas.length} filas)`);

  const sinCuentas = filas.filter(f => !cbFuentes(f.e).length).length;
  const conCliente = filas.filter(f => f.e.cliente).length;
  document.getElementById("cbPie").innerHTML =
    `${filas.length} filas. ${conCliente} tienen un código distinto para el cliente, ` +
    `${sinCuentas} no tienen ninguna cuenta asignada.` +
    (cbCambios.length
      ? `<br><strong>${cbCambios.length} cambio(s) sin guardar:</strong> ${cbCambios.map(cbEsc).join(" · ")}`
      : "");
  document.getElementById("btnCbGuardar").disabled = cbCambios.length === 0;
}

function cbEditar(code) { cbEditando = code; cbRender(); }

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
  // Si vuelve a coincidir con la cuenta real, no hace falta declarar nada.
  if (cod === e.code && nom === e.description) delete e.cliente;
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
  cbAbierto = true;
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
