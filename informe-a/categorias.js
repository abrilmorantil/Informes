// Panel "Configurar categorización" — la vista maestra de qué cuenta va a qué categoría.
//
// Por qué existe: hasta ahora la categorización se definía SOLO por la pregunta que
// aparece cuando el export trae una cuenta nueva. Una vez contestada no había forma de
// verla ni de cambiarla. Este panel es la fuente central: se ven todas las cuentas
// configuradas, las que quedaron sin categoría, y las que están repartidas en varias.
//
// Lo importante de cómo funciona: cambiar la categoría acá REESCRIBE LAS FÓRMULAS de
// Dist.de gastos, no sólo la etiqueta del mapeo. La etiqueta sola no mueve ningún
// importe (sólo decide dónde se engancha una referencia nueva), así que un panel que
// tocara nada más el mapeo daría la sensación de haber arreglado algo que sigue igual.
//
// Los cambios se acumulan en memoria sobre una copia del archivo y del mapeo, y recién
// se suben a GitHub al apretar "Guardar cambios". Hasta entonces no se toca nada.

let catWb = null;          // copia del archivo base sobre la que se edita
let catMapeo = null;       // copia del mapeo
let catCambios = [];       // qué se hizo, para el mensaje del commit y para mostrarlo
let catBusqueda = "";
let catEditando = null;    // código de la cuenta cuya categoría se está eligiendo
let catListaAbierta = false;  // el detalle cuenta por cuenta arranca cerrado

// Sólo se listan las cuentas con la numeración nueva (9 dígitos). Las de 8 son del plan
// viejo, están todas en cero y se van a sacar por completo del informe más adelante.
const CAT_ES_VIVA = (codigo) => String(codigo).length === 9;

function catEstadoDe(cuenta, filasSaldo, filasHaber) {
  const n = filasSaldo ? filasSaldo.size : 0;
  const todas = new Set([...(filasSaldo || []), ...(filasHaber || [])]);
  if (cuenta.excluida) return { clave: "excluida", texto: "No se distribuye", clase: "bad" };
  if (n > 1) return { clave: "varias", texto: `En ${n} categorías`, clase: "bad" };
  // El saldo está en una sola categoría, pero el haber (columna CR) se suma en otra.
  if (todas.size > 1) return { clave: "haber", texto: "El haber va a otra", clase: "bad" };
  if (!cuenta.categoria) return { clave: "sin", texto: "Sin categoría", clase: "bad" };
  if (todas.size === 0) return { clave: "ok", texto: "Configurada, sin usar", clase: "ok" };
  return { clave: "ok", texto: "Configurada", clase: "ok" };
}

// La categoría que el panel muestra es la que dicen las FÓRMULAS, no la etiqueta del
// mapeo: son las fórmulas las que deciden dónde cae el importe. Cuando hay más de una
// se listan todas, que es justamente lo que hay que resolver.
function catCategoriasDe(filasDist) {
  if (!filasDist || !filasDist.size) return [];
  const porFila = {};
  for (const c of catMapeo.categorias) porFila[c.dist_row] = c.desc;
  return [...filasDist].sort((a, b) => a - b).map(f => ({ fila: f, desc: porFila[f] || `(fila ${f})` }));
}

function catFilas() {
  const wsDist = catWb.getWorksheet("Dist.de gastos");
  const { saldo, haber } = mapaDeDistribucion(wsDist, catMapeo);
  const filas = [];
  for (const [codigo, cuenta] of Object.entries(catMapeo.cuentas)) {
    if (!CAT_ES_VIVA(codigo)) continue;
    const fs = saldo.get(cuenta.ss_row);
    const fh = haber.get(cuenta.ss_row);
    filas.push({
      codigo, cuenta, ssRow: cuenta.ss_row,
      cats: catCategoriasDe(fs),
      catsHaber: catCategoriasDe(fh),
      estado: catEstadoDe(cuenta, fs, fh),
    });
  }
  // primero lo que hay que resolver, después el resto
  const orden = { varias: 0, haber: 1, sin: 2, excluida: 3, ok: 4 };
  filas.sort((a, b) => (orden[a.estado.clave] - orden[b.estado.clave]) ||
                       (b.cats.length - a.cats.length) || (a.ssRow - b.ssRow));
  return filas;
}

async function abrirCategorias() {
  if (typeof wbBorrador !== "undefined" && wbBorrador) {
    alert("Hay una corrida del mes sin cerrar. Cerrá o descartá ese borrador antes de " +
          "tocar la categorización: el borrador se armó con el archivo anterior y los " +
          "cambios de acá no estarían en él.");
    return;
  }
  if (!bufferBase || !mapeoGuardado) {
    alert("Todavía no se cargó el archivo base.");
    return;
  }
  mostrar("cardCategorias", true);
  document.getElementById("catBody").innerHTML =
    '<tr><td colspan="4" class="footer-note">Leyendo el archivo del balance… ' +
    'tarda unos segundos, es el mismo archivo que abre la carga mensual.</td></tr>';
  document.getElementById("catPie").innerHTML = "";
  // el navegador tiene que llegar a pintar el aviso antes de bloquearse abriendo el libro
  await new Promise(r => setTimeout(r, 30));
  // copia propia: hasta que se guarde, el archivo del ciclo mensual no se toca
  catWb = await abrirWorkbook(bufferBase.slice(0));
  catMapeo = JSON.parse(JSON.stringify(mapeoGuardado));
  catCambios = [];
  catBusqueda = "";
  catEditando = null;
  catListaAbierta = false;
  document.getElementById("catBuscar").value = "";
  renderCategorias();
  document.getElementById("cardCategorias").scrollIntoView({ behavior: "smooth", block: "start" });
}

function cerrarCategorias() {
  if (catCambios.length &&
      !confirm(`Hay ${catCambios.length} cambio(s) sin guardar. Si cerrás se pierden. ¿Cerrar igual?`)) {
    return;
  }
  catWb = null;
  catMapeo = null;
  catCambios = [];
  mostrar("cardCategorias", false);
}

function renderCategorias() {
  const filas = catFilas();

  // El panel no muestra tablero de contadores ni filtros: con todo ya configurado no
  // aportaban y ocupaban la pantalla entera. Queda el buscador, que ademas de codigo y
  // nombre busca por el estado ("en 2 categorias", "sin categoria"...), asi que lo que
  // haya para revisar se encuentra igual.
  const q = catBusqueda.trim().toLowerCase();
  const visibles = !q ? filas : filas.filter(f =>
    f.codigo.includes(q) ||
    String(f.cuenta.label || "").toLowerCase().includes(q) ||
    f.estado.texto.toLowerCase().includes(q) ||
    f.cats.some(c => c.desc.toLowerCase().includes(q)));

  const opcionesCat = catOpcionesCategoria();
  document.getElementById("catBody").innerHTML = visibles.length ? visibles.map(f => {
    const editando = catEditando === f.codigo;
    const celdaCat = editando
      ? `<select class="catCambiar" data-cod="${f.codigo}">
           <option value="">— elegí la categoría —</option>
           ${opcionesCat}
         </select>`
      : (f.cats.length
          ? f.cats.map(c => `<div>${catEsc(c.desc)} <span class="footer-note">(fila ${c.fila})</span></div>`).join("")
          : `<span class="footer-note">${f.cuenta.categoria ? catEsc(f.cuenta.categoria) + " (todavía sin fórmulas)" : "—"}</span>`)
        + (f.estado.clave === "haber"
            ? `<div class="footer-note">haber en: ${f.catsHaber.map(c => catEsc(c.desc)).join(", ")}</div>`
            : "");
    return `
      <tr>
        <td><strong>${f.codigo}</strong><br><span class="footer-note">${catEsc(f.cuenta.label || "")} · fila ${f.ssRow}</span></td>
        <td>${celdaCat}</td>
        <td><span class="badge ${f.estado.clase}">${f.estado.texto}</span></td>
        <td style="white-space:nowrap;">
          ${editando
            ? `<button class="secundario" onclick="catAplicarCategoria('${f.codigo}')">Aplicar</button>
               <button class="secundario" onclick="catCancelarEdicion()">Cancelar</button>`
            : `<button class="secundario" onclick="catEditarCategoria('${f.codigo}')">Cambiar categoría</button>
               <button class="secundario" onclick="catRenombrar('${f.codigo}')">Renombrar</button>
               ${f.cuenta.excluida
                  ? ""
                  : `<button class="secundario" onclick="catQuitar('${f.codigo}')">Quitar</button>`}`}
        </td>
      </tr>`;
  }).join("") : '<tr><td colspan="4" class="footer-note">No hay cuentas que cumplan el filtro.</td></tr>';

  if (catEditando) conBuscadorTodos(".catCambiar", "Buscar categoría…");

  mostrar("catLista", catListaAbierta);
  document.getElementById("catFlecha").textContent = catListaAbierta ? "▾" : "▸";
  document.getElementById("catVerTexto").textContent = catListaAbierta
    ? `Ocultar el detalle (${visibles.length} de ${filas.length} cuentas)`
    : (visibles.length === filas.length
        ? `Ver el detalle cuenta por cuenta (${filas.length})`
        : `Ver el detalle (${visibles.length} de ${filas.length} cuentas)`);

  const ocultas = Object.keys(catMapeo.cuentas).filter(c => !CAT_ES_VIVA(c)).length;
  document.getElementById("catPie").innerHTML =
    `Mostrando ${visibles.length} de ${filas.length} cuentas. ` +
    `No se listan las ${ocultas} del plan viejo (códigos de 8 dígitos), que están todas en cero.` +
    (catCambios.length
      ? `<br><strong>${catCambios.length} cambio(s) sin guardar:</strong> ${catCambios.map(catEsc).join(" · ")}`
      : "");
  document.getElementById("btnCatGuardar").disabled = catCambios.length === 0;
}

function catEsc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Las opciones llevan la FILA como valor, no el nombre: hay 95 filas de categoría con
// sólo 91 nombres distintos (ALQUILERES VARIOS, ACUERDO DE INVERSIONES, ROBO y GASTOS DE
// ESPECTROMETRÍA están cada uno en dos filas). Eligiendo por nombre siempre se iba a la
// primera, sin aviso. A los repetidos se les muestra la fila para poder distinguirlos.
function catOpcionesCategoria() {
  return categoriasElegibles(catMapeo)
    .map(c => `<option value="${c.fila}">${catEsc(c.texto)}</option>`)
    .join("");
}

// El detalle son 282 renglones: desplegado siempre deja el panel imposible de leer.
// Arranca cerrado y se abre a pedido — o solo, en cuanto se busca o se filtra, que es
// cuando la lista pasa a ser corta y tiene sentido verla.
function catToggleLista() {
  catListaAbierta = !catListaAbierta;
  renderCategorias();
  if (catListaAbierta) {
    document.getElementById("catLista").scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

function catEditarCategoria(codigo) { catEditando = codigo; renderCategorias(); }
function catCancelarEdicion() { catEditando = null; renderCategorias(); }

function catAplicarCategoria(codigo) {
  const sel = document.querySelector(`.catCambiar[data-cod="${codigo}"]`);
  if (!sel || !sel.value) { alert("Elegí una categoría."); return; }
  try {
    const r = moverCuentaDeCategoria({ wb: catWb, mapeo: catMapeo, codigo, categoriaDestino: sel.value });
    catCambios.push(`${codigo} → ${r.categoria}`);
    catEditando = null;
    renderCategorias();
  } catch (e) {
    alert(e.message);
  }
}

function catRenombrar(codigo) {
  const actual = catMapeo.cuentas[codigo].label || "";
  const nuevo = prompt(`Nombre de la cuenta ${codigo}:`, actual);
  if (nuevo === null) return;
  try {
    renombrarCuenta({ wb: catWb, mapeo: catMapeo, codigo, nombre: nuevo });
    catCambios.push(`${codigo} renombrada`);
    renderCategorias();
  } catch (e) {
    alert(e.message);
  }
}

function catQuitar(codigo) {
  const c = catMapeo.cuentas[codigo];
  if (!confirm(
    `¿Sacar ${codigo} ${c.label} de la distribución?\n\n` +
    `La fila ${c.ss_row} sigue en Sumas y Saldos con su saldo, pero deja de sumar en ` +
    `ninguna categoría de Dist.de gastos, y no se va a volver a preguntar por ella.`)) return;
  try {
    const r = quitarCuentaDeDistribucion({ wb: catWb, mapeo: catMapeo, codigo });
    catCambios.push(`${codigo} sacada de la distribución (${r.quitadas} referencias)`);
    renderCategorias();
  } catch (e) {
    alert(e.message);
  }
}

// ------------------------------------------------------------ agregar a mano
function catAbrirAlta() {
  document.getElementById("catAltaCat").innerHTML =
    '<option value="">— elegí la categoría —</option>' + catOpcionesCategoria();
  document.getElementById("catAltaCodigo").value = "";
  document.getElementById("catAltaNombre").value = "";
  document.getElementById("catAltaStatus").innerHTML = "";
  mostrar("catAlta", true);
  conBuscadorTodos("#catAltaCat", "Buscar categoría…");
}

function catCerrarAlta() { mostrar("catAlta", false); }

function catAgregarCuenta() {
  const codigo = document.getElementById("catAltaCodigo").value.trim();
  const nombre = document.getElementById("catAltaNombre").value.trim();
  const cat = document.getElementById("catAltaCat").value;
  const st = document.getElementById("catAltaStatus");
  const mal = (m) => { st.innerHTML = `<div class="status-msg bad">${catEsc(m)}</div>`; };

  if (!/^\d{9}$/.test(codigo)) return mal("El código tiene que ser de 9 dígitos, como los que manda Onvio.");
  if (!codigo.startsWith("4")) return mal("Al balance sólo entran cuentas de resultado, las que empiezan con 4.");
  if (!nombre) return mal("Falta el nombre de la cuenta.");
  if (!cat) return mal("Falta elegir la categoría.");
  if (catMapeo.cuentas[codigo]) return mal(`La cuenta ${codigo} ya está, en la fila ${catMapeo.cuentas[codigo].ss_row}.`);

  try {
    const c = insertarCuentaEnBalance({ wb: catWb, mapeo: catMapeo, codigo, label: nombre, categoria: cat });
    catCambios.push(`${codigo} agregada en ${cat}`);
    mostrar("catAlta", false);
    renderCategorias();
    alert(`Se creó la fila ${c.ss_row} en Sumas y Saldos, con las fórmulas de su vecina.\n\n` +
          `Todavía no suma en ninguna columna de Dist.de gastos: sus referencias se agregan ` +
          `solas la primera vez que la cuenta traiga movimiento en un centro de costo.`);
  } catch (e) {
    mal(e.message);
  }
}

// ------------------------------------------------------------------ guardar
async function guardarCategorias() {
  if (!catCambios.length) return;
  const btn = document.getElementById("btnCatGuardar");
  const st = document.getElementById("catStatus");
  if (!confirm(`Se van a guardar ${catCambios.length} cambio(s) en el repositorio.\n\n` +
               `Se sube el archivo del balance con las fórmulas ya corregidas y el mapeo. ` +
               `Los meses ya cerrados no se tocan: sus importes están congelados.`)) return;
  btn.disabled = true;
  mostrar("spinnerCat", true);
  st.innerHTML = "";
  try {
    // Excel tiene que recalcular al abrir: las fórmulas cambiadas quedan sin resultado.
    catWb.calcProperties = catWb.calcProperties || {};
    catWb.calcProperties.fullCalcOnLoad = true;

    const buffer = await catWb.xlsx.writeBuffer();
    const mensaje = `Categorización: ${catCambios.slice(0, 3).join(", ")}` +
                    (catCambios.length > 3 ? ` y ${catCambios.length - 3} cambio(s) más` : "");
    await guardarTodo({ bufferBase: buffer, mapeo: catMapeo, estado, mensaje });

    // el ciclo mensual tiene que seguir con el archivo nuevo, no con el de antes
    bufferBase = buffer;
    mapeoGuardado = JSON.parse(JSON.stringify(catMapeo));
    const cuantos = catCambios.length;
    catCambios = [];
    renderCategorias();
    st.innerHTML = `<div class="status-msg ok">Listo: ${cuantos} cambio(s) guardados.</div>`;
  } catch (e) {
    st.innerHTML = `<div class="status-msg bad">No se pudo guardar: ${catEsc(e.message)}</div>`;
  } finally {
    mostrar("spinnerCat", false);
    btn.disabled = catCambios.length === 0;
  }
}

function catCambiarBusqueda(v) {
  catBusqueda = v;
  if (v.trim()) catListaAbierta = true;
  renderCategorias();
}
