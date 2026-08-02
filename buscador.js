// Buscador para las listas largas de los tres informes.
//
// Cuando hay que elegir una cuenta —la categoría en el Balance USD, la cuenta madre o la
// línea de la Nota 4 en el de pesos, y sobre todo el destino en el de dólares, que tiene
// 203 opciones— desplegar la lista entera y buscar a ojo es incómodo y da lugar a errores.
//
// Antes esto era un `<select>` nativo con un casillero de texto ARRIBA: había que escribir
// en un lado y elegir en otro, y no quedaba claro que una cosa filtrara a la otra. Ahora es
// un solo control: se abre el desplegable y se busca adentro, que es donde uno lo espera.
//
// El `<select>` real sigue existiendo, escondido, y es el que se actualiza al elegir. Así
// `.value` sigue respondiendo igual que siempre y nada del código que lo lee cambió.
//
// Detalles que importan:
//   - Se busca sin acentos y por partes sueltas: "alq camp" encuentra
//     "ALQUILERES DE EQUIPOS DE CAMPO".
//   - Al elegir se dispara `change` en el select, para no romper a quien lo escuche.
//   - La etiqueta del botón se refresca también al abrir: si el código cambió `.value` por
//     su cuenta (las propuestas vienen pre-elegidas), el botón lo muestra igual.
//   - Con el teclado: flechas para moverse, Enter para elegir, Escape para cerrar.

function normBuscador(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

function conBuscador(select, placeholder) {
  if (!select || select.dataset.conBuscador === "1") return select;
  select.dataset.conBuscador = "1";

  const todas = [...select.options].map(o => ({
    value: o.value, text: o.text, disabled: o.disabled,
    busca: normBuscador(o.text + " " + o.value),
  }));
  const vacia = todas.find(o => o.value === "");
  const reales = todas.filter(o => o.value !== "" && !o.disabled);

  const caja = document.createElement("div");
  caja.className = "combo";
  const boton = document.createElement("button");
  boton.type = "button";
  boton.className = "combo-boton";
  boton.setAttribute("aria-haspopup", "listbox");
  boton.setAttribute("aria-expanded", "false");
  const etiqueta = document.createElement("span");
  etiqueta.className = "combo-valor";
  boton.appendChild(etiqueta);
  boton.insertAdjacentHTML("beforeend", '<span class="combo-flecha" aria-hidden="true">▾</span>');

  const panel = document.createElement("div");
  panel.className = "combo-panel hidden";
  const input = document.createElement("input");
  input.type = "text";
  input.className = "combo-buscar";
  input.placeholder = placeholder || "Buscar…";
  input.autocomplete = "off";
  const lista = document.createElement("div");
  lista.className = "combo-lista";
  lista.setAttribute("role", "listbox");
  const aviso = document.createElement("div");
  aviso.className = "combo-aviso hidden";
  panel.appendChild(input);
  panel.appendChild(lista);
  panel.appendChild(aviso);

  select.parentNode.insertBefore(caja, select);
  caja.appendChild(boton);
  caja.appendChild(panel);
  caja.appendChild(select);
  select.classList.add("combo-real");

  let marcada = -1;

  const textoDe = (v) => {
    const o = todas.find(x => x.value === v);
    return o ? o.text : "";
  };

  function pintarBoton() {
    const v = select.value;
    const t = textoDe(v);
    const sinElegir = !v || (vacia && v === "");
    etiqueta.textContent = sinElegir ? (vacia ? vacia.text : "Elegí una opción") : t;
    boton.classList.toggle("sin-elegir", !!sinElegir);
  }

  function opcionesVisibles() {
    const partes = normBuscador(input.value).split(" ").filter(Boolean);
    if (!partes.length) return reales;
    return reales.filter(o => partes.every(p => o.busca.includes(p)));
  }

  function pintarLista() {
    const visibles = opcionesVisibles();
    lista.innerHTML = "";
    visibles.forEach((o, i) => {
      const fila = document.createElement("div");
      fila.className = "combo-opcion";
      fila.setAttribute("role", "option");
      fila.dataset.value = o.value;
      fila.textContent = o.text;
      if (o.value === select.value) fila.classList.add("elegida");
      if (i === marcada) fila.classList.add("marcada");
      fila.addEventListener("mousedown", (ev) => { ev.preventDefault(); elegir(o.value); });
      lista.appendChild(fila);
    });
    const hayFiltro = normBuscador(input.value).length > 0;
    aviso.textContent = visibles.length === 0
      ? "Ninguna opción coincide con ese texto."
      : `${visibles.length} de ${reales.length} opciones`;
    aviso.classList.toggle("hidden", !hayFiltro);
    aviso.classList.toggle("vacio", visibles.length === 0);

    const marca = lista.querySelector(".marcada");
    if (marca && marca.scrollIntoView) marca.scrollIntoView({ block: "nearest" });
  }

  function abrir() {
    panel.classList.remove("hidden");
    boton.setAttribute("aria-expanded", "true");
    input.value = "";
    // la elegida arranca marcada, así Enter no cambia nada sin querer
    marcada = reales.findIndex(o => o.value === select.value);
    pintarLista();
    input.focus();
  }

  function cerrar() {
    panel.classList.add("hidden");
    boton.setAttribute("aria-expanded", "false");
  }

  function elegir(valor) {
    select.value = valor;
    pintarBoton();
    cerrar();
    select.dispatchEvent(new Event("change", { bubbles: true }));
    boton.focus();
  }

  boton.addEventListener("click", () => {
    if (panel.classList.contains("hidden")) { pintarBoton(); abrir(); } else cerrar();
  });

  input.addEventListener("input", () => { marcada = 0; pintarLista(); });

  input.addEventListener("keydown", (ev) => {
    const visibles = opcionesVisibles();
    if (ev.key === "Escape") { ev.preventDefault(); cerrar(); boton.focus(); return; }
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
      ev.preventDefault();
      if (!visibles.length) return;
      const paso = ev.key === "ArrowDown" ? 1 : -1;
      marcada = (marcada + paso + visibles.length) % visibles.length;
      pintarLista();
      return;
    }
    if (ev.key === "Enter") {
      ev.preventDefault();
      if (visibles.length === 1) return elegir(visibles[0].value);
      if (marcada >= 0 && visibles[marcada]) return elegir(visibles[marcada].value);
    }
  });

  // click afuera cierra, sin elegir nada
  document.addEventListener("mousedown", (ev) => {
    if (!panel.classList.contains("hidden") && !caja.contains(ev.target)) cerrar();
  });

  // si alguien cambia el select por código y avisa, el botón lo sigue
  select.addEventListener("change", pintarBoton);

  pintarBoton();
  return select;
}

// Aplica el buscador a todos los selects que coincidan con el selector.
function conBuscadorTodos(selector, placeholder) {
  document.querySelectorAll(selector).forEach(s => conBuscador(s, placeholder));
}

if (typeof module !== "undefined") {
  module.exports = { conBuscador, conBuscadorTodos, normBuscador };
}
