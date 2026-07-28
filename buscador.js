// Buscador para las listas largas de los tres informes.
//
// Cuando hay que elegir una cuenta —la categoría en el Balance USD, la cuenta madre o la
// línea de la Nota 4 en el de pesos, y sobre todo el destino en el de dólares, que tiene
// 203 opciones— desplegar la lista entera y buscar a ojo es incómodo y da lugar a errores.
//
// Se le agrega arriba un campo de texto que filtra las opciones a medida que se escribe.
// El `<select>` sigue siendo el mismo elemento y sigue respondiendo `.value` igual que
// antes, así que nada del código que lo lee necesita cambiar.
//
// Detalles que importan:
//   - Las opciones no se ocultan con CSS (Safari ignora `display:none` en un `<option>`):
//     se guarda la lista completa y se reconstruye el select en cada búsqueda.
//   - La opción elegida nunca desaparece por el filtro, para no perder la selección.
//   - Se busca sin acentos y por partes sueltas: "alq camp" encuentra
//     "ALQUILERES DE EQUIPOS DE CAMPO".

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

  const caja = document.createElement("div");
  caja.className = "buscador";
  const input = document.createElement("input");
  input.type = "text";
  input.className = "buscador-input";
  input.placeholder = placeholder || "Buscar…";
  input.autocomplete = "off";
  const aviso = document.createElement("div");
  aviso.className = "buscador-aviso hidden";

  select.parentNode.insertBefore(caja, select);
  caja.appendChild(input);
  caja.appendChild(select);
  caja.appendChild(aviso);

  const totalReales = todas.filter(o => o.value !== "").length;

  function filtrar() {
    const partes = normBuscador(input.value).split(" ").filter(Boolean);
    const elegido = select.value;
    const coincide = (o) => o.value !== "" && partes.every(p => o.busca.includes(p));
    // La elegida se muestra siempre, coincida o no, para no perder la selección; pero
    // NO se cuenta como resultado, o el aviso diría "1 de 203" sin haber encontrado nada.
    const visibles = todas.filter(o => o.value === "" || o.value === elegido || coincide(o));
    const encontradas = todas.filter(coincide).length;

    select.innerHTML = "";
    for (const o of visibles) {
      const op = document.createElement("option");
      op.value = o.value;
      op.textContent = o.text;
      if (o.disabled) op.disabled = true;
      select.appendChild(op);
    }
    select.value = elegido;

    aviso.textContent = encontradas === 0
      ? "Ninguna opción coincide con ese texto."
      : `${encontradas} de ${totalReales} opciones`;
    aviso.classList.toggle("hidden", partes.length === 0);
    aviso.classList.toggle("vacio", encontradas === 0);
  }

  input.addEventListener("input", filtrar);
  // Enter no envía nada: si quedó una sola opción, la elige y listo.
  input.addEventListener("keydown", (ev) => {
    if (ev.key !== "Enter") return;
    ev.preventDefault();
    const reales = [...select.options].filter(o => o.value !== "");
    if (reales.length === 1) { select.value = reales[0].value; select.focus(); }
  });

  return select;
}

// Aplica el buscador a todos los selects que coincidan con el selector.
function conBuscadorTodos(selector, placeholder) {
  document.querySelectorAll(selector).forEach(s => conBuscador(s, placeholder));
}

if (typeof module !== "undefined") {
  module.exports = { conBuscador, conBuscadorTodos, normBuscador };
}
