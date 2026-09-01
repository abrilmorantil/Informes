// Panel de configuración de cuentas — Dólares.
//
// Hermano del de pesos (`panel_config_cuentas.js`) pero NO es el mismo, porque los balances no
// tienen la misma forma. Lo que cambia, medido en los dos maestros:
//
//   - En pesos una línea de gasto es una CUENTA MADRE con subcuentas adentro, y el Anexo II lee
//     el subtotal de la madre. En dólares no hay subcuentas: cada línea ya es la madre (las
//     hijas del export se suman ahí), y su celda es un VLOOKUP pelado. Por eso
//     `madresResultados` devuelve 0 en dólares y acá la unidad son las 125 líneas de
//     RESULTADOS, no las madres. Con eso `a2Mapa`/`a2Verificar` funcionan igual.
//   - La geometría: la clave va en la columna B y el importe en la C (en pesos, C..E y G).
//
// Lo que sí se reusa tal cual: el chip de pendientes, el detalle de pendientes, el bloque de
// líneas sin cuenta asignada y todo `anexo2.js`. Nombres con prefijo `pcd` por el ámbito
// global único que comparten los scripts del sitio.

const PCD_COL_SALDOS = "C";     // de dónde lee el Anexo II de dólares

// ------------------------------------------------------------------ lectura del maestro

// Las líneas de RESULTADOS: las que el Anexo II puede repartir en sus cuatro columnas.
// El capítulo sale del primer dígito del código, igual que en el resto del motor.
function pcdLineasResultados(wb, mapeo) {
  const p = PARAMS.dolares;
  const ws = wb.getWorksheet("SALDOS");
  const out = [];
  for (const [cod, info] of Object.entries(mapeo.cuentas || {})) {
    if (!String(cod).startsWith("4")) continue;
    let txt = "";
    for (const c of p.saldosColsCuenta) {
      const t = pcdTexto(ws.getCell(info.fila, c).value).trim();
      if (t) { txt = t; break; }
    }
    const m = /^\s*(\d{5,})\s*-\s*(.+?)\s*$/.exec(txt);
    out.push({ fila: info.fila, codigo: String(cod), nombre: m ? m[2] : txt });
  }
  out.sort((a, b) => a.fila - b.fila);
  return out;
}

function pcdTexto(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return typeof v.result === "string" ? v.result : "";
    return v.result === undefined || v.result === null ? "" : String(v.result);
  }
  return String(v);
}

// Quién lee cada fila de SALDOS fuera del Anexo II.
//
// Hace falta para no dar un aviso falso: una línea puede estar perfectamente bien fuera del
// Anexo II porque la lee `Resultados` directo — es el caso de la diferencia de cambio, el ROBO
// y los otros ingresos. La que no la lee NADIE es la única que está de verdad suelta.
//
// El total de la hoja (`SUM(C2:C232)`) no cuenta como lector: abarca todas las filas y diría
// que todas están bien.
function pcdLectoresFuera(wb) {
  const porFila = new Map();
  const RE = /SALDOS!\$?[A-Z]{1,3}\$?(\d+)/gi;
  for (const ws of wb.worksheets) {
    if (ws.name === "Anexo II" || ws.name === "Hoja1") continue;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, ci) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && v.formula)) return;
        const f = String(v.formula);
        if (/^\s*\+?SUM\(/i.test(f)) return;               // totales de la hoja
        RE.lastIndex = 0;
        let m;
        while ((m = RE.exec(f)) !== null) {
          const fila = +m[1];
          if (ws.name === "SALDOS" && r === fila) continue;
          let et = "";
          for (let k = 1; k < ci; k++) { const t = pcdTexto(ws.getCell(r, k).value).trim(); if (t) et = t; }
          if (!porFila.has(fila)) porFila.set(fila, []);
          porFila.get(fila).push({ hoja: ws.name, etiqueta: et });
        }
      });
    });
  }
  return porFila;
}

// ------------------------------------------------------------------ estado, en memoria

let pcdWb = null;
let pcdMapeo = null;
let pcdCfg = null;
let pcdAnexo = null;
let pcdOnvio = [];
let pcdSueltas = [];           // líneas de RESULTADOS que no entran a ningún estado
let pcdCambios = [];

function pcdLog(msg) {
  pcdCambios.push(msg);
  const el = $("pcdLog");
  el.textContent = pcdCambios.join("\n");
  mostrar("pcdLog", true);
  mostrar("pcdGuardarBar", true);
  $("pcdGuardarEstado").textContent = `${pcdCambios.length} cambio(s) sin guardar.`;
}

function pcdRecalcular() {
  pcdMapeo = derivarMapeoMaestro(pcdWb, "dolares", clasificacion);
  pcdCfg = derivarConfigBalance(pcdWb, "dolares", pcdMapeo, lineasDeNota4(pcdWb),
    PARAMS.dolares, { filasQueAgrega });
  pcdOnvio = pccCuentasDeHoja1(pcdWb, pcdMapeo);

  const lineas = pcdLineasResultados(pcdWb, pcdMapeo);
  const verificacion = a2Verificar(pcdWb, lineas);
  pcdAnexo = {
    lineas: a2Mapa(pcdWb, lineas).lineas,
    conceptos: a2Conceptos(pcdWb),
    columnas: a2Columnas(pcdWb),
    verificacion,
    madres: lineas,
  };

  const lectores = pcdLectoresFuera(pcdWb);
  pcdSueltas = verificacion.sinAnexo
    .filter(l => !(lectores.get(l.filaSaldos) || []).length)
    .map(l => ({ ...l, lectores: [] }));
}

async function verConfigCuentasDolares() {
  mostrar("cardConfigDolares", true);
  $("cardConfigDolares").scrollIntoView({ behavior: "smooth", block: "start" });
  if (pcdWb) { pcdRender(); return; }
  if (!bufferBaseUsd) {
    $("pcdResumen").innerHTML =
      `<div class="status-msg bad">Todavía no está cargado el maestro de dólares.</div>`;
    return;
  }
  try {
    pcdWb = await abrirWorkbook(bufferBaseUsd.slice(0));
    pcdRecalcular();
    pcdRender();
  } catch (err) {
    $("pcdResumen").innerHTML =
      `<div class="status-msg bad">No pude leer el maestro de dólares: ${gcEsc(err.message)}</div>`;
  }
}

function cerrarConfigDolares() { mostrar("cardConfigDolares", false); }

// ------------------------------------------------------------------ dibujo

// Las líneas que no entran al Anexo II Y que tampoco lee ninguna otra hoja: su importe no
// aparece en ningún estado. No es lo mismo que estar fuera del anexo, que puede estar bien.
function pcdSueltasHtml(sueltas) {
  if (!sueltas.length) {
    return `<div class="status-msg ok">Cada línea de gastos entra al Anexo II o la lee el
      estado de resultados. Ninguna queda sin aparecer en ningún lado.</div>`;
  }
  return `
    <div class="status-msg bad">
      <b>${sueltas.length} línea(s) no aparecen en ningún estado</b> — no entran a ninguna
      columna del Anexo II y tampoco las lee el estado de resultados, así que su importe no
      se ve en ninguna parte. Hoy están en cero; el día que muevan, el balance no lo va a
      mostrar.
      ${sueltas.map(s => `<br>· <b>${gcEsc(s.codigo)}</b> — ${gcEsc(s.nombre)}
        (fila ${s.filaSaldos} de SALDOS)`).join("")}
    </div>`;
}

function pcdAnexoHtml() {
  if (!pcdAnexo) return "";
  const { lineas, conceptos, columnas, verificacion } = pcdAnexo;
  const q = (($("pcdBuscar") && $("pcdBuscar").value) || "").trim().toLowerCase();
  const visibles = lineas.filter(l => !q ||
    l.nombre.toLowerCase().includes(q) || String(l.codigo).includes(q) ||
    l.donde.some(d => d.concepto.toLowerCase().includes(q)));

  const porColumna = {};
  for (const l of visibles) {
    const col = l.donde.length ? l.donde[0].columna : "(fuera del Anexo II)";
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
        <span class="gc-n">${ls.length} línea${ls.length === 1 ? "" : "s"}</span>
      </summary>
      <div class="gc-body">
        <div class="gc-miembro a2-encabezado">
          <div class="gc-nom">Cuenta</div>
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
              <select class="pcd-concepto" data-fila="${l.filaSaldos}">${opcConcepto}</select>
              <select class="pcd-columna" data-fila="${l.filaSaldos}">${opcColumna}</select>
              <button data-accion="pcd-mover" data-fila="${l.filaSaldos}">Mover</button>
            </div>
          </div>`;
        }).join("")}
      </div>
    </details>`).join("");

  return `
    <h3 style="font-family:'Instrument Serif',serif; font-weight:400; font-size:20px; margin:26px 0 4px;">
      Anexo II — en qué columna cae cada gasto</h3>
    <p class="footer-note" style="margin-top:0;">
      A diferencia del balance en pesos, acá cada línea es una cuenta suelta: no hay subcuentas
      que se sumen adentro. Se deciden dos cosas: el <b>concepto</b> —el renglón con el que sale
      impreso— y el <b>tipo de gasto</b>, que es la columna: Administración, Comercialización,
      Exploración o Financieros.
    </p>
    <div class="status-msg ${verificacion.dobles.length ? "bad" : "ok"}">
      ${verificacion.unaVez} de ${verificacion.total} líneas entran al Anexo II exactamente una
      vez${verificacion.dobles.length
        ? `, pero <b>${verificacion.dobles.length} entran más de una</b> y se cuentan doble: ` +
          verificacion.dobles.map(d => gcEsc(d.nombre)).join(", ")
        : ". Ninguna se cuenta doble."}
    </div>
    ${pcdSueltasHtml(pcdSueltas)}
    <div style="margin-top:12px;">${grupos || '<p class="footer-note">Nada coincide con la búsqueda.</p>'}</div>`;
}

function pcdRender() {
  const pendientes = pcdCfg.avisos.filter(a => a.tipo !== "linea_sin_cuenta");
  $("pcdResumen").innerHTML = pccResumenHtml(pcdCfg, pendientes.length, 0);
  $("pcdPendientes").innerHTML = pccPendientesHtml(pcdCfg) + pccSinCuentaHtml(pcdCfg, pcdOnvio);
  $("pcdAnexo").innerHTML = pcdAnexoHtml();
  pcdSincronizarSelects();
}

// Deja los desplegables mostrando dónde está HOY cada línea; poner `selected` en el HTML se
// pierde al re-renderizar.
function pcdSincronizarSelects() {
  if (!pcdAnexo) return;
  for (const l of pcdAnexo.lineas) {
    const d = l.donde[0];
    if (!d) continue;
    const sc = document.querySelector(`.pcd-concepto[data-fila="${l.filaSaldos}"]`);
    const sl = document.querySelector(`.pcd-columna[data-fila="${l.filaSaldos}"]`);
    if (sc) sc.value = String(d.anexoFila);
    if (sl) sl.value = String(d.col);
  }
}

// ------------------------------------------------------------------ acciones

// En Node (los tests) no hay document: el módulo se carga igual para probar sus funciones puras.
if (typeof document !== "undefined") {
document.addEventListener("DOMContentLoaded", () => {
  const cont = $("cardConfigDolares");
  if (!cont) return;

  const buscar = $("pcdBuscar");
  if (buscar) buscar.addEventListener("input", () => { $("pcdAnexo").innerHTML = pcdAnexoHtml(); pcdSincronizarSelects(); });

  cont.addEventListener("input", (e) => {
    const campo = e.target.closest("input.sc-cod");
    if (!campo) return;
    const cuenta = pcdOnvio.find(c => c.codigo === campo.value.trim());
    if (cuenta) $(`cc_nom_${campo.dataset.fila}`).value = cuenta.nombre;
  });

  cont.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-accion]");
    if (!btn) return;
    e.preventDefault();
    const fila = btn.dataset.fila ? +btn.dataset.fila : null;
    try {
      if (btn.dataset.accion === "pcd-mover") {
        const sc = document.querySelector(`.pcd-concepto[data-fila="${fila}"]`);
        const sl = document.querySelector(`.pcd-columna[data-fila="${fila}"]`);
        a2Mover({
          wb: pcdWb, madres: pcdAnexo.madres, filaSaldos: fila,
          anexoFilaDestino: +sc.value, colDestino: +sl.value,
          colSaldosDefecto: PCD_COL_SALDOS, log: pcdLog,
        });
        pcdRecalcular();
        pcdRender();
      } else if (btn.dataset.accion === "cc-corregir") {
        const cod = $(`cc_cod_${fila}`).value.trim();
        const nom = $(`cc_nom_${fila}`).value.trim();
        if (!nom) throw new Error("El nombre no puede quedar vacío.");
        corregirColumnaDeCuenta(pcdWb, fila, cod, nom, pcdLog);
        pcdRecalcular();
        pcdRender();
      }
    } catch (err) {
      alert(err.message);
    }
  });
});
}

async function guardarCambiosDolares() {
  if (!pcdWb || !pcdCambios.length) return;
  const boton = $("btnPcdGuardar");
  boton.disabled = true;
  $("pcdGuardarEstado").innerHTML = `<span class="spinner">Guardando en GitHub…</span>`;
  try {
    const buffer = await pcdWb.xlsx.writeBuffer();
    await ghcGuardarBaseUsd(buffer, "Configuración de cuentas (dólares): " + pcdCambios.length + " cambio(s)");
    bufferBaseUsd = buffer;
    $("pcdGuardarEstado").innerHTML =
      `<div class="status-msg ok" style="margin:0;">Guardado. La próxima corrida ya usa este maestro.</div>`;
    pcdCambios = [];
  } catch (err) {
    $("pcdGuardarEstado").innerHTML =
      `<div class="status-msg bad" style="margin:0;">No pude guardar: ${gcEsc(err.message)}</div>`;
  } finally {
    boton.disabled = false;
  }
}

if (typeof module !== "undefined") {
  module.exports = { pcdLineasResultados, pcdLectoresFuera, pcdSueltasHtml, pcdTexto };
}
