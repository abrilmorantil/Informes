// Guarda y lee en GitHub para el Informe C (Balance Pesos y Dólares):
//   - base_pesos.xlsx : el maestro de pesos vivo
//   - estado_c.json   : historial de cargas
//
// Clave de configuración propia: no pisa la del Informe A ni la de BALCOMPROBDOLARES.
// Trae de nacimiento los arreglos aprendidos en el Informe A: lecturas sin caché
// (la Contents API responde con max-age=60 y el navegador reusa lo viejo) y
// reintento ante un 409 releyendo el sha real.

const GHC_SETTINGS_KEY = "informe_c_gh_settings";
const GHC_CARPETA_DEFECTO = "informe-c";
const GHC_ARCHIVO_BASE = "base_pesos.xlsx";
const GHC_ARCHIVO_ESTADO = "estado_c.json";
const GHC_ARCHIVO_BASE_USD = "base_dolares.xlsx";
const GHC_ARCHIVO_EQUIV = "equivalencias_dolares.json";

function loadGhcSettings() {
  try { return JSON.parse(localStorage.getItem(GHC_SETTINGS_KEY) || "{}"); }
  catch (e) { return {}; }
}
function saveGhcSettings(s) { localStorage.setItem(GHC_SETTINGS_KEY, JSON.stringify(s)); }
function hasGhcSettings() { const s = loadGhcSettings(); return !!(s.token && s.repo); }

function ghcRuta(nombre) {
  const c = (loadGhcSettings().carpeta || GHC_CARPETA_DEFECTO).replace(/^\/+|\/+$/g, "");
  return c ? `${c}/${nombre}` : nombre;
}

function ghcUtf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let b = "";
  bytes.forEach(x => { b += String.fromCharCode(x); });
  return btoa(b);
}
function ghcBase64ToUtf8(b64) {
  const bin = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}
function ghcBufferABase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let b = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    b += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(b);
}
function ghcBase64ABuffer(b64) {
  const bin = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function ghcCabeceras() {
  const s = loadGhcSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  return { "Authorization": `token ${s.token}`, "Accept": "application/vnd.github+json" };
}

async function ghcLeer(nombre) {
  const s = loadGhcSettings();
  const ruta = ghcRuta(nombre);
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
              `?ref=${encodeURIComponent(s.rama || "main")}&_=${Date.now()}`;
  const res = await fetch(url, { headers: ghcCabeceras(), cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`No pude leer ${ruta} de GitHub (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return { contenidoBase64: data.content, sha: data.sha };
}

function ghcPut(ruta, contenidoBase64, sha, mensaje, rama) {
  const s = loadGhcSettings();
  const cuerpo = { message: mensaje, content: contenidoBase64, branch: rama };
  if (sha) cuerpo.sha = sha;
  return fetch(`https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}`, {
    method: "PUT",
    headers: { ...ghcCabeceras(), "Content-Type": "application/json" },
    body: JSON.stringify(cuerpo),
    cache: "no-store",
  });
}

async function ghcEscribir(nombre, contenidoBase64, sha, mensaje) {
  const s = loadGhcSettings();
  const ruta = ghcRuta(nombre);
  const rama = s.rama || "main";
  let res = await ghcPut(ruta, contenidoBase64, sha, mensaje, rama);
  if (res.status === 409) {
    const actual = await ghcLeer(nombre);
    res = await ghcPut(ruta, contenidoBase64, actual ? actual.sha : undefined, mensaje, rama);
    if (res.status === 409) {
      throw new Error(
        `${ruta} cambió en GitHub mientras se guardaba, así que no lo piso. ` +
        `Actualizá la página (Ctrl+Shift+R) y volvé a intentarlo. NO se guardó nada.`
      );
    }
  }
  if (!res.ok) throw new Error(`No pude guardar ${ruta} en GitHub (${res.status}): ${await res.text()}`);
  return await res.json();
}

// El mapping del Informe B vive en la RAÍZ del repositorio (su github.js usa
// `s.path || "mapping.json"`), no dentro de la carpeta del Informe C, así que se lo
// pide por ruta absoluta. Si no está, se sigue sin él: la app pregunta lo que no sabe.
async function ghcLeerMappingB() {
  const s = loadGhcSettings();
  for (const ruta of ["mapping.json", "informe-b/mapping.json"]) {
    const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
                `?ref=${encodeURIComponent(s.rama || "main")}&_=${Date.now()}`;
    let res;
    try { res = await fetch(url, { headers: ghcCabeceras(), cache: "no-store" }); }
    catch (e) { return null; }
    if (res.status === 404) continue;
    if (!res.ok) return null;
    const data = await res.json();
    try { return JSON.parse(ghcBase64ToUtf8(data.content)); } catch (e) { return null; }
  }
  return null;
}

async function ghcLeerEstado() {
  const r = await ghcLeer(GHC_ARCHIVO_ESTADO);
  return r ? { estado: JSON.parse(ghcBase64ToUtf8(r.contenidoBase64)), sha: r.sha } : null;
}
async function ghcLeerBase() {
  const r = await ghcLeer(GHC_ARCHIVO_BASE);
  return r ? { buffer: ghcBase64ABuffer(r.contenidoBase64), sha: r.sha } : null;
}
async function ghcLeerBaseUsd() {
  const r = await ghcLeer(GHC_ARCHIVO_BASE_USD);
  return r ? { buffer: ghcBase64ABuffer(r.contenidoBase64), sha: r.sha } : null;
}
// Las equivalencias de dólares que la usuaria fue confirmando. Cada decisión se toma
// UNA vez: acá quedan guardadas para todas las corridas siguientes.
async function ghcLeerEquivalencias() {
  const r = await ghcLeer(GHC_ARCHIVO_EQUIV);
  if (!r) return {};
  try { return JSON.parse(ghcBase64ToUtf8(r.contenidoBase64)); } catch (e) { return {}; }
}
async function ghcGuardarEquivalencias(equiv, mensaje) {
  const sha = (await ghcLeer(GHC_ARCHIVO_EQUIV))?.sha;
  await ghcEscribir(GHC_ARCHIVO_EQUIV, ghcUtf8ToBase64(JSON.stringify(equiv, null, 1)), sha, mensaje);
}
async function ghcGuardarBaseUsd(buffer, mensaje) {
  const sha = (await ghcLeer(GHC_ARCHIVO_BASE_USD))?.sha;
  await ghcEscribir(GHC_ARCHIVO_BASE_USD, ghcBufferABase64(buffer), sha, mensaje);
}

// El maestro primero (es el que más pesa y el más probable que falle): si algo se
// corta, el estado sigue apuntando a la versión anterior y se puede reintentar.
async function ghcGuardarTodo({ bufferBase, estado, mensaje }) {
  const shaBase = (await ghcLeer(GHC_ARCHIVO_BASE))?.sha;
  await ghcEscribir(GHC_ARCHIVO_BASE, ghcBufferABase64(bufferBase), shaBase, mensaje);
  const shaEstado = (await ghcLeer(GHC_ARCHIVO_ESTADO))?.sha;
  await ghcEscribir(GHC_ARCHIVO_ESTADO, ghcUtf8ToBase64(JSON.stringify(estado, null, 1)), shaEstado, mensaje);
}

if (typeof module !== "undefined") {
  module.exports = {
    loadGhcSettings, saveGhcSettings, hasGhcSettings,
    ghcLeer, ghcEscribir, ghcLeerEstado, ghcLeerBase, ghcGuardarTodo, ghcLeerMappingB,
    ghcLeerBaseUsd, ghcGuardarBaseUsd, ghcLeerEquivalencias, ghcGuardarEquivalencias,
  };
}
