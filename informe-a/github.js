// Guarda y lee en el repositorio de GitHub, con la Contents API:
//   - base_actual.xlsx : el balance vivo, con los meses cerrados congelados
//   - mapeo.json       : dónde está cada cuenta y cada centro de costo
//   - estado.json      : cuál fue el último mes cerrado
//
// Usa su propia clave de configuración: la de BALCOMPROBDOLARES apunta a otro
// archivo y no se debe pisar.

const GH_SETTINGS_KEY = "informe_a_gh_settings";
const CARPETA_DEFECTO = "informe-a";

const ARCHIVO_BASE = "base_actual.xlsx";
const ARCHIVO_MAPEO = "mapeo.json";
const ARCHIVO_ESTADO = "estado.json";

function loadGhSettings() {
  try {
    return JSON.parse(localStorage.getItem(GH_SETTINGS_KEY) || "{}");
  } catch (e) {
    return {};
  }
}

function saveGhSettings(settings) {
  localStorage.setItem(GH_SETTINGS_KEY, JSON.stringify(settings));
}

function hasGhSettings() {
  const s = loadGhSettings();
  return !!(s.token && s.repo);
}

function carpeta() {
  const s = loadGhSettings();
  return (s.carpeta || CARPETA_DEFECTO).replace(/^\/+|\/+$/g, "");
}

function rutaDe(nombre) {
  const c = carpeta();
  return c ? `${c}/${nombre}` : nombre;
}

// btoa se come cualquier carácter fuera de latin1, así que los acentos y la ñ
// se codifican pasando por UTF-8 explícitamente.
function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let binario = "";
  bytes.forEach(b => { binario += String.fromCharCode(b); });
  return btoa(binario);
}

function base64ToUtf8(b64) {
  const binario = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(binario.length);
  for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

// Los .xlsx son binarios: se pasan a base64 de a pedazos para no reventar la pila
// con archivos de cientos de miles de bytes.
function bufferABase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binario = "";
  const PASO = 0x8000;
  for (let i = 0; i < bytes.length; i += PASO) {
    binario += String.fromCharCode.apply(null, bytes.subarray(i, i + PASO));
  }
  return btoa(binario);
}

function base64ABuffer(b64) {
  const binario = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(binario.length);
  for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
  return bytes.buffer;
}

function cabeceras() {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  return {
    "Authorization": `token ${s.token}`,
    "Accept": "application/vnd.github+json",
  };
}

// Devuelve null si el archivo todavía no existe (primera vez).
//
// GitHub manda "Cache-Control: private, max-age=60" en las respuestas de su API, así
// que el navegador durante un minuto reusa lo que ya había leído sin volver a
// preguntar. Justo después de guardar eso hace daño por partida doble: la página se
// dibuja con el estado viejo (pedía cargar un mes que ya se había cerrado), y al
// reintentar se manda el sha viejo, que GitHub rechaza con un 409. Por eso cada
// lectura lleva no-store y un parámetro distinto, para que siempre vaya a buscar lo
// que hay de verdad en el repositorio.
async function ghLeer(nombre) {
  const s = loadGhSettings();
  const ruta = rutaDe(nombre);
  const rama = s.rama || "main";
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
              `?ref=${encodeURIComponent(rama)}&_=${Date.now()}`;
  const res = await fetch(url, { headers: cabeceras(), cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`No pude leer ${ruta} de GitHub (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return { contenidoBase64: data.content, sha: data.sha };
}

function ghPut(ruta, contenidoBase64, sha, mensaje, rama) {
  const s = loadGhSettings();
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}`;
  const cuerpo = { message: mensaje, content: contenidoBase64, branch: rama };
  if (sha) cuerpo.sha = sha;
  return fetch(url, {
    method: "PUT",
    headers: { ...cabeceras(), "Content-Type": "application/json" },
    body: JSON.stringify(cuerpo),
    cache: "no-store",
  });
}

// El 409 significa que el archivo en GitHub ya no es el que se leyó: el sha no coincide.
// Se vuelve a leer el sha real y se reintenta una vez, que resuelve el caso normal (una
// lectura que quedó vieja). Si vuelve a fallar, alguien más lo cambió de verdad y hay que
// avisar en vez de pisarlo.
async function ghEscribir(nombre, contenidoBase64, sha, mensaje) {
  const s = loadGhSettings();
  const ruta = rutaDe(nombre);
  const rama = s.rama || "main";

  let res = await ghPut(ruta, contenidoBase64, sha, mensaje, rama);
  if (res.status === 409) {
    const actual = await ghLeer(nombre);
    res = await ghPut(ruta, contenidoBase64, actual ? actual.sha : undefined, mensaje, rama);
    if (res.status === 409) {
      throw new Error(
        `${ruta} cambió en GitHub mientras se guardaba, así que no lo piso. ` +
        `Actualizá la página (Ctrl+Shift+R) para ver cómo quedó y, si hace falta, ` +
        `volvé a intentarlo. NO se guardó nada.`
      );
    }
  }
  if (!res.ok) throw new Error(`No pude guardar ${ruta} en GitHub (${res.status}): ${await res.text()}`);
  return await res.json();
}

// ---- las tres cosas que se guardan ----

async function leerEstado() {
  const r = await ghLeer(ARCHIVO_ESTADO);
  return r ? { estado: JSON.parse(base64ToUtf8(r.contenidoBase64)), sha: r.sha } : null;
}

async function leerMapeo() {
  const r = await ghLeer(ARCHIVO_MAPEO);
  return r ? { mapeo: JSON.parse(base64ToUtf8(r.contenidoBase64)), sha: r.sha } : null;
}

async function leerBase() {
  const r = await ghLeer(ARCHIVO_BASE);
  return r ? { buffer: base64ABuffer(r.contenidoBase64), sha: r.sha } : null;
}

// Guarda las tres cosas juntas. El orden importa: el archivo base primero, porque
// es el que más pesa y el más probable que falle; si algo se corta, el estado sigue
// apuntando al mes anterior y se puede reintentar sin quedar a mitad de camino.
async function guardarTodo({ bufferBase, mapeo, estado, mensaje }) {
  const shaBase = (await ghLeer(ARCHIVO_BASE))?.sha;
  await ghEscribir(ARCHIVO_BASE, bufferABase64(bufferBase), shaBase, mensaje);

  const shaMapeo = (await ghLeer(ARCHIVO_MAPEO))?.sha;
  await ghEscribir(ARCHIVO_MAPEO, utf8ToBase64(JSON.stringify(mapeo, null, 1)), shaMapeo, mensaje);

  const shaEstado = (await ghLeer(ARCHIVO_ESTADO))?.sha;
  await ghEscribir(ARCHIVO_ESTADO, utf8ToBase64(JSON.stringify(estado, null, 1)), shaEstado, mensaje);
}

if (typeof module !== "undefined") {
  module.exports = {
    ghLeer, ghEscribir, guardarTodo, leerEstado, leerMapeo, leerBase,
    utf8ToBase64, base64ToUtf8, bufferABase64, base64ABuffer,
  };
}
