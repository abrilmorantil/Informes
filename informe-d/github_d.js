// Guarda en GitHub la configuración del asiento de diferencia de cambio:
//   - config_difcambio.json : prefijos y cuentas no monetarias, prefijo de proveedores,
//                             cuenta de balanceo, umbrales y la allowlist de monetarias
//                             ya confirmadas por la usuaria.
//   - estado_d.json         : historial de asientos generados.
//
// Mismo mecanismo que los otros informes, con los arreglos ya aprendidos: lecturas sin
// caché (la Contents API responde con max-age=60 y el navegador reusa lo viejo) y
// reintento ante un 409 releyendo el sha real.

const GHD_SETTINGS_KEY = "informe_d_gh_settings";
const GHD_CARPETA_DEFECTO = "informe-d";
const GHD_ARCHIVO_CONFIG = "config_difcambio.json";
const GHD_ARCHIVO_ESTADO = "estado_d.json";

function loadGhdSettings() {
  try { return JSON.parse(localStorage.getItem(GHD_SETTINGS_KEY) || "{}"); }
  catch (e) { return {}; }
}
function saveGhdSettings(s) { localStorage.setItem(GHD_SETTINGS_KEY, JSON.stringify(s)); }
function hasGhdSettings() { const s = loadGhdSettings(); return !!(s.token && s.repo); }

function ghdRuta(nombre) {
  const c = (loadGhdSettings().carpeta || GHD_CARPETA_DEFECTO).replace(/^\/+|\/+$/g, "");
  return c ? `${c}/${nombre}` : nombre;
}

function ghdUtf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let b = "";
  bytes.forEach(x => { b += String.fromCharCode(x); });
  return btoa(b);
}
function ghdBase64ToUtf8(b64) {
  const bin = atob(String(b64).replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

function ghdCabeceras() {
  const s = loadGhdSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  return { "Authorization": `token ${s.token}`, "Accept": "application/vnd.github+json" };
}

async function ghdLeer(nombre) {
  const s = loadGhdSettings();
  const ruta = ghdRuta(nombre);
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
              `?ref=${encodeURIComponent(s.rama || "main")}&_=${Date.now()}`;
  const res = await fetch(url, { headers: ghdCabeceras(), cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`No pude leer ${ruta} de GitHub (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return { contenidoBase64: data.content, sha: data.sha };
}

function ghdPut(ruta, contenidoBase64, sha, mensaje, rama) {
  const s = loadGhdSettings();
  const cuerpo = { message: mensaje, content: contenidoBase64, branch: rama };
  if (sha) cuerpo.sha = sha;
  return fetch(`https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}`, {
    method: "PUT",
    headers: { ...ghdCabeceras(), "Content-Type": "application/json" },
    body: JSON.stringify(cuerpo),
    cache: "no-store",
  });
}

async function ghdEscribir(nombre, contenidoBase64, sha, mensaje) {
  const s = loadGhdSettings();
  const ruta = ghdRuta(nombre);
  const rama = s.rama || "main";
  let res = await ghdPut(ruta, contenidoBase64, sha, mensaje, rama);
  if (res.status === 409) {
    const actual = await ghdLeer(nombre);
    res = await ghdPut(ruta, contenidoBase64, actual ? actual.sha : undefined, mensaje, rama);
    if (res.status === 409) {
      throw new Error(`${ruta} cambió en GitHub mientras se guardaba. Actualizá la página ` +
                      `(Ctrl+Shift+R) y volvé a intentarlo. NO se guardó nada.`);
    }
  }
  if (!res.ok) throw new Error(`No pude guardar ${ruta} en GitHub (${res.status}): ${await res.text()}`);
  return await res.json();
}

// La primera vez todavía no existe: se arranca con la config de fábrica de la spec.
async function ghdLeerConfig() {
  const r = await ghdLeer(GHD_ARCHIVO_CONFIG);
  if (!r) return null;
  try { return JSON.parse(ghdBase64ToUtf8(r.contenidoBase64)); } catch (e) { return null; }
}

async function ghdGuardarConfig(cfg, mensaje) {
  const sha = (await ghdLeer(GHD_ARCHIVO_CONFIG))?.sha;
  await ghdEscribir(GHD_ARCHIVO_CONFIG, ghdUtf8ToBase64(JSON.stringify(cfg, null, 1)), sha,
                    mensaje || "Diferencia de cambio: actualiza la configuración");
}

async function ghdLeerEstado() {
  const r = await ghdLeer(GHD_ARCHIVO_ESTADO);
  if (!r) return { historial: [] };
  try { return JSON.parse(ghdBase64ToUtf8(r.contenidoBase64)); } catch (e) { return { historial: [] }; }
}

async function ghdGuardarEstado(estado, mensaje) {
  const sha = (await ghdLeer(GHD_ARCHIVO_ESTADO))?.sha;
  await ghdEscribir(GHD_ARCHIVO_ESTADO, ghdUtf8ToBase64(JSON.stringify(estado, null, 1)), sha,
                    mensaje || "Diferencia de cambio: nuevo asiento generado");
}

if (typeof module !== "undefined") {
  module.exports = {
    loadGhdSettings, saveGhdSettings, hasGhdSettings,
    ghdLeer, ghdEscribir, ghdLeerConfig, ghdGuardarConfig, ghdLeerEstado, ghdGuardarEstado,
  };
}
