// Lee y guarda mapeo_maestro.json en el repositorio usando la Contents API de
// GitHub, para que las cuentas nuevas que se clasifiquen queden guardadas sin
// tener que descargar y subir el archivo a mano.
//
// Usa su propia clave de configuración: la del informe BALCOMPROBDOLARES apunta a
// otro archivo y no se debe pisar.

const GH_SETTINGS_KEY = "informe_a_gh_settings";
const RUTA_MAPEO_DEFECTO = "informe-a/mapeo_maestro.json";

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

// btoa se come cualquier carácter fuera de latin1, así que los acentos y la ñ del
// mapeo se codifican pasando por UTF-8 explícitamente.
function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach(b => { binary += String.fromCharCode(b); });
  return btoa(binary);
}

function base64ToUtf8(b64) {
  const binary = atob(b64.replace(/\n/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

async function ghGetMapeo() {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  const branch = s.branch || "main";
  const path = s.path || RUTA_MAPEO_DEFECTO;
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(path)}?ref=${encodeURIComponent(branch)}`;
  const res = await fetch(url, {
    headers: { "Authorization": `token ${s.token}`, "Accept": "application/vnd.github+json" },
  });
  if (!res.ok) {
    throw new Error(`No pude leer ${path} de GitHub (${res.status}): ${await res.text()}`);
  }
  const data = await res.json();
  return { mapeo: JSON.parse(base64ToUtf8(data.content)), sha: data.sha };
}

async function ghPutMapeo(mapeo, sha, mensaje) {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  const branch = s.branch || "main";
  const path = s.path || RUTA_MAPEO_DEFECTO;
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(path)}`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": `token ${s.token}`,
      "Accept": "application/vnd.github+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: mensaje || "Actualiza mapeo_maestro.json desde el Balance de Comprobación USD",
      content: utf8ToBase64(JSON.stringify(mapeo, null, 2)),
      sha,
      branch,
    }),
  });
  if (!res.ok) {
    throw new Error(`No pude guardar ${path} en GitHub (${res.status}): ${await res.text()}`);
  }
  return await res.json();
}
