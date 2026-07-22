// Integración con GitHub: lee y escribe mapping.json directo en el repo
// usando la Contents API, para que clasificar una cuenta nueva quede
// guardado para siempre sin descargar/subir nada a mano.

const GH_SETTINGS_KEY = "balcomp_gh_settings";

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

// Codifica un string UTF-8 a base64 sin romper acentos/ñ (btoa solo, se
// come cualquier caracter fuera de latin1).
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

async function ghGetFile() {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repo).");
  const branch = s.branch || "main";
  const path = s.path || "mapping.json";
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURIComponent(path)}?ref=${encodeURIComponent(branch)}`;
  const res = await fetch(url, {
    headers: {
      "Authorization": `token ${s.token}`,
      "Accept": "application/vnd.github+json",
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`No pude leer ${path} de GitHub (${res.status}): ${body}`);
  }
  const data = await res.json();
  const content = base64ToUtf8(data.content);
  return { mapping: JSON.parse(content), sha: data.sha };
}

async function ghPutFile(mapping, sha, mensaje) {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repo).");
  const branch = s.branch || "main";
  const path = s.path || "mapping.json";
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURIComponent(path)}`;
  const body = {
    message: mensaje || "Actualiza mapping.json desde BALCOMPROBDOLARES",
    content: utf8ToBase64(JSON.stringify(mapping, null, 2)),
    sha,
    branch,
  };
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": `token ${s.token}`,
      "Accept": "application/vnd.github+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const errBody = await res.text();
    throw new Error(`No pude guardar ${path} en GitHub (${res.status}): ${errBody}`);
  }
  return await res.json();
}
