// Memoria de saldos del BALCOMPROBDOLARES.
//
// Hasta ahora la columna "Saldo anterior" del archivo salía vacía y pintada de amarillo,
// para pegarla a mano del informe del mes pasado. Acá se guarda, mes a mes, el saldo
// final de cada cuenta, y la corrida siguiente lo trae solo — igual que el informe de
// centros de costo, que recuerda su archivo y su historial.
//
// Se guarda en `informe-b/estado_b.json` del mismo repositorio:
//   { ultimo_periodo, saldos: { "111010001": 1234.56, ... }, historial: [...] }
//
// Trae de nacimiento los arreglos que costó encontrar en los otros informes: lecturas
// sin caché (la Contents API responde con max-age=60 y el navegador reusa lo viejo) y
// reintento ante un 409 releyendo el sha real.

const ARCHIVO_ESTADO_B = "informe-b/estado_b.json";

function ghbCabeceras() {
  const s = loadGhSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repo).");
  return { "Authorization": `token ${s.token}`, "Accept": "application/vnd.github+json" };
}

async function ghbLeerArchivo(ruta) {
  const s = loadGhSettings();
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
              `?ref=${encodeURIComponent(s.branch || "main")}&_=${Date.now()}`;
  const res = await fetch(url, { headers: ghbCabeceras(), cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`No pude leer ${ruta} de GitHub (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return { contenido: base64ToUtf8(data.content), sha: data.sha };
}

function ghbPut(ruta, contenidoBase64, sha, mensaje) {
  const s = loadGhSettings();
  const cuerpo = { message: mensaje, content: contenidoBase64, branch: s.branch || "main" };
  if (sha) cuerpo.sha = sha;
  return fetch(`https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}`, {
    method: "PUT",
    headers: { ...ghbCabeceras(), "Content-Type": "application/json" },
    body: JSON.stringify(cuerpo),
    cache: "no-store",
  });
}

async function ghbEscribirArchivo(ruta, texto, mensaje) {
  const actual = await ghbLeerArchivo(ruta);
  let res = await ghbPut(ruta, utf8ToBase64(texto), actual ? actual.sha : undefined, mensaje);
  if (res.status === 409) {
    const reintento = await ghbLeerArchivo(ruta);
    res = await ghbPut(ruta, utf8ToBase64(texto), reintento ? reintento.sha : undefined, mensaje);
    if (res.status === 409) {
      throw new Error(`${ruta} cambió en GitHub mientras se guardaba. Actualizá la página ` +
                      `(Ctrl+Shift+R) y volvé a intentarlo. NO se guardó nada.`);
    }
  }
  if (!res.ok) throw new Error(`No pude guardar ${ruta} en GitHub (${res.status}): ${await res.text()}`);
  return await res.json();
}

const ESTADO_B_VACIO = { ultimo_periodo: null, saldos: {}, historial: [] };

async function leerEstadoB() {
  try {
    const r = await ghbLeerArchivo(ARCHIVO_ESTADO_B);
    if (!r) return { ...ESTADO_B_VACIO };
    const e = JSON.parse(r.contenido);
    return {
      ultimo_periodo: e.ultimo_periodo || null,
      saldos: e.saldos || {},
      historial: e.historial || [],
    };
  } catch (e) {
    return { ...ESTADO_B_VACIO, error: e.message };
  }
}

// El saldo con el que cierra cada cuenta este mes: es el que va a aparecer como
// "Saldo anterior" en la corrida siguiente.
function saldosDeCierre(lineas, saldosAnteriores) {
  const previos = saldosAnteriores || {};
  const out = {};
  for (const l of lineas || []) {
    const antes = typeof previos[l.code] === "number" ? previos[l.code] : 0;
    const mov = (Number(l.debe) || 0) - (Number(l.haber) || 0);
    out[l.code] = Math.round((antes + mov) * 100) / 100;
  }
  return out;
}

async function guardarEstadoB(estado, periodo, lineas) {
  const saldos = saldosDeCierre(lineas, estado.saldos);
  const total = Object.values(saldos).reduce((a, v) => a + v, 0);
  const historial = [...(estado.historial || []).filter(h => h.periodo !== periodo), {
    periodo,
    fecha: new Date().toISOString(),
    cuentas: Object.keys(saldos).length,
    total: Math.round(total * 100) / 100,
  }].sort((a, b) => String(a.periodo).localeCompare(String(b.periodo)));

  const nuevo = { ultimo_periodo: periodo, saldos, historial };
  await ghbEscribirArchivo(ARCHIVO_ESTADO_B, JSON.stringify(nuevo, null, 1),
                           `BALCOMPROBDOLARES: registra los saldos de ${periodo}`);
  return nuevo;
}

// Siembra la memoria con los saldos leídos de un informe ya terminado. A diferencia de
// `guardarEstadoB`, acá no se calcula nada: los saldos ya vienen cerrados del archivo.
// Sirve para arrancar sin tener que cargar la columna a mano una última vez.
async function sembrarEstadoB(estado, periodo, saldos, origen) {
  const limpios = {};
  for (const [codigo, v] of Object.entries(saldos || {})) {
    if (typeof v === "number" && isFinite(v)) limpios[codigo] = Math.round(v * 100) / 100;
  }
  if (!Object.keys(limpios).length) throw new Error("El archivo no traía ningún saldo para importar.");

  const total = Object.values(limpios).reduce((a, v) => a + v, 0);
  const historial = [...(estado.historial || []).filter(h => h.periodo !== periodo), {
    periodo,
    fecha: new Date().toISOString(),
    cuentas: Object.keys(limpios).length,
    total: Math.round(total * 100) / 100,
    importado: origen || true,
  }].sort((a, b) => String(a.periodo).localeCompare(String(b.periodo)));

  const nuevo = { ultimo_periodo: periodo, saldos: limpios, historial };
  await ghbEscribirArchivo(ARCHIVO_ESTADO_B, JSON.stringify(nuevo, null, 1),
                           `BALCOMPROBDOLARES: importa los saldos de ${periodo}`);
  return nuevo;
}

// Coteja los saldos que trae un informe importado contra el plan de cuentas.
//
// El informe no imprime TODAS las filas del plan: las marcadas `ocultar_si_cero` se saltean
// mientras estén en cero. Que esas falten no es un problema y su saldo es, justamente, cero:
// se completan. Antes quedaban sin saldo guardado, y una cuenta sin saldo guardado sale
// AMARILLA en la corrida siguiente, para cargarla a mano. O sea: una cuenta en cero pedía
// atención todos los meses sin tener nada que revisar.
//
// Cualquier OTRA cuenta que falte sí es un problema, y no se completa con cero: rellenar a
// ciegas convertiría un informe incompleto en ceros que parecen datos. Se devuelven aparte
// para avisar antes de guardar nada. Es lo que hacía falta el 02/09/2026, cuando se importó
// un informe de julio de 102 cuentas en lugar del de 197: se perdieron los saldos de 33
// cuentas que tenían plata y nada lo dijo.
function cotejarConElPlan(saldos, mapping) {
  const completados = { ...(saldos || {}) };
  const enCero = [], faltan = [];
  for (const x of (mapping || [])) {
    const cod = x && x.code !== undefined && x.code !== null ? String(x.code) : null;
    if (!cod || Object.prototype.hasOwnProperty.call(completados, cod)) continue;
    // Hay un renglón del plan cuyo "código" es texto ("Diferencia Resultados no Asignados"):
    // es un rótulo, no una cuenta, y no tiene saldo que buscar. Avisar por él sería un aviso
    // que aparece siempre y no pide nada, que es la forma más rápida de que se ignoren todos.
    if (!/^\d+$/.test(cod)) continue;
    if (x.ocultar_si_cero) {
      completados[cod] = 0;
      enCero.push({ code: cod, description: x.description || "" });
    } else {
      faltan.push({ code: cod, description: x.description || "" });
    }
  }
  return { completados, enCero, faltan };
}

if (typeof module !== "undefined") {
  module.exports = { ARCHIVO_ESTADO_B, ESTADO_B_VACIO, saldosDeCierre, leerEstadoB,
                     guardarEstadoB, sembrarEstadoB, cotejarConElPlan };
}
