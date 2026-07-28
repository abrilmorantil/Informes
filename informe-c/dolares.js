// El balance en dólares NO tiene el mismo plan de cuentas que el export.
//
// Medido sobre junio 2026: de sus 203 cuentas, 142 usan códigos viejos de 8 dígitos, y
// esos códigos fueron REASIGNADOS, no reformateados. La regla que parecía obvia
// (`11401003` → `114010003`) es falsa: de 40 códigos que "resuelve", 14 caen en una
// cuenta distinta — `42102000 Honorarios profesionales` daría `421020000 Servicios
// Públicos`. Emparejar por parecido de nombre es todavía peor: propone meter
// `Dep. Ac. Muebles y Útiles` dentro de `Muebles y Útiles`.
//
// Así que se resuelve con lo que SÍ es confiable, en este orden, y lo que sobra se le
// pregunta a la usuaria una sola vez y queda guardado:
//   1. código idéntico
//   2. proveedores: el prefijo configurado en el Informe B (21101) va todo a la línea
//      única "Proveedores" (en dólares se agrupan; en pesos se detallan)
//   3. cuenta madre declarada en el Informe B: el balance en dólares está a nivel de
//      madre, así que las hijas del export se SUMAN en la fila de su madre
//   4. nombre exacto (normalizado)
//   5. equivalencias confirmadas a mano en corridas anteriores
// Nada de emparejar por parecido: una sola equivocación mete plata en la línea que no es.

function normNombre(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// --------------------------------------------------- propuesta por código + nombre
//
// El plan viejo tiene un dígito menos en el grupo del medio: `114010005` (export) es
// `11401005` (maestro). Esa regla SOLA no sirve —probada sobre las 142 cuentas viejas,
// 14 de los 40 códigos que "resuelve" caen en otra cuenta—, y el parecido de nombre
// solo tampoco. Pero exigiendo LAS DOS cosas a la vez se vuelve confiable: el código
// candidato tiene que existir en el maestro Y el nombre tiene que corroborarlo.
//
// Lo importante es que los casos peligrosos quedan afuera por construcción: para
// "124020001 Dep. Ac. Muebles y Útiles" el código candidato no existe, así que no hay
// propuesta —en vez de proponer meterla dentro del bien de uso, como hacía el parecido
// de nombre—. Control sobre las cuentas que ya tienen destino seguro: 0 choques.
//
// Aun así la propuesta NUNCA se aplica sola: se deja pre-elegida en la pantalla para
// que la usuaria la confirme de un vistazo.

const PALABRAS_VACIAS = ["DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "A", "EN", "CTA", "PART", "SA", "SRL"];

function tokensNombre(s) {
  return normNombre(s).split(" ").filter(t => t.length > 1 && PALABRAS_VACIAS.indexOf(t) < 0);
}

// "SALDO" casa con "SALDOS" y "GCIA" con "GCIA": se compara por prefijo, porque los
// dos planes escriben la misma cuenta con abreviaturas distintas.
function casanTokens(a, b) {
  if (a === b) return true;
  const corto = a.length <= b.length ? a : b;
  const largo = a.length <= b.length ? b : a;
  return corto.length >= 3 && largo.indexOf(corto) === 0;
}

function similitudNombre(n1, n2) {
  const A = tokensNombre(n1), B = tokensNombre(n2);
  if (!A.length || !B.length) return 0;
  const usados = [];
  let comun = 0;
  for (const a of A) {
    for (let j = 0; j < B.length; j++) {
      if (usados.indexOf(j) < 0 && casanTokens(a, B[j])) { comun++; usados.push(j); break; }
    }
  }
  return comun / Math.max(A.length, B.length);
}

// Los códigos del plan viejo que podrían corresponder a este: el mismo, y los que
// resultan de sacarle un cero del grupo del medio.
function codigosCandidatos(codigo) {
  const c = String(codigo);
  const out = [c];
  for (let i = 3; i <= 6 && i < c.length; i++) {
    if (c.charAt(i) === "0") {
      const alt = c.slice(0, i) + c.slice(i + 1);
      if (out.indexOf(alt) < 0) out.push(alt);
    }
  }
  return out;
}

const SIMILITUD_MINIMA = 0.4;

function proponerDestino(cuenta, cuentasMaestro) {
  const porCodigo = new Map(cuentasMaestro.map(f => [f.codigo, f]));
  let mejor = null;
  for (const cod of codigosCandidatos(cuenta.codigo)) {
    if (cod === cuenta.codigo) continue;          // ese ya lo probó la vía directa
    const f = porCodigo.get(cod);
    if (!f) continue;
    const s = similitudNombre(cuenta.nombre, f.nombre);
    if (s >= SIMILITUD_MINIMA && (!mejor || s > mejor.similitud)) {
      mejor = { fila: f.fila, codigo: f.codigo, nombre: f.nombre, clave: f.clave, similitud: s };
    }
  }
  return mejor;
}

// Las cuentas del maestro, tal como están escritas en su hoja SALDOS.
function cuentasDelMaestro(wb, moneda) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El maestro no tiene la hoja 'SALDOS'.");
  const filas = [];
  const vistas = new Set();
  for (let r = 1; r <= ws.rowCount; r++) {
    for (const c of p.saldosColsCuenta) {
      const t = textoPlano(ws.getCell(r, c).value).trim();
      const m = RE_CUENTA_TXT.exec(t);
      if (!m) continue;
      const clave = t.replace(/\s+/g, " ");
      if (!vistas.has(clave)) { vistas.add(clave); filas.push({ fila: r, col: c, codigo: m[1], nombre: m[2], clave }); }
      break;
    }
  }
  return filas;
}

// A qué fila del maestro va cada cuenta del export. Devuelve también lo que no se pudo
// resolver, para preguntarlo.
function resolverDestinosDolares({ cuentasExport, cuentasMaestro, clasificacion, equivalencias = {} }) {
  const porCodigo = new Map(cuentasMaestro.map(f => [f.codigo, f]));
  const porNombre = new Map();
  for (const f of cuentasMaestro) {
    const k = normNombre(f.nombre);
    if (!porNombre.has(k)) porNombre.set(k, f);
  }
  // La línea de proveedores se busca por el CÓDIGO que el Informe B tiene configurado
  // (21101000), no por el nombre: buscar /proveedor/ agarraba primero
  // "11405004 - Anticipo a Proveedores", que es una cuenta del ACTIVO, y los 49
  // proveedores terminaban sumados ahí.
  const codProv = (clasificacion && clasificacion.codigoProveedores) || "21101000";
  const filaProveedores = porCodigo.get(codProv)
    || cuentasMaestro.find(f => normNombre(f.nombre) === "PROVEEDORES") || null;
  const madre = (clasificacion && clasificacion.madrePorHija) || new Map();

  const destinos = new Map();     // clave del maestro -> { fila, clave, aportes: [] }
  const pendientes = [];
  const ignoradas = [];

  const sumar = (f, cuenta, via) => {
    if (!destinos.has(f.clave)) destinos.set(f.clave, { fila: f.fila, clave: f.clave, codigo: f.codigo, nombre: f.nombre, aportes: [] });
    destinos.get(f.clave).aportes.push({ codigo: cuenta.codigo, nombre: cuenta.nombre, saldo: cuenta.saldo_usd, via });
  };

  for (const c of cuentasExport) {
    // Las decisiones ya tomadas mandan sobre todo lo demás, incluido el "no cargar".
    if (Object.prototype.hasOwnProperty.call(equivalencias, c.codigo)) {
      const destino = equivalencias[c.codigo];
      if (destino === null) { ignoradas.push(c); continue; }
      const f = porCodigo.get(String(destino)) || cuentasMaestro.find(x => x.clave === destino);
      if (f) { sumar(f, c, "confirmada"); continue; }
    }
    if (c.saldo_usd === 0) continue;          // sin saldo en dólares no toca este balance

    const directa = porCodigo.get(c.codigo);
    if (directa) { sumar(directa, c, "código"); continue; }

    if (filaProveedores && esProveedor(clasificacion, c.codigo)) { sumar(filaProveedores, c, "proveedores"); continue; }

    const m = madre.get(String(c.codigo));
    const filaMadre = m ? porCodigo.get(m.code) : null;
    if (filaMadre) { sumar(filaMadre, c, "cuenta madre"); continue; }

    const porNom = porNombre.get(normNombre(c.nombre));
    if (porNom) { sumar(porNom, c, "nombre"); continue; }

    pendientes.push(c);
  }

  return { destinos, pendientes, ignoradas, filaProveedores };
}

// Escribe en Hoja1 el importe que le toca a cada cuenta del maestro. A diferencia de
// pesos —donde el export y el archivo hablan el mismo plan de cuentas y se copia uno a
// uno— acá se escribe el TOTAL AGREGADO por cuenta del maestro.
//
// Igual que en pesos, Hoja1 no se reordena: se actualiza en el lugar y lo que no
// recibe nada queda en cero.
function volcarHoja1Dolares(wb, destinos, moneda, log = () => {}) {
  const p = PARAMS[moneda];
  const ws = wb.getWorksheet("Hoja1");
  if (!ws) throw new Error("El maestro no tiene la hoja 'Hoja1'.");

  const filaDeClave = new Map();
  for (let r = 1; r <= ws.rowCount; r++) {
    const v = ws.getCell(r, p.hoja1ColClave).value;
    if (v && typeof v === "object" && typeof v.formula === "string") continue;
    const t = String(v === null || v === undefined ? "" : v).trim().replace(/\s+/g, " ");
    if (!RE_CUENTA_TXT.test(t)) continue;
    if (!filaDeClave.has(t)) filaDeClave.set(t, r);
    ws.getCell(r, p.hoja1ColValor).value = 0;      // cada corrida reemplaza
  }

  let escritas = 0;
  const sinFila = [];
  for (const d of destinos.values()) {
    const total = d.aportes.reduce((a, x) => a + x.saldo, 0);
    const fila = filaDeClave.get(d.clave);
    if (fila === undefined) { sinFila.push(d); continue; }
    ws.getCell(fila, p.hoja1ColValor).value = total;
    escritas++;
  }

  log(`  Hoja1 (dólares): ${escritas} cuentas con importe; el resto en cero.`);
  for (const d of sinFila) {
    log(`  ⚠ "${d.clave}" está en SALDOS pero no tiene fila en Hoja1: su importe no se puede cargar.`);
  }
  return { escritas, sinFila };
}

// La corrida de dólares. No inserta cuentas nuevas: el balance en dólares es un
// resumen a nivel de cuenta madre, y toda cuenta del export termina sumando en alguna
// de sus líneas (o quedando explícitamente afuera).
function procesarDolares({ wb, cuentasExport, clasificacion, equivalencias = {}, log = () => {} }) {
  const cuentasMaestro = cuentasDelMaestro(wb, "dolares");
  const r = resolverDestinosDolares({ cuentasExport, cuentasMaestro, clasificacion, equivalencias });

  if (r.pendientes.length) {
    const e = new Error("Hay cuentas en dólares sin destino definido.");
    e.pendientesDolares = r.pendientes;
    e.cuentasMaestro = cuentasMaestro;
    throw e;
  }

  const porVia = {};
  for (const d of r.destinos.values()) for (const a of d.aportes) porVia[a.via] = (porVia[a.via] || 0) + 1;
  log(`  ${cuentasExport.filter(c => c.saldo_usd !== 0).length} cuentas del export con saldo en dólares.`);
  log(`  Ubicadas: ${Object.entries(porVia).map(([k, v]) => `${v} por ${k}`).join(", ")}.`);
  if (r.ignoradas.length) log(`  ${r.ignoradas.length} dejadas afuera a propósito (marcadas "no cargar").`);

  const volcado = volcarHoja1Dolares(wb, r.destinos, "dolares", log);

  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  const total = cuentasExport.reduce((a, c) => a + c.saldo_usd, 0);
  const cargado = [...r.destinos.values()].reduce((a, d) => a + d.aportes.reduce((b, x) => b + x.saldo, 0), 0);
  return {
    resumen: {
      moneda: "dolares",
      cuentasMaestro: cuentasMaestro.length,
      lineasConImporte: volcado.escritas,
      total, cargado,
      ignoradas: r.ignoradas.map(c => `${c.codigo} - ${c.nombre}`),
      sinFilaEnHoja1: volcado.sinFila.map(d => d.clave),
    },
  };
}

if (typeof module !== "undefined") {
  const m = require("./motor_balances.js");
  global.PARAMS = m.PARAMS;
  global.RE_CUENTA_TXT = /^\s*(\d{6,})\s*-\s*(.+?)\s*$/;
  global.textoPlano = m.textoPlano;
  global.esProveedor = require("./clasificacion.js").esProveedor;
  module.exports = {
    normNombre, cuentasDelMaestro, resolverDestinosDolares, volcarHoja1Dolares, procesarDolares,
    proponerDestino, similitudNombre, codigosCandidatos,
  };
}
