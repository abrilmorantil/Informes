// La clasificación de cuentas no se vuelve a pedir: ya está configurada en el
// mapping.json del Informe B (BALCOMPROBDOLARES), que vive en el mismo repositorio.
//
// Ahí está, por código de cuenta:
//   - `category`  ACTIVO / PASIVO / CAPITAL Y PATRIMONIO / RESULTADOS
//   - `type: "parent"` + `children`  la relación madre/hija de los gastos
//   - `type: "range"` + `prefix`     el agrupamiento de proveedores (prefijo 21101)
//   - `aliases`                      nombres alternativos de la misma cuenta
//
// El Informe C lo usa para no preguntar lo que ya está resuelto. Lo que NO sale de
// acá es dónde va cada cosa DENTRO del balance formal (qué fila de SALDOS, qué línea
// de la Nota 4): eso es la geometría del archivo de la usuaria y se sigue leyendo del
// archivo. Uno dice "qué es cada cuenta", el otro "dónde está".

const PREFIJO_PROVEEDORES_DEFECTO = "21101";
const CODIGO_PROVEEDORES_DEFECTO = "21101000";

function indexarClasificacion(mappingB) {
  const porCodigo = new Map();
  const madrePorHija = new Map();
  let prefijoProveedores = null;
  let codigoProveedores = null;

  for (const x of mappingB || []) {
    if (!x || !x.code) continue;
    porCodigo.set(String(x.code), x);
    if (x.type === "range" && x.prefix) {
      prefijoProveedores = String(x.prefix);
      codigoProveedores = String(x.code);   // la línea única donde se agrupan
    }
    for (const h of (x.children || [])) {
      if (h && h.code) madrePorHija.set(String(h.code), { code: String(x.code), description: x.description });
    }
  }

  return {
    porCodigo,
    madrePorHija,
    // Sin mapping se usa 21101, que es el prefijo real de proveedores. Ojo con
    // acortarlo a 2110: ahí caen también "Provisión de Gastos" (211030000) y
    // "Previsión IGMP" (211060000), que no son proveedores.
    prefijoProveedores: prefijoProveedores || PREFIJO_PROVEEDORES_DEFECTO,
    codigoProveedores: codigoProveedores || CODIGO_PROVEEDORES_DEFECTO,
    vacio: porCodigo.size === 0,
  };
}

function esProveedor(clasif, codigo) {
  const p = (clasif && clasif.prefijoProveedores) || PREFIJO_PROVEEDORES_DEFECTO;
  return String(codigo).startsWith(p);
}

// La cuenta madre declarada en el Informe B, traducida a la fila que ocupa en el
// SALDOS de este maestro. Devuelve null si el Informe B no la conoce o si esa madre
// no existe en el archivo (ahí sí hay que preguntar).
function madreEnArchivo(clasif, madresDelArchivo, codigo) {
  const m = clasif && clasif.madrePorHija.get(String(codigo));
  if (!m) return null;
  const enArchivo = madresDelArchivo.find(x => String(x.codigo) === m.code);
  return enArchivo ? { fila: enArchivo.fila, codigo: m.code, nombre: m.description } : null;
}

function categoriaDe(clasif, codigo) {
  const x = clasif && clasif.porCodigo.get(String(codigo));
  return x ? x.category : null;
}

if (typeof module !== "undefined") {
  module.exports = {
    indexarClasificacion, esProveedor, madreEnArchivo, categoriaDe,
    PREFIJO_PROVEEDORES_DEFECTO, CODIGO_PROVEEDORES_DEFECTO,
  };
}
