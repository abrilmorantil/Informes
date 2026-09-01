// Gestión de categorías del balance en pesos: qué cuentas de Onvio integran cada categoría
// de RESULTADOS (las "cuenta madre" que arma Anexo II — Alquileres, Honorarios, etc.), y
// operaciones para sacar una cuenta, agregar una nueva o corregirle el código/nombre.
//
// La categoría vive en SALDOS como una fila madre con un subtotal (`SUM(F..:F..)` o
// `+F..+F..`) que agrega las filas de sus cuentas (madresResultados, en motor_balances.js).
// Estas funciones tocan esas fórmulas de verdad: sacar una cuenta reacomoda el subtotal de
// la madre y todo lo que dependía de esa fila, igual que si se borrara a mano en Excel.

// Todas las filas de SALDOS que integran el bloque de una madre, en orden.
function gcFilasDelBloque(bloque) {
  if (bloque.tipo === "rango") {
    return Array.from({ length: bloque.hasta - bloque.desde + 1 }, (_, i) => bloque.desde + i);
  }
  return bloque.filas.slice();
}

// Categorías de RESULTADOS con sus cuentas miembro, cruzando madresResultados (la
// geometría del archivo) con mapeo.cuentas (qué código tiene cada fila).
function categoriasPesos(wb, mapeo) {
  const madres = madresResultados(wb, "pesos");
  const porFila = new Map();
  for (const [codigo, f] of Object.entries(mapeo.cuentas)) porFila.set(f.fila, { codigo, ...f });

  const categorias = madres.map(m => {
    const filas = gcFilasDelBloque(m.bloque);
    const miembros = filas
      .map(fila => {
        const c = porFila.get(fila);
        if (!c) return null;
        return {
          codigo: c.codigo, nombre: c.nombre, fila, col: c.col,
          esFilaCompartida: fila === m.fila,
        };
      })
      .filter(Boolean);
    return { filaMadre: m.fila, codigo: m.codigo, nombre: m.nombre, bloque: m.bloque, miembros };
  });
  categorias.sort((a, b) => a.filaMadre - b.filaMadre);

  // Cuentas de RESULTADOS que no integran ninguna categoría: la fila propia de cada madre
  // (algunas tienen su propio VLOOKUP, pero ese valor no lo suma ningún subtotal — no es
  // un bug de esta app, así viene el archivo) y las líneas de resultado que van sueltas
  // (recupero de siniestro, venta de bien de uso, etc.), antes del bloque de gastos.
  const filasEnCategoria = new Set(categorias.flatMap(c => c.miembros.map(m => m.fila)));
  const sueltas = Object.entries(mapeo.cuentas)
    .filter(([, f]) => f.capitulo === "RESULTADOS" && !filasEnCategoria.has(f.fila))
    .map(([codigo, f]) => ({ codigo, nombre: f.nombre, fila: f.fila, col: f.col }))
    .sort((a, b) => a.fila - b.fila);

  return { categorias, sueltas };
}

// Reescribe el subtotal de la madre para que ya no cuente `filaSacada`, ANTES de borrar la
// fila. Hace falta para los dos casos que `borrarFilaEn` no arregla solo:
//   - bloque "lista" (+F63+F64+..): el término de la fila que se va no desaparece solo, y
//     si no se saca primero `borrarFilaEn` corta la operación (ve la propia madre
//     referenciando la fila que se quiere borrar).
//   - bloque "rango" cuando se saca la ÚLTIMA fila (`hasta`): los rangos se achican solos
//     al insertar/borrar, pero solo cuando la fila borrada queda ADENTRO del rango. Si es
//     el límite de arriba, el número de esa punta no se toca (nada se corrió por encima de
//     ella) y el rango se queda una fila de más, ahora apuntando a lo que sigue después de
//     la categoría — sumaría una cuenta ajena. Se lo achica a mano antes de borrar.
function gcAjustarSubtotalAntesDeBorrar(ws, madre, filaSacada) {
  const celda = ws.getCell(madre.fila, 7); // columna G, saldosColValor de pesos
  if (madre.bloque.tipo === "lista") {
    const nueva = celda.value.formula.replace(new RegExp(`\\+F${filaSacada}\\b`), "");
    if (nueva === celda.value.formula) {
      throw new Error(`No encontré "+F${filaSacada}" en el subtotal de "${madre.nombre}" (${celda.value.formula}). NO se tocó el archivo.`);
    }
    celda.value = { formula: nueva };
    return;
  }
  // rango
  if (filaSacada === madre.bloque.hasta) {
    const nuevoHasta = madre.bloque.hasta - 1;
    celda.value = { formula: `SUM(F${madre.bloque.desde}:F${nuevoHasta})` };
  }
  // si está en el medio o es el primero, el rango se achica solo al borrar la fila.
}

// Saca una cuenta de su categoría: reacomoda el subtotal de la madre y todas las fórmulas
// del archivo que dependían de esa fila (como al borrar cualquier fila de SALDOS).
function quitarCuentaDeCategoria(wb, categoria, filaCuenta, log = () => {}) {
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El archivo no tiene la hoja 'SALDOS'.");

  const miembro = categoria.miembros.find(m => m.fila === filaCuenta);
  if (!miembro) {
    throw new Error(`La fila ${filaCuenta} no integra la categoría "${categoria.nombre}". NO se tocó el archivo.`);
  }
  if (miembro.esFilaCompartida) {
    throw new Error(
      `"${miembro.codigo} - ${miembro.nombre}" comparte la fila con la categoría "${categoria.nombre}" ` +
      `en el Excel (la madre y esta cuenta viven en la misma fila). Sacarla requiere mover antes la ` +
      `categoría a otra fila del bloque, y esa operación todavía no está soportada acá — se puede ` +
      `hacer a mano en Excel. NO se tocó el archivo.`
    );
  }
  if (categoria.miembros.length <= 1) {
    throw new Error(
      `"${categoria.nombre}" se quedaría sin ninguna cuenta. Si la categoría ya no hace falta, se saca ` +
      `a mano en Excel junto con su fila de subtotal. NO se tocó el archivo.`
    );
  }

  const madresActuales = madresResultados(wb, "pesos");
  const madre = madresActuales.find(m => m.fila === categoria.filaMadre);
  if (!madre) throw new Error(`No encontré la categoría "${categoria.nombre}" en la fila ${categoria.filaMadre}. NO se tocó el archivo.`);

  gcAjustarSubtotalAntesDeBorrar(ws, madre, filaCuenta);
  const modificadas = borrarFilaEn(wb, "SALDOS", filaCuenta);
  log(`  "${miembro.codigo} - ${miembro.nombre}" sacada de "${categoria.nombre}" (fila ${filaCuenta}, ${modificadas} referencias reacomodadas).`);
  return { modificadas };
}

// Corrige el código o el nombre de una cuenta. NO es sólo texto: ese texto ES la clave con la
// que SALDOS busca su importe en Hoja1 (`VLOOKUP(E200, Hoja1!$A$2:$E$399, 5, FALSE)`).
//
// Cambiarlo sólo en SALDOS apaga la cuenta: el VLOOKUP deja de encontrarla, el IFERROR lo
// vuelve cero y no avisa nada. Probado sobre el maestro real — renombrando "111010001 - Caja"
// a cualquier otra cosa, Caja pasa a valer 0. Y no se arregla solo el mes siguiente:
// `actualizarHoja1` empareja por CÓDIGO y a propósito NO pisa el texto de Hoja1, así que la
// clave vieja se queda ahí para siempre.
//
// Por eso se renombra en los DOS lados y después se comprueba que la cuenta siga siendo
// encontrable. Si no se puede, se deshace todo y se explica: es preferible no dejar renombrar
// a dejar una cuenta en cero sin que se note.
//
// Ojo con el caso de una cuenta que todavía no está en Hoja1 (no vino en ningún export): ahí
// no hay nada que renombrar, y cuando aparezca, Hoja1 la va a insertar con el nombre que
// mande ONVIO. Si el nombre puesto acá no es ese, la cuenta va a leer cero. Se avisa.
function gcNormClave(s) {
  return String(s === null || s === undefined ? "" : s).trim().replace(/\s+/g, " ").toUpperCase();
}

// La fila de Hoja1 que alimenta una clave de SALDOS. Se busca por TEXTO, que es como compara
// el VLOOKUP; si no aparece, por código, que es como empareja `actualizarHoja1`.
function gcFilaEnHoja1(hoja1, colClave, texto) {
  const objetivo = gcNormClave(texto);
  const codigo = (/^\s*(\d{5,})\s*-/.exec(String(texto)) || [])[1];
  let porCodigo = null;
  for (let r = 1; r <= hoja1.rowCount; r++) {
    const v = hoja1.getCell(r, colClave).value;
    if (v && typeof v === "object" && typeof v.formula === "string") continue;   // celda de control
    const t = String(v === null || v === undefined ? "" : v).trim();
    if (!t) continue;
    if (gcNormClave(t) === objetivo) return r;
    if (porCodigo === null && codigo) {
      const m = /^\s*(\d{5,})\s*-/.exec(t);
      if (m && m[1] === codigo) porCodigo = r;
    }
  }
  return porCodigo;
}

function editarCuenta(wb, fila, col, nuevoCodigo, nuevoNombre, log = () => {}) {
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El archivo no tiene la hoja 'SALDOS'.");
  if (!/^\d{6,}$/.test(String(nuevoCodigo))) {
    throw new Error(`"${nuevoCodigo}" no parece un código de cuenta válido. NO se tocó el archivo.`);
  }
  const hoja1 = wb.getWorksheet("Hoja1");
  if (!hoja1) throw new Error("El archivo no tiene la hoja 'Hoja1'. NO se tocó nada.");
  const colClave = 1;                       // Hoja1: la clave va en la columna A

  const celda = ws.getCell(fila, col);
  const antes = String(celda.value === null || celda.value === undefined ? "" : celda.value);
  const nuevo = `${nuevoCodigo} - ${nuevoNombre}`;
  if (gcNormClave(antes) === gcNormClave(nuevo)) return { sinCambios: true };

  const filaH1 = gcFilaEnHoja1(hoja1, colClave, antes);
  const antesH1 = filaH1 ? String(hoja1.getCell(filaH1, colClave).value || "") : null;

  celda.value = nuevo;
  if (filaH1) hoja1.getCell(filaH1, colClave).value = nuevo;

  // El control: la clave nueva tiene que poder encontrarse en Hoja1, que es lo único que hace
  // que la cuenta siga levantando su importe.
  if (filaH1 && gcFilaEnHoja1(hoja1, colClave, nuevo) === null) {
    celda.value = antes;
    hoja1.getCell(filaH1, colClave).value = antesH1;
    throw new Error(
      `Con "${nuevo}" la cuenta dejaría de encontrar su importe en Hoja1 y quedaría en cero. ` +
      `Se deshizo el cambio, NO se tocó el archivo.`);
  }

  log(`  Fila ${fila}: "${antes}" → "${nuevo}".` +
      (filaH1 ? ` Hoja1 fila ${filaH1} renombrada igual, para que siga encontrando su importe.`
              : ` OJO: esta cuenta todavía no está en Hoja1 (no vino en ningún export). Cuando ` +
                `aparezca, Hoja1 la va a escribir con el nombre que mande Onvio: si no es ` +
                `exactamente "${nuevoNombre}", va a leer cero.`));
  return { fila, filaHoja1: filaH1, antes, nuevo };
}

// La cuenta está escrita en la columna equivocada.
//
// Cada fila de SALDOS levanta su importe con un VLOOKUP cuya clave es UNA columna concreta
// (`VLOOKUP(C40, Hoja1!...)`). Si el texto de la cuenta quedó en la columna de al lado, la
// fórmula lee una celda vacía y la fila no puede traer un importe nunca. En el bloque del
// activo las filas sanas tienen el mismo texto en C y en D; las rotas tienen la C vacía.
//
// Se escribe SOLO la columna clave y no se toca la otra: en la fila 43 la D trae el código
// viejo del cliente (`12301000` contra el `123010000` de Onvio), y ése es justamente el dato
// que hay que conservar — es la equivalencia entre los dos planes de cuentas.
function cfgColumnaClaveDe(ws, fila, colValor) {
  const RE = /VLOOKUP\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*[,;]\s*Hoja1!/i;
  for (const c of [colValor, colValor - 1]) {
    const f = ws.getCell(fila, c).formula;
    const m = f && RE.exec(String(f));
    if (m && +m[2] === fila) {
      return m[1].split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
    }
  }
  return null;
}

function corregirColumnaDeCuenta(wb, fila, codigo, nombre, log = () => {}) {
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El archivo no tiene la hoja 'SALDOS'.");
  if (!/^\d{5,}$/.test(String(codigo))) {
    throw new Error(`"${codigo}" no parece un código de cuenta válido. NO se tocó el archivo.`);
  }
  const colClave = cfgColumnaClaveDe(ws, fila, 7);
  if (colClave === null) {
    throw new Error(`La fila ${fila} de SALDOS no levanta su importe con un VLOOKUP contra Hoja1, ` +
                    `así que no hay ninguna columna clave que corregir. NO se tocó el archivo.`);
  }
  const yaHabia = String(ws.getCell(fila, colClave).value || "").trim();
  if (yaHabia) {
    throw new Error(`La fila ${fila} ya tiene "${yaHabia}" en su columna clave. ` +
                    `Si hay que cambiarla, se hace desde "corregir". NO se tocó el archivo.`);
  }

  const texto = `${codigo} - ${nombre}`;
  ws.getCell(fila, colClave).value = texto;
  // el formato de una fila vecina sana, para que no quede desalineada
  for (const vecina of [fila - 1, fila + 1, fila - 2]) {
    const v = ws.getCell(vecina, colClave);
    if (v && v.value !== null && v.value !== undefined && String(v.value).trim()) {
      ws.getCell(fila, colClave).style = v.style;
      break;
    }
  }

  // ¿va a encontrar su importe? Todavía puede no estar en Hoja1 si nunca vino en un export.
  const hoja1 = wb.getWorksheet("Hoja1");
  const enHoja1 = hoja1 ? gcFilaEnHoja1(hoja1, 1, texto) : null;
  const letra = String.fromCharCode(64 + colClave);
  log(`  Fila ${fila}: la cuenta pasa a la columna ${letra}, que es la que lee la fórmula — ` +
      `"${texto}".` +
      (enHoja1 ? ` Ya la encuentra en Hoja1 (fila ${enHoja1}).`
               : ` Todavía no está en Hoja1 porque no vino en ningún export; va a leer su ` +
                 `importe el mes que Onvio la mande.`));
  return { fila, columna: letra, texto, enHoja1 };
}

// Agrega una cuenta nueva a una categoría: es la misma operación que usa el alta de
// cuentas nuevas durante la corrida mensual (motor_balances.js).
function agregarCuentaACategoria(wb, mapeo, cuenta, categoria, log = () => {}) {
  return insertarHijaEnMadre(wb, mapeo, cuenta, categoria.filaMadre, "pesos", log);
}

if (typeof module !== "undefined") {
  module.exports = {
    gcFilasDelBloque, categoriasPesos, quitarCuentaDeCategoria, editarCuenta, agregarCuentaACategoria,
    gcFilaEnHoja1, gcNormClave, corregirColumnaDeCuenta, cfgColumnaClaveDe,
  };
}
