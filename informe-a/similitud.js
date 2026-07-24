// Equivalente en JS de difflib.SequenceMatcher.ratio() y difflib.get_close_matches
// de Python, que el motor original usa para emparejar el nombre de un centro de
// costo del export de Onvio contra el nombre que tiene en el balance.
// Se replica el algoritmo de bloques coincidentes de difflib (no una distancia de
// edición cualquiera) para que el resultado sea el mismo que daba el Python.

function construirB2J(b) {
  const b2j = new Map();
  for (let j = 0; j < b.length; j++) {
    const ch = b[j];
    if (!b2j.has(ch)) b2j.set(ch, []);
    b2j.get(ch).push(j);
  }
  return b2j;
}

function findLongestMatch(a, b, b2j, alo, ahi, blo, bhi) {
  let besti = alo, bestj = blo, bestsize = 0;
  let j2len = new Map();

  for (let i = alo; i < ahi; i++) {
    const newj2len = new Map();
    const indices = b2j.get(a[i]) || [];
    for (const j of indices) {
      if (j < blo) continue;
      if (j >= bhi) break;
      const k = (j2len.get(j - 1) || 0) + 1;
      newj2len.set(j, k);
      if (k > bestsize) {
        besti = i - k + 1;
        bestj = j - k + 1;
        bestsize = k;
      }
    }
    j2len = newj2len;
  }

  return [besti, bestj, bestsize];
}

function totalCoincidencias(a, b) {
  const b2j = construirB2J(b);
  let total = 0;
  const cola = [[0, a.length, 0, b.length]];
  while (cola.length) {
    const [alo, ahi, blo, bhi] = cola.pop();
    const [i, j, k] = findLongestMatch(a, b, b2j, alo, ahi, blo, bhi);
    if (k === 0) continue;
    total += k;
    if (alo < i && blo < j) cola.push([alo, i, blo, j]);
    if (i + k < ahi && j + k < bhi) cola.push([i + k, ahi, j + k, bhi]);
  }
  return total;
}

function ratio(a, b) {
  const t = a.length + b.length;
  if (t === 0) return 1.0;
  return (2.0 * totalCoincidencias(a, b)) / t;
}

// Equivalente a difflib.get_close_matches(palabra, posibilidades, n=1, cutoff).
// difflib compara con la palabra buscada como segunda secuencia, así que se
// respeta ese orden porque los bloques coincidentes no son simétricos.
function getCloseMatches(palabra, posibilidades, n = 1, cutoff = 0.6) {
  const puntuadas = [];
  for (const x of posibilidades) {
    const r = ratio(x, palabra);
    if (r >= cutoff) puntuadas.push([r, x]);
  }
  puntuadas.sort((p, q) => q[0] - p[0]);
  return puntuadas.slice(0, n).map(p => p[1]);
}

if (typeof module !== "undefined") {
  module.exports = { ratio, getCloseMatches };
}
