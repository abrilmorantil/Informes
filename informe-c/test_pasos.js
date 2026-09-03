// La guía por pasos: qué paso se ve, cuántos hay y adónde vuelve el botón de atrás.
//
// Se prueba la lógica sola, sin pantalla. Lo que importa no es que dibuje: es que NO se saltee
// un paso cuando la app habilita dos de una vez, y que el paso que no hace falta desaparezca de
// la numeración en vez de mostrarse vacío.
//
//   node informe-c/test_pasos.js
const path = require("path");
const P = require(path.join(__dirname, "pasos.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };

// Reproduce lo que hace la app: una tanda de `mostrar()` seguidos.
// Una tanda de mostrar() seguidos, y el cierre de la tanda: es lo que hace la pantalla en cada
// tick. La decisión de qué paso se ve se toma al cerrarla, no en cada aviso.
function tanda(est, pares) {
  for (const [id, v] of pares) est = P.pasosAlMostrar(est, id, v);
  return P.pasosAsentar(est);
}
const donde = (est) => {
  const u = P.pasosUbicacion(est);
  return `${est.actual} (paso ${u.indice + 1} de ${u.total})`;
};

console.log("=== 1) el mes normal: hay cuentas nuevas ===");
{
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true]]);
  check(e.actual === "cardExport" && P.pasosUbicacion(e).total === 1,
    `arranca en la carga del balance — ${donde(e)}`);

  // se procesa y aparecen cuentas nuevas
  e = tanda(e, [["cardNuevas", true]]);
  check(e.actual === "cardNuevas", `pasa a las cuentas nuevas — ${donde(e)}`);
  check(P.pasosUbicacion(e).nombreAnterior === "Carga del balance",
    "y el botón de volver dice a dónde vuelve");

  // se confirman: la app apaga esa tarjeta y habilita controles + revisión JUNTOS
  e = tanda(e, [["cardNuevas", false], ["cardResultado", true], ["cardCierre", true]]);
  check(e.actual === "cardResultado",
    `de dos que se habilitan a la vez, se queda en la primera — ${donde(e)}`);
  check(P.pasosUbicacion(e).total === 3,
    "y las cuentas nuevas ya no cuentan: quedaron atrás y resueltas");
}

console.log("\n=== 2) el mes sin cuentas nuevas: ese paso no existe ===");
{
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true]]);
  e = tanda(e, [["cardResultado", true], ["cardCierre", true]]);
  const u = P.pasosUbicacion(e);
  check(e.actual === "cardResultado" && u.total === 3 && u.indice === 1,
    `no se muestra vacío: desaparece — ${donde(e)}`);
  check(u.nombreAnterior === "Carga del balance",
    "y volver saltea el paso que no existió, no lleva a una pantalla en blanco");
}

console.log("\n=== 3) volver, y volver a avanzar ===");
{
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true], ["cardNuevas", true]]);
  e = P.pasosVolver(e);
  check(e.actual === "cardExport", "volver lleva al paso anterior");
  check(P.pasosUbicacion(e).total === 2, "sin perder el paso al que se puede volver a entrar");
  e = P.pasosVolver(e);
  check(e.actual === "cardExport", "y desde el primero no se puede volver más atrás");
  e = P.pasosIrA(e, "cardNuevas");
  check(e.actual === "cardNuevas", "se puede volver adelante a un paso ya habilitado");
  e = P.pasosIrA(e, "cardCierre");
  check(e.actual === "cardNuevas", "pero no saltar a uno que el circuito todavía no habilitó");
}

console.log("\n=== 4) volver a procesar borra lo que había ===");
{
  // Al apretar Procesar de nuevo, la app apaga las tres tarjetas del final.
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true], ["cardNuevas", true]]);
  e = tanda(e, [["cardNuevas", false], ["cardResultado", true], ["cardCierre", true]]);
  e = tanda(e, [["cardNuevas", false], ["cardResultado", false], ["cardCierre", false]]);
  check(e.actual === "cardExport" && P.pasosUbicacion(e).total === 1,
    `vuelve a la carga del balance, con un solo paso — ${donde(e)}`);
}

console.log("\n=== 5) la revisión de dólares, cuando aparece ===");
{
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true]]);
  e = tanda(e, [["cardRevisionUsd", true]]);
  check(e.actual === "cardRevisionUsd", `se muestra — ${donde(e)}`);
  e = tanda(e, [["cardRevisionUsd", false], ["cardResultado", true], ["cardCierre", true]]);
  check(e.actual === "cardResultado" && P.pasosUbicacion(e).total === 3,
    `y al resolverse deja de contar — ${donde(e)}`);
}

console.log("\n=== 6) lo que NO es un paso ===");
{
  let e = P.pasosNuevoEstado();
  e = tanda(e, [["cardExport", true]]);
  const antes = JSON.stringify(e);
  // los paneles, el historial y la carga del maestro no son parte del cierre del mes
  e = tanda(e, [["cardConfigCuentas", true], ["cardHistorial", true], ["cardAlta", true],
                ["cardGuardados", true], ["cardEerrAnterior", true]]);
  check(JSON.stringify(e) === antes,
    "los paneles, el historial y la carga del maestro no entran en la numeración");
}

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
