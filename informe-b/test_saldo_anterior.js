// El saldo anterior de una fila sale de TODAS las cuentas que la alimentan, no sólo del
// código de la fila.
//
// El caso que lo destapó: se juntaron las dos "Retenciones SUSS a pagar" en una sola línea,
// haciendo a `21409000` madre de `212020002`. Los saldos guardados eran 0,00 y −30,97; el
// informe imprimía sólo el de la madre —0,00— y la columna "Saldo anterior" pasaba de cerrar
// en −0,02 a dar 30,95. La plata no se perdía en ningún lado: simplemente dejaba de sumarse.
//
//   node informe-b/test_saldo_anterior.js
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;
const { saldoAnteriorDeLinea, codigosQueAlimentan, buildBalance } = require(path.join(AQUI, "core.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const cerca = (a, b) => Math.abs(a - b) < 0.005;

const PREV = {
  "21409000": 0, "212020002": -30.97,
  "111010001": 1014.2,
  "211010001": 7, "211010002": 3, "21101000": 5,
};

console.log("=== 1) una cuenta madre suma la suya y las de sus hijas ===");
{
  const madre = { code: "21409000", type: "parent",
                  children: [{ code: "21409000" }, { code: "212020002" }] };
  const r = saldoAnteriorDeLinea(madre, PREV);
  check(cerca(r.total, -30.97), `suma 0,00 + (−30,97) = ${r.total}`);
  check(r.conocido, "y se da por conocido");
  // la madre figura tambien como hija suya: no se puede contar dos veces
  const dosVeces = { code: "111010001", type: "parent",
                     children: [{ code: "111010001" }, { code: "111010001" }] };
  check(cerca(saldoAnteriorDeLinea(dosVeces, PREV).total, 1014.2),
    "una cuenta repetida entre madre e hijas se cuenta UNA vez");
}

console.log("\n=== 2) el renglón de proveedores junta todo su prefijo ===");
{
  const rango = { code: "21101000", type: "range", prefix: "21101" };
  const cods = codigosQueAlimentan(rango, PREV);
  check(cods.has("211010001") && cods.has("211010002") && cods.has("21101000"),
    `toma las ${cods.size} cuentas que empiezan con 21101`);
  check(cerca(saldoAnteriorDeLinea(rango, PREV).total, 15), "y suma sus saldos: 5 + 7 + 3 = 15");
}

console.log("\n=== 3) cuándo sale en amarillo ===");
{
  check(!saldoAnteriorDeLinea({ code: "999999999", type: "simple" }, PREV).conocido,
    "una cuenta sin saldo guardado: amarilla");
  check(!saldoAnteriorDeLinea({ code: "21409000", sin_cuentas: true }, PREV).conocido,
    "una fila declarada sin cuentas: amarilla, no lee nada de Onvio");
  // si UNA de las cuentas tiene saldo, el numero sale: mejor el dato parcial que nada
  const media = { code: "21409000", type: "parent",
                  children: [{ code: "212020002" }, { code: "999999999" }] };
  const r = saldoAnteriorDeLinea(media, PREV);
  check(r.conocido && cerca(r.total, -30.97),
    "si al menos una de sus cuentas tiene saldo, el importe sale igual");
  check(saldoAnteriorDeLinea({ code: "21409000", type: "simple" }, PREV).conocido,
    "un saldo guardado en 0 cuenta como conocido, no como faltante");
}

console.log("\n=== 4) buildBalance lo arrastra hasta la línea ===");
{
  const mapping = [
    { code: "21409000", description: "Retenciones SUSS a pagar", category: "PASIVO",
      type: "parent", orden: 1,
      children: [{ code: "21409000", description: "SUSS" }, { code: "212020002", description: "SUSS 2" }] },
  ];
  const cuentas = { "212020002": { descripcion: "SUSS 2", debe: 31.24, haber: 30.82, saldo: 0.42, fila: 9 } };
  const { lineas } = buildBalance(cuentas, mapping, PREV);
  const l = lineas[0];
  check(cerca(l.saldo_anterior, -30.97), `la línea sale con saldo anterior ${l.saldo_anterior}`);
  check(l.saldo_anterior_conocido === true, "marcada como conocida, así que no va en amarillo");
  check(cerca(l.saldo_nuevo, -30.97 + 0.42), `y el saldo final es anterior + movimiento: ${l.saldo_nuevo}`);
}

console.log("\n=== 5) contra el informe real de agosto ===");
{
  const mapping = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));
  const estado = JSON.parse(fs.readFileSync(path.join(AQUI, "estado_b.json"), "utf8"));
  // la suma de los saldos anteriores de TODAS las filas tiene que cerrar como un balance
  let total = 0;
  const amarillas = [];
  for (const e of mapping) {
    const r = saldoAnteriorDeLinea(e, estado.saldos);
    total += r.total;
    // "Diferencia Resultados no Asignados" es un rotulo, no una cuenta: su "codigo" es texto
    // y no hay saldo que buscarle. Contarlo como amarillo seria un aviso que aparece siempre.
    if (!r.conocido && !e.sin_cuentas && /^\d+$/.test(String(e.code))) amarillas.push(e.code);
  }
  console.log(`       ${mapping.length} filas · suma ${total.toFixed(2)} · ${amarillas.length} sin saldo`);
  check(Math.abs(total) <= 1,
    `la columna "Saldo anterior" cierra: ${total.toFixed(2)} (antes daba 30,95 por la SUSS que no se sumaba)`);
  check(amarillas.length === 0,
    amarillas.length ? `salen en amarillo: ${amarillas.join(", ")}` : "ninguna cuenta sale en amarillo");
}

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
