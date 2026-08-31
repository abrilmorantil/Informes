// "Sin cuentas asignadas" tiene que desconectar de verdad, y la plata que quede suelta tiene
// que avisar, no desaparecer.
//
// El caso que lo motivó: el renglón que el cliente ve como "421170000 Gastos Legales" tiene
// ese mismo número que en Onvio es "Alojamiento Rel. Comunitarias Catamarca" — dos cosas
// distintas que coinciden de casualidad, porque los códigos fueron reasignados. Antes la
// marca era sólo un cartel: la fila igual se llevaba esa cuenta y encima sin trazabilidad.
const fs = require("fs");
const path = require("path");
const AQUI = __dirname;

const XLSX = require(path.join(AQUI, "..", "informe-a", "vendor", "xlsx.full.min.js"));
const core = require(path.join(AQUI, "core.js"));

let fallos = 0;
const check = (ok, m) => { console.log((ok ? "  OK  " : " FALLA") + " " + m); if (!ok) fallos++; };
const n2 = (x) => Math.round((x || 0) * 100) / 100;
const f2 = (x) => n2(x).toFixed(2);

const mapping = JSON.parse(fs.readFileSync(path.join(AQUI, "..", "mapping.json"), "utf8"));
const SIN = mapping.filter(e => e.sin_cuentas).map(e => e.code);

check(SIN.length > 0, `hay ${SIN.length} renglón(es) declarado(s) sin cuentas asignadas: ${SIN.join(", ")}`);
check(!SIN.includes("426320000"),
  '"426320000 Canon Minero Gst Campo" NO está entre ellos: en Onvio existe con ese número y ese nombre');

// Un export de mentira: sólo las cuentas que se quieren probar
const cuentas = {};
for (const code of SIN) {
  cuentas[code] = { code, descripcion: "lo que sea que diga Onvio", debe: 1000, haber: 0, fila: 9 };
}

const { lineas, totalDebe } = core.buildBalance(cuentas, mapping, {});
for (const code of SIN) {
  const l = lineas.find(x => x.code === code);
  check(l && l.debe === 0 && l.haber === 0,
    `"${l ? l.description : code}" queda en cero aunque Onvio tenga ${code} con 1.000`);
  check(l && (l.detalle || []).length === 0, `   y no dice que ese saldo salga de ningún lado`);
}
check(n2(totalDebe) === 0, `no entra nada al informe: total ${f2(totalDebe)}`);

const sueltas = core.findUnmapped(cuentas, mapping);
for (const code of SIN) {
  check(!!sueltas.find(x => x.code === code),
    `y la app avisa que ${code} quedó sin destino, en vez de tragárselo en silencio`);
}

// El control de siempre: con el export real de julio nada de esto cambia el informe.
const EXPORT = "C:/Users/amoran/Downloads/Balance de SyS por Cod. de Cta. (4).xls";
if (fs.existsSync(EXPORT)) {
  const libro = XLSX.read(fs.readFileSync(EXPORT), { type: "buffer", raw: true });
  const filas = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]], { header: 1, raw: true, defval: null });
  const r = core.parseSiseExport(filas);
  const estado = JSON.parse(fs.readFileSync(path.join(AQUI, "estado_b.json"), "utf8"));
  const b = core.buildBalance(r.cuentas, mapping, estado.saldos);
  check(Math.abs(b.totalDebe - r.control.debe) < 0.005 && Math.abs(b.totalHaber - r.control.haber) < 0.005,
    `julio sigue atando con Onvio: ${f2(b.totalDebe)} / ${f2(b.totalHaber)}`);
  check(core.findUnmapped(r.cuentas, mapping).length === 0, "y sin ninguna cuenta suelta");
} else {
  console.log("  (sin el export de julio a mano, salteo el control contra Onvio)");
}

console.log(fallos ? `\n${fallos} FALLA(S)` : "\ntodo OK");
process.exit(fallos ? 1 : 0);
