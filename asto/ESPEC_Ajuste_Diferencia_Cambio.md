# Especificación — Motor de "Ajuste por Conversión" / Diferencia de Cambio (bimonetario ARS/USD)

**Cliente de referencia:** Southern Copper Argentina S.R.L. (SCA)
**Objetivo:** automatizar por completo el armado del asiento de diferencia de cambio que hoy se hace a mano en `Asto_dif_cambio_<periodo>.xlsx`, para integrarlo como módulo de la app de informes (NestJS + Next.js).
**Entrada:** el Balance de Sumas y Saldos bimonetario (`SyS_<periodo>_sin_ajuste.xls`) + dos tipos de cambio de cierre.
**Salida:** archivo importador Onvio (formato "Asientos", idéntico a `0.xls`).

> **Estado de validación:** la lógica de esta especificación fue implementada y **contrastada contra el asiento real Nº 2630 (período 06-2026)**. Reproduce las **89 líneas materiales al centavo**; el asiento cierra exactamente en 0 y el único desvío frente al manual es **0,01 USD** de residuo en la cuenta de balanceo (por una línea sub-centavo que se incluyó manualmente). Ver sección 9.

---

## 1. Qué hace este asiento (contexto contable, para no perder el sentido)

SCA lleva contabilidad **bimonetaria**: cada cuenta tiene saldo en pesos (`$`) y en dólares (`u$s`). La moneda funcional / de presentación es el **dólar**. Al cierre de cada mes, las **partidas monetarias** (caja, bancos, créditos y débitos fiscales, cuentas con relacionadas, anticipos, proveedores, cargas sociales e impuestos a pagar) deben quedar re-expresadas en USD al **tipo de cambio de cierre**.

El asiento "Ajuste por Conversión" **sólo mueve el mayor en dólares (MEP)** y deja los pesos intactos. Para cada cuenta monetaria calcula cuál *debería* ser su saldo USD (`peso / TC de cierre`), lo compara con el saldo USD que hay en libros, y ajusta la diferencia. El neto de todos esos ajustes se imputa a la cuenta de resultado **`423050000 - Diferencia de Cambio USD`**, que hace de contrapartida y cierra el asiento.

> Nota: el workbook original tiene dos hojas de cálculo. **`TRADUCC` ("TRANSLATION DIFFERENCE") es la activa** y es la que describe esta espec (pesos → USD). La hoja `REVALUAC` es una plantilla vieja (2019, en cero) con el sentido inverso (USD → pesos) y **no se usa**; ignorarla.

---

## 2. Archivos de entrada y salida

| Rol | Archivo ejemplo | Formato | Descripción |
|---|---|---|---|
| **Entrada** | `SyS_06-2026_sin_ajuste__.xls` | `.xls` (Crystal Reports) | Balance de Sumas y Saldos bimonetario. **Único insumo de datos.** |
| **Entrada (parámetros)** | — | — | Fecha de cierre + T.C. compra + T.C. venta (los carga el usuario). |
| **Salida** | `0.xls` | `.xls` | Modelo importador de Onvio (hoja `Asientos`). El motor debe **producir un archivo con esta misma estructura**. |
| Referencia | `Grilla_Listado_de_Asientos_...xlsx` | `.xlsx` | Export del asiento real **Nº 2630** ya contabilizado. Sirve como *golden file* para tests. |
| Referencia | `Asto_dif_cambio_6-2026.xlsx` | `.xlsx` | Workbook manual actual (hojas `TRADUCC`, `Proveedores (7)`, `PARA IMPORTAR`). Es lo que estamos reemplazando. |

---

## 3. Estructura del SyS (parseo)

Hoja única. Las columnas relevantes (índice 0-based):

| Columna | Índice | Header (fila 10) | Uso |
|---|---|---|---|
| A | 0 | `Cuenta - Denominación` | `"111010001 - Caja"` → se separa en código (9 díg.) + denominación. También trae los rótulos de sección. |
| P | 15 | `Saldo ($)` | **Saldo en pesos** de la cuenta. |
| AB | 27 | `Saldo (u$s)` | **Saldo en dólares en libros** (histórico, a ajustar). |

**Secciones:** el texto en la columna A marca los capítulos, en este orden: `ACTIVO`, `PASIVO`, `PATRIMONIO NETO`, `RESULTADOS`. Cada cuenta pertenece a la última sección vista. Sólo interesan `ACTIVO` y `PASIVO` (más la cuenta de resultado de balanceo, que se agrega aparte).

**Regex de cuenta:** `^(\d{9})\s*-\s*(.*)$` sobre el texto de la columna A. Si no matchea (o es un rótulo de sección), se saltea la fila.

**Cuidado:** el header dice `AA` pero el valor del saldo USD cae en `AB` (índice 27). Usar índice 27.

---

## 4. Parámetros del período (los carga el usuario en la UI)

| Parámetro | Ejemplo 06-2026 | Notas |
|---|---|---|
| `periodo_fin` | `2026-06-30` | Último día del mes. Es la fecha del asiento. |
| `tc_compra` | `1473` | T.C. **COMPRA** de cierre (BNA). Se aplica al **ACTIVO**. |
| `tc_venta` | `1482` | T.C. **VENTA** de cierre (BNA). Se aplica al **PASIVO**. |
| `numero_asiento` | `1` | Opcional; número de asiento para el importador. |
| `concepto` | `"Ajuste por Conversión"` | Texto fijo. |

---

## 5. Clasificación de cuentas — **el corazón de la lógica**

Sólo se ajustan las **partidas monetarias** del ACTIVO y del PASIVO. La selección resulta de combinar **dos filtros**:

### 5.1 Filtro contable (config persistente): monetaria vs. no monetaria

Una cuenta **NO es monetaria** (y por lo tanto se excluye del asiento) si cumple alguna:

- Su código empieza con un **prefijo no monetario**: `"1240"` (Bienes de Uso y sus Depreciaciones Acumuladas) o `"1250"` (Cargos Diferidos).
- Está en la **lista de exclusión puntual**:
  - `114010016` — Impuesto Crédito Diferido
  - `114050005` — Seguros a Devengar (pago anticipado)
  - `211050000` — Previsión AID (impuesto diferido pasivo)
- Pertenece a `PATRIMONIO NETO` o `RESULTADOS` (se excluyen enteras).

> Estas cuentas tienen diferencia peso/USD **real** (se mantienen a costo histórico), así que **no se auto-excluyen** por materialidad: hay que excluirlas por criterio. Esta lista es **configuración que se mantiene entre períodos** y cambia rara vez. En la app conviene guardarla en una tabla editable por el usuario (ver sección 10), no hardcodeada.

Todo lo demás dentro de `ACTIVO` / `PASIVO` es **monetario** y entra como candidato.

### 5.2 Filtro de materialidad (automático)

Se publica la línea sólo si `|round(ajuste_usd, 2)| ≥ 0.005` (es decir, redondea a ≥ 0,01 USD).

Esto **auto-excluye** dos casos frecuentes sin necesidad de configurarlos:
- **Cuentas denominadas en USD** (ej. `Banco Macro (u$s)`, `Alonso Miguel Ángel (u$s)`, `Datamine Chile`, proveedores USD): su saldo en pesos ya es `USD × TC_cierre`, así que el ajuste da ≈ 0.
- **Cuentas ya cuadradas** o de saldo cero.

> **Verificación cruzada útil:** para una cuenta USD-denominada del pasivo, `saldo_pesos / saldo_usd ≈ 1482` (venta); para una del activo, `≈ 1473` (compra). Ese cociente igual al TC de cierre es la firma de "ya está en dólares, no ajustar".

---

## 6. Cálculo del ajuste (paso a paso)

Para cada cuenta que pase el filtro de la sección 5:

1. **Elegir el tipo de cambio según la sección:**
   - `ACTIVO` → `tc = tc_compra` (1473)
   - `PASIVO` → `tc = tc_venta` (1482)

2. **Calcular el USD teórico** (lo que el saldo USD *debería* ser):
   - **Cuentas de proveedores** (código empieza con `"21101"`): **sin redondear**
     `usd_teorico = saldo_pesos / tc`
     *(replica el subdiario de proveedores: `C = B/TC`, `E = C − D`)*
   - **Resto de las cuentas**: **redondeado a 2 decimales**
     `usd_teorico = round(saldo_pesos / tc, 2)`
     *(replica la hoja `TRADUCC`: `E = ROUND(C/TC, 2)`, `I = E − G`)*

3. **Ajuste con signo:**
   `ajuste_usd = usd_teorico − saldo_usd`

4. **Filtro de materialidad** (sección 5.2). Si pasa, es una línea del asiento.

> El **signo** del ajuste es la clave del asiento: **positivo = Debe MEP**, **negativo = Haber MEP**. No hay que separar en dos columnas: el importador usa un único importe firmado (sección 8).

---

## 7. Cuenta de balanceo (contrapartida de resultado)

Después de calcular todas las líneas:

1. Redondear cada `ajuste_usd` a 2 decimales (ese es el importe que se publica).
2. `neto = round(Σ ajuste_usd de todas las líneas, 2)`
3. Agregar una línea final:
   - **Código:** `423050000`
   - **Denominación:** `Diferencia de Cambio USD`
   - **Importe USD:** `round(-neto, 2)`

Así el asiento **cierra exactamente en 0** (Σ de todas las líneas incluida la de balanceo = 0), y esa cuenta de resultado absorbe además el residuo de redondeo. En 06-2026 su importe es **+58.374,80 USD (Debe MEP)**.

---

## 8. Formato de salida — archivo importador Onvio (idéntico a `0.xls`)

Hoja `Asientos`. Fila 1 = headers (exactos, respetar tildes y puntos). Una fila por línea del asiento (incluida la de balanceo).

| Col | Header exacto | Contenido |
|---|---|---|
| A | `Número de asiento` | `numero_asiento` (mismo para todas las filas, ej. `1`). |
| B | `Número de Pase` | Número de línea correlativo `1, 2, 3, …` |
| C | `Fecha` | `periodo_fin` como **serial de Excel** (base 1899-12-30). Ej. `2026-06-30` → `46203`. Aplicar formato de fecha a la celda. |
| D | `Concepto` | `"Ajuste por Conversión"` |
| E | `Código de cuenta` | **Código completo de 9 dígitos** (ej. `"111010001"`). *(Los códigos cortos `14111/12512/15555` del template `0.xls` son datos dummy: ignorarlos. La hoja `PARA IMPORTAR` del workbook confirma que se importa con el código de 9 dígitos.)* |
| F | `Importe en moneda local` | **`0`** (el asiento no toca pesos). |
| G | `Importe en moneda ext.present.` | **`round(ajuste_usd, 2)` con signo** (positivo = Debe MEP, negativo = Haber MEP). |
| H | `Leyenda` | Vacío (opcional). |
| I | `Código de centro de costos` | Vacío. |
| J | `Porcentaje de distribución` | Vacío. |
| K | `Imp.mon.local dist.C.Costos` | Vacío. |
| L | `Imp.mon.present.dist.C.Costos` | Vacío. |

**Serial de fecha:** `serial = (date(y,m,d) − date(1899,12,30)).days`. Comprobado: `2026-06-30 → 46203`.

> El export de referencia (`Grilla`) usa columnas separadas `Debe MEP` / `Haber MEP`; el **importador** (`0.xls`) usa **un solo importe firmado en G**. La relación es: `G = DebeMEP − HaberMEP`. Y además `G` coincide directamente con `ajuste_usd`.

---

## 9. Validación (tests de aceptación)

Usar `Grilla_Listado_de_Asientos_...xlsx` (asiento real Nº 2630, 06-2026) como *golden file*:

- Construir de él un dict `{codigo: round(DebeMEP − HaberMEP, 2)}` (columnas K y L).
- Correr el motor sobre `SyS_06-2026_sin_ajuste__.xls` con `tc_compra=1473`, `tc_venta=1482`.
- **Aserciones:**
  1. Toda cuenta con `|importe| ≥ 0.01` en el golden aparece en el generado con **el mismo importe (tolerancia 0,01)**.
  2. El asiento generado **suma 0** en USD.
  3. La línea `423050000` da ≈ `+58.374,79/80` USD.
  4. No aparecen cuentas no monetarias (`1240xx`, `1250xx`, `114010016`, `114050005`, `211050000`).

**Resultado obtenido con la implementación de referencia:** 89/89 líneas materiales coinciden al centavo; el asiento suma 0; la cuenta de balanceo difiere en **0,01 USD** frente al manual, por una línea sub-centavo (proveedor *Toro Franco*, 0,0026 USD) que el operador incluyó a mano. Es inmaterial y el residuo queda absorbido en la cuenta de resultado. Si se quisiera replicar el manual bit-a-bit, bajar el umbral de materialidad de proveedores a 0, pero **no se recomienda** (agrega ~9 líneas de importe cero).

---

## 10. Modelo de datos y firmas sugeridas (para el port a NestJS/TypeScript)

**Config persistente** (tabla editable, no hardcode):

```ts
// tabla: exchange_diff_config (por cliente)
interface ExchangeDiffConfig {
  clienteId: string;
  nonMonetaryPrefixes: string[];   // ["1240", "1250"]
  nonMonetaryExact: string[];      // ["114010016","114050005","211050000"]
  supplierPrefix: string;          // "21101"
  cuentaBalanceo: string;          // "423050000"
  denomBalanceo: string;           // "Diferencia de Cambio USD"
  materialidad: number;            // 0.005
  concepto: string;                // "Ajuste por Conversión"
}
```

**Parámetros del período** (input de la UI):

```ts
interface Parametros {
  periodoFin: Date;      // cierre
  tcCompra: number;      // ACTIVO
  tcVenta: number;       // PASIVO
  numeroAsiento?: number;
}
```

**Pipeline:**

```ts
type CuentaSaldo = { codigo: string; denominacion: string;
                     seccion: 'ACTIVO'|'PASIVO'|'PATRIMONIO NETO'|'RESULTADOS';
                     saldoPesos: number; saldoUsd: number };

type LineaAjuste = { codigo: string; denominacion: string; seccion: string;
                     tcAplicado: number; usdTeorico: number; usdLibros: number;
                     ajusteUsd: number };

parseSysXls(buffer): CuentaSaldo[]              // leer el .xls (col A, P=15, AB=27)
esMonetaria(c, cfg): boolean                    // sección 5.1
calcularLineas(cuentas, params, cfg): LineaAjuste[]   // secciones 5.2 + 6
armarAsiento(lineas, cfg): LineaAjuste[]        // sección 7 (+ balanceo)
escribirImportadorXls(asiento, params): Buffer  // sección 8 (hoja "Asientos")
```

**Librerías Node sugeridas:** para leer `.xls` viejos → `xlsx` (SheetJS) leyendo con `cellDates`/`raw`; para escribir el `.xls` de salida → SheetJS también (`XLSX.utils.aoa_to_sheet` + `bookType:'xls'`), o si hay problemas de compatibilidad con Onvio, generar `.xlsx` y confirmar que el importador lo acepta. El serial de fecha se puede escribir con celda tipo fecha o directamente como número.

---

## 11. Casos borde y notas de mantenimiento

- **Cuentas que aparecen/desaparecen cada mes:** el motor es data-driven sobre el SyS, así que soporta altas/bajas de cuentas monetarias sin tocar código. Sólo hay que mantener la lista de **no monetarias** (sección 5.1) cuando se crea una cuenta nueva de bienes de uso, cargos diferidos, previsiones o diferidos impositivos fuera de los prefijos `1240/1250`.
- **Nuevos prefijos de proveedores:** si el plan agrega proveedores fuera de `21101`, actualizar `supplierPrefix` (o volverlo lista de prefijos).
- **Un TC por sección** es la regla actual (compra=activo, venta=pasivo). Si algún día se necesitara un TC por cuenta, extender la config a un override por código.
- **Redondeo:** proveedores sin redondear el cociente; el resto redondeado a 2. Mantener esa distinción para que el neto cierre igual que el manual.
- **La cuenta de balanceo** siempre absorbe el residuo → el asiento cierra en 0 por construcción; conviene una aserción en runtime `assert Σ importes == 0`.
- **Validación de saldos:** si el SyS trae `Saldo (u$s)` en una columna distinta en otro export, parametrizar los índices de columna (P/AB) en la config del cliente.

---

## 12. Mantenimiento del plan de cuentas y salvaguarda interactiva

Toda cuenta nueva respeta la numeración de su familia, así que la clasificación por prefijo es la columna vertebral y se auto-mantiene en la mayoría de los casos. Hay que distinguir dos situaciones:

### 12.1 Se resuelve solo (cero mantenimiento)
Cuando la cuenta nueva cae en una familia ya clasificada por prefijo:
- Bien de uso nuevo → `1240xxxxx` → excluida automáticamente.
- Cargo diferido nuevo → `1250xxxxx` → excluida automáticamente.
- Cualquier cuenta **monetaria** nueva (banco, crédito fiscal, proveedor, carga social) → entra al cálculo sola, sin tocar nada.

### 12.2 Necesita decisión (el único caso no mecánico)
Una cuenta **no monetaria que nace *dentro* de un prefijo monetario**. Son las tres excepciones actuales: `114010016` (impuesto diferido) dentro de `1140`, `114050005` (seguros a devengar) dentro de `11405`, `211050000` (previsión AID) dentro de `2110`. El prefijo no alcanza —ni siquiera el de 5 dígitos, porque `21105` (Previsión AID, no monetaria) convive con `21106` (Previsión IGMP, monetaria)—, así que van en la **lista de exclusión puntual mantenida** (`nonMonetaryExact`).

### 12.3 Salvaguarda: detectar, avisar y preguntar
Para que 12.2 no se escape en silencio, el módulo debe **detectar automáticamente las líneas sospechosas y pausar para preguntarle al usuario si las incluye o no**, antes de generar el importador.

**Criterio de sospecha** (se marca la línea que cumple **las dos** condiciones a la vez):
- **Desproporción:** `|ajuste_usd| > umbral_ratio × |saldo_usd|` con **`umbral_ratio = 0.30`**.
- **Y materialidad:** `|ajuste_usd| > umbral_abs` con **`umbral_abs = 500` USD**.

> **Calibración con datos reales (06-2026):** las cuentas no monetarias mal incluidas dan ratio **~0,95–0,98** (se mantienen a TC histórico, muy lejos del de cierre); las monetarias legítimas dan **~0,05** (el TC se movió ~5 % en el mes). La brecha es enorme, así que `0.30` las separa sin ambigüedad. Exigir **ambas** condiciones evita los falsos positivos de saldo chico (ej. Caja: ajuste 100 % del saldo pero solo 24 USD → no se marca). Test verificado: con la config real marca **0 líneas**; simulando el olvido de clasificar (sin exclusiones) **captura las 20** cuentas no monetarias (bienes de uso, depreciaciones, cargos y créditos diferidos) y ninguna monetaria legítima.

Opcionalmente, marcar también como sospechosa una cuenta **nueva**: código que no figura en la config como ya clasificado (ni en `nonMonetaryExact` ni en la allowlist `confirmedMonetary`) y que no cae en un prefijo no monetario. Requiere persistir un registro de cuentas ya vistas.

> Nota: `114050005 Seguros a Devengar` (no monetaria, ratio 0,06) **no** dispara la salvaguarda —se ve igual que una monetaria—, por eso permanece en la lista mantenida `nonMonetaryExact`. La salvaguarda es la red para las no monetarias evidentes (activos/diferidos a histórico), no un reemplazo de la lista.

**Flujo (dos fases, apto para la app):**
1. **Fase cálculo/revisión.** El motor calcula todas las líneas y devuelve dos grupos: `lineasOk` (pasan sin dudas) y `lineasARevisar` (marcadas por el criterio de sospecha), sin generar todavía el importador.
2. **Fase confirmación (UI).** Si `lineasARevisar` no está vacía, la app **avisa y pregunta por cada cuenta**: muestra código, denominación, saldo pesos, saldo USD libros, TC aplicado, USD teórico y ajuste, con dos acciones: **Incluir** / **Excluir**.
3. **Persistir la decisión** para no volver a preguntar cada mes:
   - Si el usuario **excluye** → agregar el código a `nonMonetaryExact` (es no monetaria).
   - Si el usuario **incluye** → agregar el código a una allowlist `confirmedMonetary` (queda confirmada y no se vuelve a marcar como nueva).
4. **Recién entonces** armar el asiento (con las incluidas) y generar el importador.

> Regla de oro: **ninguna cuenta sospechosa entra al importador sin confirmación explícita del usuario.** El resto del flujo es automático.

**Firmas sugeridas:**
```ts
type LineaRevisar = LineaAjuste & { motivo: 'ratio' | 'monto' | 'nueva' };

// Fase 1: no genera nada, solo separa
calcularConRevision(cuentas, params, cfg):
    { lineasOk: LineaAjuste[]; lineasARevisar: LineaRevisar[] };

// Fase 2 (tras la UI): decisiones = { [codigo]: 'incluir' | 'excluir' }
aplicarDecisiones(lineasARevisar, decisiones):
    { incluidas: LineaAjuste[]; nuevaConfig: Partial<ExchangeDiffConfig> };
// nuevaConfig trae los códigos a persistir en nonMonetaryExact / confirmedMonetary

// Config extendida
interface ExchangeDiffConfig {
  // ...campos previos...
  confirmedMonetary: string[];   // cuentas monetarias ya confirmadas por el usuario
  umbralRatio: number;           // 0.15
  umbralAbs: number;             // 1000
}
```

Agregar al final una **aserción de cierre** (`Σ importes == 0`) como red de contención independiente de todo lo anterior.

---

## Anexo — Resumen de la mecánica en una línea

> Para cada cuenta **monetaria** del ACTIVO/PASIVO: `ajuste_USD = (saldo_pesos / TC_cierre) − saldo_USD_libros`, con **TC compra** para activos y **TC venta** para pasivos, **redondeando salvo proveedores**; se publican las de `|ajuste| ≥ 0,01`; la contrapartida por el neto va a **`423050000 Diferencia de Cambio USD`**; y se exporta al importador con **pesos = 0** y **USD firmado** (positivo Debe MEP / negativo Haber MEP), fecha = cierre del período.
