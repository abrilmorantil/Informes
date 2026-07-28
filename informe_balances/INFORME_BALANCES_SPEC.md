# Informe C — Balance Pesos y Dólares (especificación para implementar)

## Qué es

Tercer informe del menú "Informes". A partir de UN export de Onvio
("Balance de SyS por Cód. de Cta.", que trae ambas monedas), actualiza DOS
archivos Excel: el balance formal en pesos y el balance formal en dólares
(estados contables completos: Balance, Resultados, Pat.Neto, Activo y Pasivo,
Anexos I y II).

Archivos de referencia (en `informe_balances/ejemplos/`):
- `Balance_de_SyS_por_Cod__de_Cta_.xls` — el export de entrada (Crystal/Onvio).
- `SCA_Balance_a_06-2026_Pesos.xls` y `SCA_Balance_al_06-2026_dolares_.xls` — los
  dos archivos maestros a actualizar.
- `mapeo_balances_pesos_dolares.json` — mapeo extraído de ambos archivos
  (cuentas→fila de SALDOS, referencias de cada hoja de estados, conceptos de
  Anexo II, rubros de Nota 4). Generado con `extraer_mapeo_balances.py`.
- `equivalencias_dolares_propuestas.json` — ver sección "Códigos viejos" abajo.

## Cómo fluye la información dentro de cada archivo maestro

1. `Hoja1` = zona de pegado del export. Clave = texto "código - nombre" (col A).
   Valor = saldo (col **D** en dólares, col **E** en pesos).
2. `SALDOS` = lista ordenada por capítulo (ACTIVO/PASIVO/PATRIMONIO NETO/
   RESULTADOS). Cada fila busca su cuenta contra Hoja1 con VLOOKUP y expone el
   valor final en una columna (col **C** en dólares, col **G** en pesos).
3. Las hojas de estados se cuelgan de SALDOS:
   - `Activo y Pasivo` (Nota 4): rubros que suman cuentas puntuales de SALDOS
     (pasivo con signo invertido).
   - `Anexo II`: conceptos de gasto × 4 columnas (F=Administración,
     G=Comercialización, H=Exploración, I=Financieros).
   - `Anexo I` (Bienes de Uso): amortizaciones acumuladas desde SALDOS;
     valores de origen y aumentos son MANUALES.
   - `Balance`, `Resultados`, `Pat.Neto`: cúspide, totales de las anteriores.
     `Pat.Neto` tiene asientos manuales (aportes de capital, AREA).

## Qué automatiza la app (y qué NO)

AUTOMÁTICO:
- Parsear el export: capítulos por sección, cuentas "código - nombre",
  saldos de AMBAS monedas (detectar columnas por encabezado "Saldo (u$s)" /
  "Saldo ($)", nunca por posición — igual criterio que parser_onvio.py del
  Informe A).
- Reescribir `Hoja1` de cada archivo con los datos frescos (limpiar lo viejo
  primero — cada corrida reemplaza, no acumula).
- Detectar cuentas nuevas (no mapeadas en SALDOS de ese archivo).
- Capítulo de una cuenta nueva: AUTOMÁTICO por primer dígito del código
  (1=ACTIVO, 2=PASIVO, 3=PATRIMONIO NETO, 4=RESULTADOS). Validado: las 163
  cuentas del export de junio cumplen la regla sin excepción.
- Para cada cuenta nueva, preguntar al usuario (formulario web):
  a) línea de destino en la Nota 4 / Anexo II según capítulo:
     - ACTIVO/PASIVO → qué rubro de `Activo y Pasivo` (lista en
       `nota4_rubros` del mapeo)
     - RESULTADOS → qué concepto de `Anexo II` (lista en `anexo2_conceptos`)
       y qué columna (Administración/Comercialización/Exploración/Financieros)
  b) insertar la fila en SALDOS en el capítulo correcto (replicando el patrón
     de fórmula de las filas vecinas de ese capítulo) y agregar la referencia
     en la hoja de estados correspondiente.
- Chequeos finales: Activo total = Pasivo + PN (celda de control L23 de
  `Balance` debe dar 0), totales del export vs totales cargados, y lista de
  cuentas del export que no quedaron enganchadas a ningún estado.

MANUAL (la app NO lo toca, solo avisa):
- `Anexo I`: valores de origen, aumentos/disminuciones de bienes de uso.
- `Pat.Neto`: aportes de capital, AREA, ajustes.
- Ajustes manuales preexistentes en columnas auxiliares de SALDOS.

## Problemas de datos conocidos (medidos con el export de junio 2026)

1. **El archivo de DÓLARES usa códigos de cuenta VIEJOS (8 dígitos)** mientras
   el export usa el plan nuevo (9 dígitos). Solo 13/163 matchean por código
   directo. `equivalencias_dolares_propuestas.json` trae 31 matches automáticos
   por nombre (ALGUNOS SON DUDOSOS — ej. emparejó "IIBB saldo a favor" con
   "IVA saldo a favor") y 119 sin resolver.
   → DISEÑO REQUERIDO: pantalla de "revisión de equivalencias" la primera vez
   que se usa el informe: mostrar cada match propuesto para confirmar/corregir,
   y guardar la tabla validada en el mapeo. NO dar por buenos los matches
   automáticos sin confirmación humana.
2. **Proveedores**: dólares los agrupa en UNA línea ("Proveedores"); pesos los
   detalla uno por uno en SALDOS. Regla propuesta (a confirmar con la usuaria):
   en dólares, todas las cuentas 2110xxxxx suman a la línea Proveedores; en
   pesos, cada proveedor nuevo inserta su fila en el detalle.
3. El archivo de pesos tiene 7 cuentas duplicadas dentro de SALDOS y el de
   dólares 2 (ver `cuentas_duplicadas` en el mapeo). Además hay `#REF!` en
   `Balance`/`Anexo II` del de dólares y un rango sospechoso
   (`I102=SUM(I10:J100)`). Avisar, no "arreglar" en silencio.
4. En pesos, SALDOS es geométricamente caótico: cuentas en columnas C, D o E
   según la sección, VLOOKUP a veces en F y a veces en G, algunos saldos
   escritos a mano (ej. G100). Para insertar filas nuevas: copiar el patrón
   exacto de la fila vecina anterior del mismo capítulo.

## Decisiones ya tomadas por la usuaria

- Parametrizar (un solo motor, dos configs — ver `params` en el mapeo).
- Un solo ítem en el menú ("Balance Pesos y Dólares") que genera los dos
  archivos de una sola corrida.
- Capítulo automático por primer dígito; el formulario pregunta solo el
  destino fino (rubro/concepto/columna).

## Pendiente de decidir (preguntar antes de implementar)

- Confirmar la regla de proveedores (punto 2 de arriba).
- Si la salida pisa los archivos maestros o genera copias con nombre nuevo
  (recomendado: copias `SCA_Balance_<mes>_<moneda>.xls`, nunca pisar).
- Los `.xls` son formato viejo (BIFF). Escribirlos requiere convertirlos a
  `.xlsx` o mantenerlos: decidir formato de salida con la usuaria (openpyxl
  no escribe .xls; LibreOffice headless puede convertir en ambos sentidos).
