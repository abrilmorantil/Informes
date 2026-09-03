# Informes — Southern Copper Argentina

Sitio único con los dos informes. Corre entero en el navegador: no hay que instalar
Python ni arrancar ningún servidor, se entra por un link.

```
index.html      pantalla de inicio: elegís qué informe generar
estilos.css     paleta y componentes compartidos
informe-a/      Balance de Comprobación USD  (export de Onvio)
informe-b/      BALCOMPROBDOLARES            (export de Onvio)
```

## Cómo publicarlo

Subí **el contenido de esta carpeta** a la raíz del repositorio que ya tenés publicado
en GitHub Pages (el mismo de BALCOMPROBDOLARES): entrá al repo en github.com →
"Add file" → "Upload files" → arrastrá todo.

**No hay que cambiar ninguna configuración de BALCOMPROBDOLARES.** Su `mapping.json` se
deja donde está, en la raíz del repositorio, y su ⚙ sigue apuntando ahí. Es a propósito:
ese archivo lo actualiza la app sola cada vez que se clasifica una cuenta, así que moverlo
obligaría a sincronizarlo a mano con riesgo de pisar clasificaciones.

---

## Balance de Comprobación USD (informe-a)

### Cómo funciona el archivo (importante)

El balance **se calcula solo**. La hoja `Sumas y Saldos` no guarda importes: los trae de
la hoja `SyS` con fórmulas `VLOOKUP`, y de ahí salen `Dist.de gastos`, los totales y la
columna de totales por cuenta.

Por eso el motor escribe **únicamente en la hoja `SyS`**. Todo lo demás se completa con
las fórmulas que el archivo ya tiene. Escribir importes a mano en `Sumas y Saldos`
pisaría esas fórmulas y rompería el archivo.

Lo único que se toca fuera de `SyS` es cuando aparece una cuenta que todavía no existe:
ahí hay que crearle la fila (copiándole las fórmulas de una fila vecina) y referenciarla
en `Dist.de gastos`.

### Las columnas de mes

En `Dist.de gastos` hay una columna por mes (`TOTAL ENERO` … `TOTAL DICIEMBRE`) más
`TOTAL AÑO`. Los meses cerrados guardan un importe fijo; **el mes en curso es el único
"vivo"**, con la fórmula `=D<fila>`, que sigue a la columna MOVIMIENTO MES.

De ahí salen las dos reglas del sistema:

- **No se puede saltear un mes.** El mes a cargar no se elige de una lista: sale del
  último mes cerrado que quedó guardado. Si se saltearan meses, el TOTAL AÑO quedaría
  corto sin que nadie se entere.
- **No puede haber dos meses abiertos.** Si el mes anterior quedó sin cerrar, cargar el
  siguiente le escribiría los importes encima. El motor lo detecta y frena antes de
  generar nada.

### Uso

**La primera vez** se sube el archivo del balance tal como esté. La app detecta sola cuál
es el mes en curso y pregunta si ya está cerrado. A partir de ahí el archivo queda
guardado en el repositorio y no hay que volver a subirlo.

**Cada mes:**

1. La pantalla muestra qué mes toca cargar. Se sube el export de Onvio.
2. Si hay cuentas nuevas, se elige la categoría de cada una.
3. Se descarga el **borrador**, las veces que haga falta. No guarda nada.
4. Se revisa en Excel y, si está bien, **se guarda ahí mismo** (Ctrl+G).
5. Se sube ese archivo y se **cierra el mes**. Recién ahí queda guardado y se habilita
   el mes siguiente.

El paso 4 no es un capricho: para cerrar un mes hay que dejar escrito el importe
calculado, y quien calcula esas fórmulas es Excel. La columna `E` de `Dist.de gastos`
tiene fórmulas variadas (referencias a Debe, a Haber, negativas, y una que apunta a otra
celda de la hoja), así que replicar ese cálculo en el navegador sería frágil y podría dar
totales mal sin avisar. Usando el archivo que Excel ya calculó, el importe que se congela
es exactamente el que se vio en pantalla.

### Qué se guarda en el repositorio

| Archivo | Qué es |
|---|---|
| `base_actual.xlsx` | el balance vivo, con los meses cerrados congelados |
| `mapeo.json` | dónde está cada cuenta y a qué centro de costo corresponde cada columna |
| `estado.json` | cuál fue el último mes cerrado, y el historial |

El mapeo **se deriva del propio archivo**, mirando las fórmulas y no los títulos de las
columnas. Los títulos no son confiables: la columna `L` de `Dist.de gastos` dice
"LOS MORTERITOS" pero suma PROYECTO LOS MORTERITOS.

### Pendientes conocidos

- **SAMENTA no tiene columna en `Dist.de gastos`.** Si algún mes trae movimiento en ese
  centro de costo, el importe queda en `Sumas y Saldos` pero no aparece en la
  distribución. La app avisa en pantalla cuando pasa. Lo mismo con TRES CONOS, LA CHILENA,
  SOL y LOS MORTERITOS, que hoy están sin uso.
- **La columna `S`** dice "PROYECTO SHEYLA I" pero varias de sus fórmulas suman el Haber
  de SHEYLA. Queda para revisar con el archivo abierto.

### Archivos

- `motor.js` — la carga: limpia `SyS`, escribe los importes del mes, inserta las cuentas
  nuevas y activa la columna del mes. También cierra el mes.
- `meses.js` — las columnas de mes: cuál está viva, activarla y congelarla.
- `mapeo.js` — deriva del archivo base dónde está cada cuenta y cada centro de costo.
- `parser_onvio.js` — lee el export de Onvio. Ubica las columnas en dólares por su
  encabezado, nunca por posición, para no tomar por error las columnas en pesos.
- `formula_utils.js` — reacomoda las fórmulas de todo el archivo cuando se inserta una
  fila, incluidas las de la propia hoja (Excel lo hace solo; ExcelJS no).
- `similitud.js` — emparejamiento de nombres de centro de costo entre Onvio y el balance.
- `github.js` — lee y guarda el archivo base, el mapeo y el estado.
