# Informes — Southern Copper Argentina

Sitio único con los dos informes. Corre entero en el navegador: no hay que instalar
Python ni arrancar ningún servidor, se entra por un link.

```
index.html      pantalla de inicio: elegís qué informe generar
estilos.css     paleta y componentes compartidos
informe-a/      Balance de Comprobación USD  (export de Onvio)
informe-b/      BALCOMPROBDOLARES            (export de SISE)
```

## Cómo publicarlo

Subí **el contenido de esta carpeta** a la raíz del repositorio que ya tenés publicado
en GitHub Pages (el mismo de BALCOMPROBDOLARES). La forma más simple: entrá al repo en
github.com → "Add file" → "Upload files" → arrastrá todo.

Después de subirlo, el link de siempre abre la pantalla de "Informes" y desde ahí se
entra a cualquiera de los dos. **No hay que cambiar ninguna configuración.**

### Por qué el mapeo de BALCOMPROBDOLARES sigue en la raíz

BALCOMPROBDOLARES pasó a vivir en `informe-b/`, pero su `mapping.json` **se deja donde
está, en la raíz del repositorio**, y la configuración (⚙) se deja apuntando ahí.

Es a propósito: ese archivo lo viene actualizando la app sola cada vez que se clasifica
una cuenta, así que es el que tiene los datos buenos. Moverlo obligaría a sincronizarlo a
mano y se corre el riesgo de pisar clasificaciones. Dejándolo quieto, la app sigue
funcionando igual que siempre.

El `informe-b/mapping.json` que viene en esta carpeta queda solo como copia de respaldo y
no se usa mientras la configuración de GitHub esté puesta.

## Balance de Comprobación USD (informe-a)

Reemplaza al motor de Python que se corría por consola. Hace exactamente lo mismo:

1. Subís el balance maestro (`.xlsm`) y el export de Onvio (`.xls`), y ponés el período.
2. Si aparece una cuenta que todavía no está en el balance, te la muestra en una tabla
   para que elijas a qué categoría pertenece. El motor le crea la fila en el lugar
   correcto y reacomoda todas las fórmulas que dependían de esa posición.
3. Descargás el balance ya cargado.

**El archivo sale en `.xlsx`**, no `.xlsm`: las macros viejas ya no se usan (este motor
reemplaza a `CargarOnvio`). Las hojas, las fórmulas y el formato se mantienen, y Excel
recalcula los totales al abrirlo.

### Guardar las cuentas nuevas

Igual que en BALCOMPROBDOLARES, se configura con ⚙ un token de GitHub y queda guardado
solo en tu navegador. La ruta por defecto del mapeo es `informe-a/mapeo_maestro.json`.

El botón para guardar aparece **solo cuando hubo cuentas nuevas**. Es a propósito: al
insertarse una fila, las cuentas que estaban más abajo cambian de posición, así que a
partir de ese momento el mapeo corresponde al archivo que generó esa corrida. Si el mes
siguiente arrancás del balance viejo, el sistema lo detecta y frena antes de escribir
nada, avisando qué pasó.

### Archivos

- `motor.js` — la lógica de carga: limpia los saldos del mes anterior, carga los nuevos,
  inserta cuentas nuevas, reconstruye la fila TOTALES y la columna de totales por cuenta.
- `parser_onvio.js` — lee el export de Onvio. Ubica las columnas en dólares por su
  encabezado, nunca por posición, para no tomar por error las columnas en pesos.
- `formula_utils.js` — reacomoda las fórmulas de todo el archivo cuando se inserta una
  fila (ni Excel-en-el-navegador ni el Python original lo hacen solos).
- `similitud.js` — emparejamiento de nombres de centro de costo entre Onvio y el balance.
- `github.js` — lee y guarda el mapeo en el repositorio.
- `mapeo_maestro.json` — la "memoria": qué fila ocupa cada cuenta, sus categorías y los
  22 centros de costo con sus columnas.
