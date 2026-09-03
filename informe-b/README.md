# BALCOMPROBDOLARES — versión web (sin Python)

Misma app de siempre, pero corre 100% en el navegador. No hay que instalar
Python, ni Flask, ni nada — se abre con un link.

## Publicarla en GitHub Pages (una sola vez)

1. **Creá un repositorio nuevo en GitHub** (puede ser privado). Por ejemplo
   `balcomp-app`.
2. **Subí todos los archivos de esta carpeta** al repositorio (`index.html`,
   `app.js`, `core.js`, `writer.js`, `github.js`, `mapping.json`,
   `xlsx.full.min.js`, `exceljs.min.js`). La forma más simple: entrá al
   repo en github.com → "Add file" → "Upload files" → arrastrá todo.
3. En el repositorio: **Settings → Pages** → en "Source" elegí la rama
   donde subiste los archivos (`main`) y la carpeta `/ (root)` → Save.
4. GitHub te va a dar una URL parecida a
   `https://tuusuario.github.io/balcomp-app/` — esa es la app. Tarda un
   par de minutos en estar disponible la primera vez.

## Configurar el guardado automático a GitHub (una sola vez)

Para que la app pueda guardar las cuentas nuevas que vayas clasificando
directo en el repo (sin que tengas que descargar y subir el archivo a
mano):

1. Abrí la app → botón **"⚙ Configuración de GitHub"** arriba a la derecha.
2. **Generar el token** (el link "¿cómo lo genero?" en la app tiene estos
   mismos pasos):
   - Entrá a github.com → tu foto de perfil (arriba a la derecha) →
     **Settings**.
   - En el menú de la izquierda, abajo del todo: **Developer settings**.
   - **Personal access tokens → Tokens (classic)**.
   - **Generate new token (classic)**.
   - Ponele un nombre (ej. "balcomp"), marcá el casillero **`repo`**
     (acceso completo al repositorio) y generá.
   - Copiá el token (empieza con `ghp_`) — GitHub solo lo muestra una vez.
3. En la app, pegá ese token, y en "Usuario/organización y repositorio"
   poné algo como `tuusuario/balcomp-app` (el mismo repo del paso
   anterior). Dejá "Rama" en `main` y "Ruta" en `mapping.json`.
4. Guardar. Si dice "conectado correctamente", ya está.

Ese token queda guardado solo en tu navegador (en esta computadora) — no
se manda a ningún lado más que a GitHub.

## Uso diario

Exactamente el mismo flujo de siempre:

1. Subís el export de Onvio, el Sumas y Saldos (`.xls`).
2. Si aparecen cuentas nuevas, las clasificás — al apretar "Guardar
   clasificación en GitHub y recalcular" queda commiteado en el repo al
   toque, no hay que descargar/subir nada.
3. Revisás los chequeos.
4. Descargás el `.xlsx` final — se genera en tu navegador y se descarga
   directo a tu carpeta de Descargas, con la columna de Saldo anterior en
   amarillo para pegar a mano.

## Si alguna vez algo no anda

- **"No encontré mapping.json"**: revisá que el archivo esté en el
  repositorio y que la ruta configurada (⚙) coincida.
- **Botón de clasificar no guarda / da error**: revisá el token — puede
  haber vencido o no tener el permiso `repo` marcado. Generá uno nuevo.
- El **historial de cambios al mapeo** queda en GitHub como commits
  normales — podés ver quién cambió qué y cuándo desde la pestaña
  "Commits" del repositorio.

## Archivos

- `index.html` — la página.
- `app.js` — conecta todo con la pantalla.
- `core.js` — la lógica de siempre: parsea el export de Onvio, cruza con
  el mapeo, corre la cascada de validación.
- `writer.js` — genera el `.xlsx` final en el navegador.
- `github.js` — lee y guarda `mapping.json` en GitHub.
- `mapping.json` — la tabla de mapeo (se actualiza sola desde la app).
- `xlsx.full.min.js`, `exceljs.min.js` — las librerías que leen/escriben
  Excel, incluidas directo en el repo para no depender de ningún servicio
  externo.
