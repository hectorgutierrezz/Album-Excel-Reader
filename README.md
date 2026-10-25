# Music Library Browser

Explorador y filtro interactivo para tu biblioteca personal de música en Excel. Funciona como **web estática** (navegador o GitHub Pages): al abrir la página descarga el Excel desde **Google Drive** y lo procesa en el navegador.

Sitio online: `https://hectorgutierrezz.github.io/Album-Excel-Reader/`

---

## Cómo usarlo

### Online
Abre el enlace de GitHub Pages. Si configuraste el proxy de Drive (abajo), la biblioteca se actualiza sola cada vez que entras.

### En local
```bash
python3 -m http.server 8080
```
Luego abre `http://localhost:8080`. Sin `driveProxyUrl`, carga el fichero local `Albums List.xlsx`.

---

## Autoactualizar desde Google Drive

Los navegadores no pueden leer Drive directamente (CORS). Hace falta un **proxy mínimo** con Google Apps Script que lee tu **Google Sheet** y la sirve como JSON.

### 1. ID de la hoja
Abre la hoja en Drive/Sheets. La URL es del estilo:
`https://docs.google.com/spreadsheets/d/FILE_ID/edit`  
Copia el `FILE_ID`.

### 2. Crear el proxy (Apps Script)
1. Ve a [script.google.com](https://script.google.com) → **New project**.
2. Sustituye el código por el de `drive-proxy/Code.gs`.
3. Pon tu `FILE_ID` en `DRIVE_FILE_ID`.
4. **Deploy** → **New deployment** → tipo **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Autoriza acceso a Sheets cuando lo pida.
6. Copia la URL que termina en `/exec`.

Si cambias el código después: **Implementar → Gestionar implementaciones → lápiz → Nueva versión → Implementar**.

### 3. Conectar la web
En `js/config.js`:
```js
driveProxyUrl: 'https://script.google.com/macros/s/XXXX/exec',
```
Haz commit y push para que GitHub Pages lo use.

Cada visita carga los datos actuales de la Sheet. Si el proxy falla, la web intenta el `Albums List.xlsx` local como respaldo.

---

## Filtros

| Filtro | Cómo funciona |
| :--- | :--- |
| **Artista / Banda** | Texto (sin distinguir mayúsculas). **Pick from list** para elegir de la lista. |
| **Varios artistas (`;`)** | Colaboraciones tipo `Bruno Mars; Anderson .Paak` se indexan por cada artista. |
| **Nombre del registro** | Coincidencia parcial en el título. |
| **Género e Idioma** | Selector emergente con buscador. |
| **Tipo** | `Album`, `EP`, `Single`, etc. |
| **Años** | *From* / *To*. |
| **Calificación** | Operador (`>=`, `<=`, `==`) y deslizador 0–10. |
| **Apply / Clear** | Aplica o limpia filtros. |
| **Vistas** | **DataFrame** (texto monoespaciado) o **Data Table** (columnas ordenables). |
| **Zoom** | `-` / `+`. Clic en una fila para resaltarla. |
| **Tema** | Oscuro, claro o sistema. |

---

## Formato del Excel

Columnas en la primera fila:

* **RECORDS** — título del álbum o disco
* **CREATOR** — artista(s); varios separados con `;`
* **TYPE** — `Album`, `EP`, etc.
* **GENRES** — géneros (coma o `/`)
* **LANGUAGE** — idioma
* **YEAR** — año (4 dígitos)
* **MARK(/10)** o **MARK** — nota sobre 10
