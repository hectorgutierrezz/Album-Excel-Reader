# 🎵 Music Library Browser

Un explorador y filtro interactivo para tu biblioteca personal de música en Excel (`Albums List.xlsx`). Permite consultar, buscar y filtrar tu colección de álbumes por artista, título, género, idioma, año y calificación.

La aplicación está disponible en **3 modalidades**:
1. **🌐 Versión Web estática** (desde el navegador o GitHub Pages, sin necesidad de instalar nada).
2. **🖥️ Aplicación de escritorio** (interfaz gráfica moderna con CustomTkinter).
3. **💻 Modo terminal** (interfaz por consola basada en pandas).

---

## 🚀 Cómo acceder y utilizar el programa

### Opción 1: Versión Web (Recomendada - Sin instalación)

Puedes usar la aplicación directamente desde cualquier navegador web moderno:

* **Enlace online**: Accede directamente a:  
  `https://hectorgutierrezz.github.io/Album-Excel-Reader/`
* **Ejecutar en local**: Si tienes el proyecto descargado en tu ordenador, abre una terminal en la carpeta y ejecuta:
  ```bash
  python3 -m http.server 8080
  ```
  Luego abre `http://localhost:8080` en tu navegador.

### Opción 2: Aplicación de Escritorio con Interfaz Gráfica (GUI)

#### A) Con el ejecutable listo para usar (sin necesidad de tener Python instalado)
Coloca en una misma carpeta el archivo ejecutable y tu hoja de cálculo:
```text
readerUI (o readerUI.exe en Windows)
Albums List.xlsx
```
* **En Linux**:
  ```bash
  chmod +x readerUI
  ./readerUI
  ```
* **En Windows**: Haz doble clic en `readerUI.exe`.

#### B) Ejecutando desde el código fuente con Python
Si tienes Python instalado:
```bash
# 1. Crear y activar entorno virtual
python3 -m venv .venv
source .venv/bin/activate       # En Linux/macOS
# .venv\Scripts\activate        # En Windows

# 2. Instalar dependencias
pip install pandas openpyxl customtkinter

# 3. Lanzar la aplicación
python readerUI.py
```

---

### Opción 3: Modo Terminal (Consola)

Para consultar tus datos de manera rápida desde la línea de comandos:
```bash
python readerTerminal.py
```
Sigue las opciones numéricas del menú interactivo para aplicar filtros individuales o combinados.

---

## 🔍 Guía de Filtros y Funcionalidades

| Filtro / Función | Cómo funciona |
| :--- | :--- |
| **Artista / Banda** | Búsqueda por texto (insensible a mayúsculas/minúsculas). También puedes pulsar **Pick from list** en la web para elegir de la lista ordenada de artistas. |
| **Múltiples artistas (`;`)** | Los álbumes con colaboraciones separadas por punto y coma (ej. `Bruno Mars; Anderson .Paak` o `Clipse; Pusha T; Malice`) se indexan para cada artista individualmente. Nombres propios con coma (como `Tyler, The Creator`) no se dividen. |
| **Nombre del registro** | Busca por coincidencias parciales en el título del álbum o EP. |
| **Género e Idioma** | Pulsa el botón del filtro para abrir un selector emergente con buscador integrado para encontrar rápidamente cualquier género o idioma. |
| **Tipo de registro** | Desplegable para filtrar entre `Album`, `EP`, `Single`, etc. |
| **Rango de años** | Campos *From* y *To* para limitar por año de lanzamiento (ej. de 2018 a 2022). |
| **Calificación (Nota)** | Elige el operador (`>=`, `<=`, `==`) y ajusta la barra deslizante entre 0.0 y 10.0. Admite tanto puntos como comas decimales (ej. `8.5` y `8,5`). |
| **Botones Apply / Clear** | **Apply** aplica los filtros activos. **Clear** restablece todos los filtros para mostrar la colección completa. |
| **Cambio de Vista (Web)** | Alterna entre **DataFrame (App)** (vista monoespaciada exacta a la aplicación original de Python) y **Data Table (Web)** (tabla interactiva con columnas ordenables por clic). |
| **Zoom y Resaltado** | Botones **`-`** y **`+`** para cambiar el tamaño de letra. Haz clic sobre cualquier fila para resaltarla en azul. |
| **Tema Claro / Oscuro** | Selector en la esquina inferior para alternar entre modo oscuro, claro o sincronizado con el sistema operativo. |

---

## 📊 Formato del Archivo Excel (`Albums List.xlsx`)

La aplicación detecta automáticamente las siguientes columnas en la primera fila de tu hoja:

* **RECORDS**: Título del álbum o disco.
* **CREATOR**: Nombre del artista o banda. Si hay varios artistas, sepáralos con punto y coma (ej: `Artista 1; Artista 2`).
* **TYPE**: Tipo de lanzamiento (`Album`, `EP`, etc.).
* **GENRES**: Géneros musicales (separados por coma o barra `/`).
* **LANGUAGE**: Idioma principal de las canciones.
* **YEAR**: Año de publicación (número entero de 4 dígitos).
* **MARK(/10)** o **MARK**: Calificación numérica sobre 10 (admite decimales con punto o coma).
* **WOKE?**: Indicador opcional (`Yeh` / `Nah`, `Yes` / `No`, `True` / `False`).

---

## 🛠️ Para Desarrolladores y Compilación

### Compilar ejecutables con PyInstaller

* **En Linux**:
  ```bash
  python -m PyInstaller --noconsole --onefile --clean --collect-all customtkinter readerUI.py
  ```
  El binario se generará en `dist/readerUI`.

* **En Windows (PowerShell)**:
  ```powershell
  py -m PyInstaller --noconsole --onefile --clean --collect-all customtkinter readerUI.py
  ```
  El binario se generará en `dist\readerUI.exe`.
