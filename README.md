# 📌 Extracción de Datos desde Documentos Word

Este proyecto permite extraer información clave de documentos Word (`.docx`) exportados desde **ArcGIS** y almacenarla en un archivo Excel.

## 🚀 Características

- Extrae datos de **párrafos y tablas** dentro de los documentos Word.
- Utiliza **expresiones regulares** para mejorar la precisión de la extracción.
- Convierte fechas a un formato estándar (`DD/MM/YYYY HH:MM`).
- Permite **procesamiento en paralelo** para agilizar la extracción de múltiples archivos.

## 📂 Estructura del Proyecto

```
📁 word-to-excel-arcgis
│── 📁 Archivos  # Carpeta donde deben estar los archivos .docx
│── 📄 extraer_word.py  # Código principal de extracción
│── 📄 requirements.txt  # Dependencias del proyecto
│── 📄 LICENSE  # Licencia del proyecto
│── 📄 README.md  # Documentación del proyecto
```

## 🛠️ Instalación y Uso

### 📥 Descarga del Proyecto

Puedes descargar el proyecto directamente desde el siguiente enlace:
[Descargar word-to-excel-arcgis](https://github.com/dashmonroy/word-to-excel-arcgis/archive/refs/heads/main.zip)

### 🔧 Instalación

1. **Clona este repositorio** o descárgalo manualmente:
   ```bash
   git clone https://github.com/dashmonroy/word-to-excel-arcgis.git
   cd word-to-excel-arcgis
   ```
2. **Instala las dependencias necesarias**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Ejecuta el script para extraer los datos**:
   ```bash
   python extraer_word.py
   ```
4. **Los datos extraídos se guardarán en un archivo Excel en la misma carpeta.**

## ℹ️ ¿Qué hace este proyecto?

Este proyecto está diseñado para **automatizar la extracción de datos** desde documentos Word generados en ArcGIS, organizándolos en un archivo Excel para su posterior análisis. Es especialmente útil para proyectos de **geoprocesamiento, reportes técnicos y consolidación de información estructurada**.

## 📜 Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo `LICENSE` para más detalles.
