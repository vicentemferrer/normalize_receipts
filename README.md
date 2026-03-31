# 📄 APR Receipt Normalizer

Herramienta automatizada para procesar, validar y renombrar boletas de Agua Potable Rural (APR) en formato PDF basándose en registros de Excel.

## 🚀 Funcionalidades

- **Extracción de Metadatos**: Localiza automáticamente el número de medidor y el nombre del cliente dentro del PDF.
- **Limpieza de Datos**: Elimina automáticamente honoríficos (SR, SRA, DON), símbolos de moneda y normaliza espacios o caracteres especiales.
- **Sincronización con Excel**: Cruza la información del PDF con una base de datos maestra para asegurar la integridad de los datos.
- **Estandarización de Archivos**: Renombra los archivos siguiendo el formato: `[Medidor]_[Nombre]_[Apellido].pdf`.

## 🛠️ Requisitos Técnico

El script requiere Python 3.x y las siguientes librerías:

- `pandas` & `openpyxl`: Para el manejo de bases de datos Excel.
- `pdfplumber`: Para la lectura precisa de documentos PDF.
- `unidecode`: Para la normalización de caracteres.
- `python-dotenv`: Para la gestión segura de rutas y configuración.

## ⚙️ Configuración

1.  Instala las dependencias: `pip install pandas pdfplumber unidecode python-dotenv openpyxl`
2.  Crea un archivo `.env` en la raíz con el siguiente formato:
    ```env
    EXCEL_FILE=tu_archivo.xlsx
    SHEET_NAME=Hoja1
    PDF_FOLDER=./boletas_entrada
    COL_MEDIDOR=NumeroMedidor
    COL_NOMBRE=Nombre
    COL_APELLIDO1=ApellidoPaterno
    COL_APELLIDO2=ApellidoMaterno
    ```

## 📜 Licencia

Copyright © 2026. Todos los derechos reservados. El uso, copia o distribución de este software está estrictamente prohibido sin autorización expresa del autor.
