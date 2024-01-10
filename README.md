**Importador de notas PDF a Excel**

Este script en Python está diseñado para extraer información de archivos PDF y almacenarla en una hoja de cálculo de Excel.  Se requiere una carpeta con las notas en PDF, cuya ruta de acceso deberá actualizarse en el código. Las notas contienen: número de nota, referencia, fecha y el texto siempre comienza con "De mi mayor consideracion:" y finaliza en "Sin otro particular saluda atte.".

---

Resumen:

1. **Importación de bibliotecas necesarias:**
    - `os`: Proporciona una forma de interactuar con el sistema operativo.
    - `re`: Se utiliza para expresiones regulares.
    - `openpyxl`: Se utiliza para crear y manipular archivos de Excel.
    - `fitz` (PyMuPDF): Un enlace de Python para MuPDF, una biblioteca de renderizado de PDF.
    - `locale`: Se utiliza para establecer el idioma predeterminado para el análisis de fechas.

2. **Configuración del idioma predeterminado para el análisis de fechas en español.**

3. **Creación de un libro y hoja de Excel:**
    - Se crea un nuevo libro de Excel y se obtiene la hoja activa.

4. **Creación de encabezados de columnas en la hoja de Excel:**
    - Se establecen las cabeceras de las columnas para "Número de Nota", "Referencia", "Fecha" y "Texto".

5. **Búsqueda de archivos PDF en una carpeta específica:**
    - Se especifica una carpeta para buscar archivos PDF y se crea una lista de archivos PDF en esa carpeta.

6. **Expresión regular para fechas en formato específico en español.**
    - Se define una expresión regular para buscar fechas en un formato específico.

7. **Diccionario para mapear nombres de meses en español a números de mes.**

8. **Iteración a través de archivos PDF y extracción de información:**
    - Se itera a través de cada archivo PDF.
    - Se abre cada archivo PDF utilizando PyMuPDF.
    - Se inicializan variables para la información necesaria.
    - Se lee el contenido de cada página del PDF.
    - Se busca información específica en el texto, como el número de nota, referencia y fechas.
    - La información se almacena en la hoja de Excel.

9. **Guardado del libro de Excel:**
    - El libro de Excel se guarda en un archivo específico.