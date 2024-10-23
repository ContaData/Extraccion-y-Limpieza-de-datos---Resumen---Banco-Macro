Extracción y Procesamiento de Resúmenes Bancarios del Banco Macro Modelo 1

Este proyecto es un script en Python diseñado para procesar automáticamente archivos PDF de resúmenes bancarios y convertirlos en archivos Excel, generando hojas de cálculo por cuenta bancaria y filtrando las transacciones relevantes.

Características

Procesamiento masivo de PDFs: El script procesa todos los archivos PDF dentro de una carpeta dada.
Extracción de datos clave: Extrae la fecha, descripción, referencia, débitos, créditos y saldos de cada transacción bancaria.
Filtrado de información: Filtra los movimientos financieros importantes (incluyendo saldos) para su exportación.
Generación de archivos Excel: Para cada resumen bancario, se genera un archivo Excel con una hoja por cada cuenta bancaria identificada en el PDF.

Manejo de errores: Si ocurre un error durante el procesamiento de un archivo PDF, se captura y se informa.

Requisitos
El script requiere Python 3.x y las siguientes bibliotecas de Python:
 - pandas: Para manipulación y análisis de datos.
 - tabula-py: Para extraer tablas de archivos PDF.
 - xlsxwriter: Para generar archivos Excel.
Puedes instalar las dependencias con el siguiente comando:

bash
pip install pandas tabula-py xlsxwriter

Además, necesitarás tener instalada Java, ya que tabula-py depende de ella para procesar los PDFs.

Configuración y Uso
Configura las rutas:
 - Define la ruta de la carpeta que contiene los archivos PDF a procesar (carpeta_pdf).
 - Define la carpeta de destino donde se guardarán los archivos Excel generados (carpeta_guardar).

Ejecuta el script:

Una vez que hayas configurado las rutas, puedes ejecutar el script principal. 
Este recorrerá la carpeta de PDFs, procesará cada archivo, y generará un archivo Excel correspondiente.

bash
python procesar_resumenes.py

Estructura de las hojas de cálculo:

Cada archivo Excel generado contendrá una o más hojas, donde cada hoja representa una cuenta bancaria encontrada en el resumen.
Las hojas de Excel tendrán columnas que representan la fecha, descripción, referencia, débitos, créditos y saldos de las transacciones.
Cómo funciona el script

Lectura de PDFs: El script utiliza tabula-py para extraer tablas de cada PDF, separando las transacciones bancarias.
Procesamiento de datos: Se renombra y organiza la información en columnas, y se extrae el número de cuenta desde la columna "Descripción".
Generación del Excel: Los datos se guardan en un archivo Excel, creando una hoja por cada cuenta bancaria.

Manejo de errores: Si un archivo PDF no se puede procesar correctamente, el script captura la excepción e imprime un mensaje de error, indicando el archivo problemático.

Ejemplo de Salida
Al ejecutar el script, generará archivos Excel con nombres que siguen el formato:


Copiar código
[Denominación] Periodo del Extracto [desde_fecha] al [hasta_fecha].xlsx
Por ejemplo:
Banco XYZ Periodo del Extracto 2023-01-01 al 2023-12-31.xlsx

Manejo de Errores
Si ocurre un error durante el procesamiento de un archivo PDF, el script continuará con los siguientes archivos y te mostrará un mensaje en la consola indicando qué archivo causó el error y una descripción del mismo.

Contribuir
Si deseas contribuir a este proyecto, puedes hacer un fork del repositorio, crear una nueva rama para tus cambios y luego hacer un pull request.

Realiza un fork del proyecto.
 - Crea una nueva rama: git checkout -b feature-nueva-funcionalidad.
 - Realiza tus cambios y haz commit: git commit -m 'Agrega nueva funcionalidad'.
 - Envía los cambios a tu fork: git push origin feature-nueva-funcionalidad.
 - Crea un pull request en GitHub.
