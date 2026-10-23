# ClasseurRouges
Programme qui stocke de manière locale les notes des élèves sur leurs 3 ans de cycle et génère des rapports par classe ou par élève.

Aplicación local (100 % offline) para almacenar el seguimiento académico de los alumnos durante sus 3 años de Cycle d'Orientation (9e, 10e, 11e) y generar vistas e informes Excel por clase/cohorte o por alumno.

## Características principales

- **Importación multiformato con guardado automático:**
  - **Archivos XML (`Notes*.xml` y `Absences*.xml`):** importa automáticamente las notas por trimestre (`T1`, `T2`, `T3`), calcula las medias anuales y ponderadas (respetando las notas con `Compte="0"` entre paréntesis sin incluirlas en la media), y registra las ausencias, retrasos y dispensas de Educación Física. Los datos se guardan automáticamente en la base de datos al importarlos.
  - **Listado de profesores (`*MaitresParGroupe.xls`, `.xlsx`, `.csv`):** vincula los profesores de cada asignatura y el *Maître de classe* a cada grupo y año escolar. Puede importarse junto con los XML o con posterioridad para actualizar los boletines ya guardados.
  - **Boletines en PDF (`*.pdf`):** mantiene la importación tradicional de boletines escolares en PDF con revisión previa antes de guardar.
- **Nombres de profesores en fichas individuales:**
  - El nombre del profesor aparece debajo de cada asignatura únicamente en las **Fichas Individuales de Alumnos** (tanto en la aplicación como en las pestañas individuales del informe Excel exportado), manteniendo limpia la **Vista Clase / Cohorte**.
- **Filtro por Año Escolar y Clase:**
  - Permite filtrar por **Año escolar** (`Année scolaire`) y por **Clase** tanto en la vista de cohorte como en las fichas individuales de los alumnos.
  - Al seleccionar un año escolar, se muestran las clases correspondientes a ese año y los alumnos que cursaron en dicha clase durante ese año escolar, conservando el historial completo de 9e, 10e y 11e de cada alumno.
- **Exportación a Excel (`.xlsx`):**
  - Genera libros Excel completos con hojas resumen por cohorte/clase y pestañas individuales formateadas para cada alumno (incluyendo asignaturas, profesores, medias por grupo, ausencias y observaciones).

## Estructura del proyecto

- `Lancer_Classeurs_Rouges.bat`: lanzador de la aplicación (abre el servidor local en PowerShell y la interfaz en Microsoft Edge en modo aplicación).
- `src/`: código fuente de la aplicación:
  - `Server.ps1`: servidor HTTP local y endpoints de la API.
  - `Engine.ps1`: lógica de negocio, importación de XML/Excel/CSV/PDF, mapeo de códigos de asignaturas del CO, cálculo de medias y base de datos JSON.
  - `ui.html`: interfaz gráfica de usuario.
  - `NativePdfExtractor.cs`: extractor nativo en C# para boletines en formato PDF.
  - `XlsxWriter.cs`: generador nativo en C# de archivos Excel `.xlsx` OpenXML sin dependencias externas.
- `data/`: base de datos local (`database.json`) y copias de seguridad automáticas (`backups/`) — *ignorado por Git*.
- `Bulletins_Exemples/`, `Rapports_Excel/`: carpetas para documentos de ejemplo e informes Excel generados — *ignoradas por Git*.

## Requisitos y uso

1. Ejecutar `Lancer_Classeurs_Rouges.bat` en Windows (requiere PowerShell 5.1+ y Microsoft Edge; si se importan archivos `.xls`/`.xlsx` directamente, utiliza Microsoft Excel instalado o acepta archivos `.csv` sin necesidad de Excel).
2. En la pestaña **Importation Documents**, arrastrar o seleccionar los archivos XML de notas/ausencias, el archivo Excel/CSV de profesores y/o los boletines PDF.
3. Consultar los datos en **Vue Classe / Cohorte** o **Fiches Individuelles Élèves** filtrando por **Año escolar** y **Clase**, o pulsar **Exporter Excel (.xlsx)** para generar el informe completo.
