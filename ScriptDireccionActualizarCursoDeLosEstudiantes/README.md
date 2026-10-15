# Actualizar curso de estudiantes desde CSV

Esto es un script para Google Sheets que actualiza el curso de los
estudiantes a partir del listado oficial en un archivo CSV guardado en
Google Drive.

Compara los estudiantes que ya están en la hoja con los del CSV,
actualiza sus cursos, marca las incidencias y añade los estudiantes que
falten.

<br>

## Para poder usarlo

-   Abre el documento de Google Sheets donde quieras actualizar los
    estudiantes.
-   Ve al menú **Extensiones** y selecciona **Apps Script**.
    -   Se abrirá una pestaña nueva.
-   Sustituye el código del editor por el contenido de
    `Script_actualizar_curso.js`.
-   Cambia el nombre de *Proyecto sin título* por algo más descriptivo,
    por ejemplo *Actualizar cursos de estudiantes*.
-   Guarda los cambios.
-   Vuelve a la hoja de cálculo y recárgala para que aparezca el menú
    **Estudiantes**.
-   En el menú **Estudiantes**, selecciona **Actualizar desde CSV**.
-   La primera vez, Google te pedirá permisos para acceder al archivo de
    Drive y modificar la hoja. Autorízalos para poder ejecutar el
    script.

<br>

## Cómo se usa

1.  **Prepara el CSV.** Sube a Google Drive el archivo oficial
    actualizado con el listado de estudiantes. (Da a compartir y copia el enlace)
2.  En la hoja de cálculo, abre **Estudiantes → Actualizar desde CSV**.
3.  Comprueba el nombre de la pestaña que aparece en el aviso. El script
    trabaja sobre la **hoja activa**, así que selecciona antes la
    pestaña correcta.
4.  Pega el enlace para compartir del archivo CSV. También se puede
    indicar el ID o el nombre del archivo.
5.  Si hay estudiantes con los mismos apellidos pero nombres distintos,
    el script puede pedir confirmación manual para decidir si se trata
    de la misma persona.
7.  Al terminar, aparecerá un resumen con los totales del proceso.

<br>
<br>


## Formato que debe tener el CSV

Esto es el archivo obtenido en la plataforma.
El archivo debe ser un CSV **sin fila de encabezado**, con una línea por
estudiante y tres campos separados por comas. El ejemplo esperado es:

``` csv
"18","CARPINTERO , MARTA","ESO1A"
```

Los campos son:

1.  Número de orden: solo dígitos.
2.  Apellidos y nombre, separados por la primera coma.
3.  Curso en el formato que exporta el sistema del centro.

<br>

## Qué hace con los estudiantes

-   **Coincidencia exacta:** compara apellidos y nombre normalizados,
    ignorando mayúsculas, tildes, comas y espacios repetidos.
-   **Coincidencia parcial de nombre:** si está activada, acepta que el
    nombre de la hoja sea el principio del nombre del CSV cuando los
    apellidos coinciden y la coincidencia es única y no ambigua.
-   **Confirmación manual:** si no hay coincidencia exacta ni parcial,
    puede preguntar si una entrada del CSV con los mismos apellidos
    corresponde al estudiante de la hoja. Se puede elegir actualizar
    nombre y curso, actualizar solo el curso o indicar que no es la
    misma persona.
-   **Estudiantes que ya están en la hoja:** si se encuentran en el CSV,
    se actualiza siempre el curso usando el dato del CSV.
-   **Estudiantes nuevos:** si aparecen en el CSV y no en la hoja, se
    añaden al final del listado, en mayúsculas y conservando tildes y ñ.
-   **No encontrados:** si un estudiante de la hoja no aparece en el
    CSV, su curso se deja en blanco y se marca la incidencia en la
    columna de información adicional.
-   **Duplicados:** si el nombre normalizado aparece repetido en la hoja
    o en el CSV, las filas afectadas se marcan como duplicadas y el
    curso queda en blanco.

<br>
<br>

## Configuración

Si necesitas adaptar el script a otra hoja o a otros cursos, busca el
bloque **CONFIGURACIÓN** al principio del código.

Estas son las opciones principales:

-   `HEADER_ROWS = 2`: número de filas de encabezado. Por defecto, el
    primer estudiante está en la fila 3.
-   `COL_APELLIDO1 = 'A'`: columna del primer apellido.
-   `COL_APELLIDO2 = 'B'`: columna del segundo apellido.
-   `COL_NOMBRE = 'C'`: columna del nombre.
-   `COL_CURSO = 'Q'`: columna donde se escribe el curso normalizado.
-   `COL_CURSO_EXTRA = 'R'`: columna para el resto de información del
    curso y las incidencias.
-   `CURSOS_VALIDOS`: valores permitidos para la columna del curso
    (`1ESO`, `2ESO`, `3ESO`, `4ESO`, `1BTO`, `2BTO`, `AE` y `FPB`).
-   `COINCIDIR_POR_NOMBRE_PARCIAL = true`: activa la búsqueda por nombre
    parcial.
-   `CONFIRMAR_MISMOS_APELLIDOS = true`: activa las preguntas de
    confirmación manual para posibles coincidencias con los mismos
    apellidos.
-   `UMBRAL_CONTROL_COINCIDENCIAS = 150`: si, después de revisar 150
    estudiantes, no hay ninguna coincidencia, pregunta si se desea
    continuar o abortar.
-   `AE_TIENE_PRIORIDAD = true`: si aparece `AE` junto con un nivel, el
    curso resultante será `AE`.
-   `REGLAS_NIVEL`: define cómo se convierten los prefijos `ESO`, `BTO`
    y `FPB`.

Si cambias las columnas o las reglas de cursos, comprueba también que la
estructura de la hoja y la validación de datos de la columna del curso
coincidan con esa configuración.

<br>

## Avisos y resumen final

El resumen muestra el número de estudiantes del CSV, los cursos
actualizados, las coincidencias parciales, las confirmaciones manuales,
los nuevos añadidos, los no encontrados, los duplicados, las líneas
ignoradas, los cursos no reconocidos y los estudiantes cuyo curso estaba
vacío. No muestra nombres de estudiantes.

Si el script detecta un problema con el archivo, una estructura
incorrecta o un posible error de configuración, puede detenerse o pedir
confirmación antes de escribir datos. Si se aborta en esas
comprobaciones, no se modifica la hoja.
