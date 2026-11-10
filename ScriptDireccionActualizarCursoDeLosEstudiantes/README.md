# Actualizar datos de estudiantes desde CSV

Este script para Google Sheets actualiza la información de los estudiantes a partir de dos listados oficiales exportados desde la plataforma del centro.

El proceso se realiza en **dos pasos independientes**, cada uno con su propio CSV:

1. **Actualizar cursos:** sincroniza los cursos, marca incidencias y añade los estudiantes que falten.
2. **Completar datos personales:** actualiza la fecha de nacimiento, el país y la nacionalidad de los estudiantes que encuentre.

Los dos archivos deben obtenerse previamente de la plataforma y subirse a Google Drive.

<br>

## Para poder usarlo *(primera vez)*

- Abre el documento de Google Sheets donde quieras actualizar los estudiantes.
- Ve al menú **Extensiones** y selecciona **Apps Script**.
  - Se abrirá una pestaña nueva.
- Sustituye el código del editor por el contenido de `Script_actualizar_curso.js`.
- Cambia el nombre de *Proyecto sin título* por algo más descriptivo, por ejemplo *Actualizar datos de estudiantes*.
- Guarda los cambios.
- Vuelve a la hoja de cálculo y recárgala para que aparezca el menú **Estudiantes**.
- ⚠️ La primera vez que ejecutes el script, Google te pedirá permisos para acceder a los archivos de Drive y modificar la hoja. Autorízalos para poder utilizarlo.

<br>

## 1. Obtener los listados y subirlos a Google Drive

Antes de ejecutar el script, obtén **los dos listados** en la plataforma del centro:

1. Ve a **Central de listados / Filiación / Listado de matrículas**.
2. Genera estos dos archivos:

   - **Filiación** (contiene nombre, apellidos y curso).
     - ⚠️ En **Configuración de impresión**, selecciona **CSV con cabecera y pie**.
   - **Estandar / ALXE-25CL** (contiene los datos para completar la fecha de nacimiento, el país y la nacionalidad).
     - ⚠️ En **Configuración de impresión**, selecciona **CSV con cabecera y pie**.

3. Sube **ambos CSV** a Google Drive. Puedes guardarlos en la misma carpeta.
4. Comprueba que los archivos se han subido correctamente. Después, al ejecutar cada paso, tendrás que pegar el enlace para compartir del CSV correspondiente. (**Archivo** → **Compartir** → **Copiar enlace**)

💡Es recomendable conservar nombres de archivo que permitan distinguir claramente el listado de **Filiación** del de **Estandar / ALXE-25CL**.

<br>

## 2. Ejecutar el proceso desde Google Sheets

En la hoja de cálculo, abre el menú **Estudiantes**. Encontrarás dos opciones:

- **Paso 1: Actualizar cursos desde CSV**
- **Paso 2: Completar fecha de nacimiento, país y nacionalidad**

Ejecuta cada opción con el archivo correspondiente.

### Paso 1: Actualizar cursos desde CSV

1. Selecciona **Estudiantes → Paso 1: Actualizar cursos desde CSV**.
2. Comprueba el nombre de la pestaña que aparece en el aviso. El script trabaja sobre la **hoja activa**, así que selecciona antes la pestaña correcta.
3. Cuando lo solicite, pega el enlace para compartir del CSV de **Filiación** que has subido a Drive. También puedes indicar el ID o el nombre del archivo.
4. Si hay estudiantes con los mismos apellidos pero nombres distintos, el script puede pedir confirmación manual para decidir si se trata de la misma persona.
5. Al terminar, aparecerá un resumen con los totales del proceso.

Este paso actualiza el curso de los estudiantes, añade los que aparecen en el listado oficial pero todavía no están en la hoja y señala las incidencias (No encontrado, Duplicado, etc).

### Paso 2: Completar fecha de nacimiento, país y nacionalidad

1. Selecciona **Estudiantes → Paso 2: Completar fecha de nacimiento, país y nacionalidad**.
2. Comprueba de nuevo el nombre de la pestaña indicada. El proceso se ejecuta sobre la **hoja activa**.
3. Cuando lo solicite, pega el enlace para compartir del CSV **Estandar / ALXE-25CL** que has subido a Drive. También puedes indicar el ID o el nombre del archivo.
4. Al terminar, aparecerá un resumen con los totales del proceso.

Este paso busca coincidencias exactas y únicas por nombre y apellidos. Cuando encuentra al estudiante, completa o actualiza su fecha de nacimiento, país y nacionalidad. No añade estudiantes ni modifica nombres o cursos.

<br>

## Información extra del script y su funcionamiento
### Formato de los CSV

Los dos archivos deben ser los CSV exportados desde la plataforma con la opción **CSV con cabecera y pie**. El script comprueba que cada archivo corresponde al paso seleccionado y que contiene las columnas esperadas.

#### CSV del paso 1: Filiación

Se utiliza para actualizar los cursos. Debe contener las columnas:

```csv
"NIF/Pasaporte","Apellidos","Nombre","Sección"
"12345678Z","CARPINTERO","ANA MARÍA","ESO1A"
```

Se utilizan los campos **Apellidos**, **Nombre** y **Sección**. El campo **NIF/Pasaporte** se ignora.

#### CSV del paso 2: Estandar / ALXE-25CL

Se utiliza para completar los datos personales. Debe contener las columnas:

```csv
"Nº","Alumno/a","NIF/Pasaporte","Fecha nacimiento","Localidad","Provincia","País","NACIONALIDAD"
"18","ANA MARÍA CARPINTERO","12345678Z","03/05/2009","MADRID","MADRID","ESPAÑA","RUMANA"
```

Se utilizan los campos **Alumno/a**, **Fecha nacimiento**, **País** y **NACIONALIDAD**. Los demás campos se ignoran.

**Importante:** no intercambies los archivos. El CSV de Filiación corresponde al paso 1 y el CSV Estandar / ALXE-25CL corresponde al paso 2. Si seleccionas el archivo equivocado, el script detectará que la estructura no es la esperada y no modificará la hoja.

<br>

### Qué hace con los estudiantes

#### Paso 1: cursos

- **Coincidencia exacta:** compara apellidos y nombre normalizados, ignorando mayúsculas, tildes y espacios repetidos.
- **Posibles coincidencias con los mismos apellidos:** si no hay coincidencia exacta, puede pedir confirmación manual. Se puede elegir actualizar el nombre y el curso, actualizar solo el curso (con una segunda confirmación) o indicar que no es la misma persona.
- **Estudiantes que ya están en la hoja:** si se encuentran en el CSV, se actualiza siempre el curso usando el dato del CSV.
- **Estudiantes nuevos:** si aparecen en el CSV y no en la hoja, se añaden al final del listado, en mayúsculas y conservando tildes y ñ.
- **No encontrados:** si un estudiante de la hoja no aparece en el CSV, su curso se deja en blanco y se marca la incidencia en la columna de información adicional.
- **Duplicados:** si el nombre normalizado aparece repetido en la hoja o en el CSV, las filas afectadas se marcan como duplicadas y el curso queda en blanco.
- **Cursos no reconocidos o vacíos:** el curso queda en blanco y se indica la incidencia en la columna de información adicional.

#### Paso 2: datos personales

- **Coincidencia exacta y única:** compara el nombre completo del CSV con el nombre y los apellidos de la hoja, sin distinguir mayúsculas, tildes ni signos de puntuación.
- **Datos encontrados:** actualiza la fecha de nacimiento, el país y la nacionalidad.
- **Sin coincidencia:** no modifica al estudiante.
- **Nombres duplicados o ambiguos:** no modifica esos registros y los cuenta en el resumen.
- **Campos vacíos en el CSV:** se conserva el valor que ya hubiera en la hoja.

<br>

### Configuraciones que permite el script

Si necesitas adaptar el script a otra hoja o a otros cursos, busca el bloque **CONFIGURACIÓN** al principio del código.

Estas son algunas de las opciones principales:

- `HEADER_ROWS = 2`: número de filas de encabezado. Por defecto, el primer estudiante está en la fila 3.
- `COL_APELLIDO1 = 'A'`: columna del primer apellido.
- `COL_APELLIDO2 = 'B'`: columna del segundo apellido.
- `COL_NOMBRE = 'C'`: columna del nombre.
- `COL_CURSO = 'Q'`: columna donde se escribe el curso normalizado.
- `COL_CURSO_EXTRA = 'R'`: columna para el resto de información del curso y las incidencias.
- `COL_FECHA_NACIMIENTO = 'D'`: columna de la fecha de nacimiento.
- `COL_PAIS = 'F'`: columna del país.
- `COL_NACIONALIDAD = 'G'`: columna de la nacionalidad.
- `CURSOS_VALIDOS`: valores permitidos para la columna del curso (`1ESO`, `2ESO`, `3ESO`, `4ESO`, `1BTO`, `2BTO`, `AE` y `FPB`).
- `CONFIRMAR_MISMOS_APELLIDOS = true`: activa las preguntas de confirmación manual del paso 1 para posibles coincidencias con los mismos apellidos.
- `UMBRAL_CONTROL_COINCIDENCIAS = 150`: si, después de revisar 150 estudiantes en el paso 1, no hay ninguna coincidencia, pregunta si se desea continuar o abortar.
- `AE_TIENE_PRIORIDAD = true`: si aparece `AE` junto con un nivel, el curso resultante será `AE`.
- `REGLAS_NIVEL`: define cómo se convierten los prefijos `ESO`, `BTO` y `FPB`.
- `CONVERTIR_FECHA = true`: intenta guardar la fecha de nacimiento como una fecha real de la hoja.
- `SOBRESCRIBIR_CON_VACIOS = false`: conserva los datos existentes cuando el CSV del paso 2 trae vacío un campo personal.

Si cambias las columnas o las reglas de cursos, comprueba también que la estructura de la hoja y la validación de datos de la columna del curso coincidan con esa configuración.

<br>

## Avisos y resumen final

Cada paso muestra su propio resumen:

- **Paso 1:** incluye totales de estudiantes del CSV, actualizados, nuevos, no encontrados, duplicados, líneas ignoradas, cursos no reconocidos y estudiantes cuyo curso estaba vacío. También informa de las confirmaciones manuales y de los posibles guardados parciales.
- **Paso 2:** incluye los totales de estudiantes procesados, actualizados, sin coincidencia y ambiguos, además de las líneas ignoradas.

Los resúmenes no muestran nombres de estudiantes.

Si el script detecta un problema con el archivo, una estructura incorrecta o un posible error de configuración, puede detenerse o pedir confirmación antes de escribir datos. Si se aborta en esas comprobaciones, no se modifica la hoja.

<br/><br/><br/>
## ⌛ Control del tiempo de ejecución (Paso 1)

El **Paso 1** puede necesitar mostrar preguntas para confirmar si dos registros con los mismos apellidos corresponden a la misma persona. Como Google Apps Script limita cada ejecución a **6 minutos**, el programa incorpora un aviso preventivo al cumplirse **5 minutos** desde el inicio del paso, cuando va a mostrar la siguiente pregunta.

En ese momento aparece un cuadro con dos opciones:

- **Sí — Guardar los avances y detener (recomendado):** el script guarda los estudiantes que ya ha podido emparejar, tanto por coincidencia exacta como mediante las confirmaciones realizadas, y termina esa ejecución. No marca todavía como «no encontrados» a los estudiantes pendientes ni añade los nuevos: esas operaciones se reservan para cuando se complete el proceso.
- **No — Seguir preguntando:** el script continúa con las preguntas. Sin embargo, si la ejecución supera el límite de 6 minutos de Google Apps Script, Google puede interrumpirla y se perderán las respuestas de esa ejecución que aún no se hayan guardado.

### Cómo continuar después de un guardado parcial

Si eliges **Sí**, vuelve a ejecutar **Estudiantes → Paso 1: Actualizar cursos desde CSV** y selecciona el mismo CSV. El proceso continuará con los registros pendientes:

- Los estudiantes ya emparejados y guardados coincidirán exactamente al repetir el proceso, por lo que no volverán a pedir confirmación.
- Los casos en los que respondiste **No** o **Cancelar** se volverán a preguntar.
- Los estudiantes que aún no se habían resuelto se revisarán en esta nueva ejecución.
- Los estudiantes nuevos y los que no se encuentren se gestionarán cuando el paso llegue a completarse.

El aviso de los 5 minutos es, por tanto, una medida de seguridad para poder trabajar con listados grandes sin arriesgar todas las confirmaciones realizadas. Si aparece, lo recomendable es elegir **Sí** y reanudar el Paso 1.