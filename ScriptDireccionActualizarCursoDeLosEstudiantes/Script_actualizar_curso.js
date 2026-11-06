/**
 * ============================================================================
 * INSTRUCCIONES PARA LA IA (leer antes de modificar este script)
 * ============================================================================
 * PROPÓSITO
 *   Script de Google Apps Script ligado a un Google Spreadsheet con el listado
 *   de estudiantes de un centro educativo. Se usa en DOS PASOS, cada uno con su
 *   propio CSV subido a Google Drive y su propia opción del menú:
 *     PASO 1 (actualizarCursos -> procesarCursos_): sincroniza la hoja con el
 *       listado oficial de cursos (CSV de cursos).
 *     PASO 2 (actualizarDatosPersonales -> procesarDatosPersonales_): rellena la
 *       fecha de nacimiento, el país y la nacionalidad (CSV de datos personales).
 *   Los dos CSV vienen de listados distintos y NO se comparan igual con la hoja
 *   (ver REGLAS DE NEGOCIO de cada paso).
 *
 * ESTRUCTURA DE LA HOJA (configurable en la sección CONFIGURACIÓN)
 *   - Columna A: apellido 1 | Columna B: apellido 2 | Columna C: nombre
 *   - Columna D: fecha de nacimiento | Columna F: país | Columna G: nacionalidad
 *     (configurable, ver COL_FECHA_NACIMIENTO, COL_PAIS y COL_NACIONALIDAD).
 *     Solo las escribe el PASO 2. La columna E y cualquier otra no indicada en
 *     esta lista NO se tocan.
 *   - Columna Q: curso del estudiante, con un valor de CURSOS_VALIDOS
 *     (configurable, ver COL_CURSO). Solo la escribe el PASO 1.
 *   - Columna R: resto de la información del curso del CSV (grupo, "(R)",
 *     etc.) (configurable, ver COL_CURSO_EXTRA). Solo la escribe el PASO 1.
 *   - Las primeras HEADER_ROWS filas son encabezado (el primer estudiante está
 *     en la fila 3 por defecto).
 *   - Se trabaja siempre sobre la HOJA ACTIVA; antes de empezar se muestra un
 *     aviso con el nombre de la hoja y hay que confirmar.
 *
 * FORMATO DEL CSV DEL PASO 1 (cursos; ver FORMATO_CSV_CURSOS)
 *   CON fila de encabezado, separador coma, campos entre comillas:
 *     "NIF/Pasaporte","Apellidos","Nombre","Sección"
 *     "12345678Z","ORBULESCU","ANA MARÍA","ESO1A"
 *   - Encabezado: se salta (CSV_TIENE_CABECERA), pero antes se comprueba que las
 *     columnas usadas se llaman como se espera (cabecerasEsperadas), para no
 *     leer una columna equivocada si cambia el orden.
 *   - Campo 0: NIF/Pasaporte -> SE IGNORA.
 *   - Campo 1 (CSV_CURSOS_IDX_APELLIDOS): apellidos.
 *   - Campo 2 (CSV_CURSOS_IDX_NOMBRE): nombre.
 *   - Campo 3 (CSV_CURSOS_IDX_CURSO): sección = curso en formato libre del CSV
 *     (ver TRANSFORMACIÓN DEL CURSO).
 *
 * FORMATO DEL CSV DEL PASO 2 (datos personales; ver FORMATO_CSV_DATOS)
 *   CON fila de encabezado, separador coma, campos entre comillas:
 *     "Nº","Alumno/a","NIF/Pasaporte","Fecha nacimiento","Localidad","Provincia","País","NACIONALIDAD"
 *     "18","ANA MARÍA ORBULESCU","12345678Z","03/05/2009","MADRID","MADRID","ESPAÑA","RUMANA"
 *   - Campo 1 (CSV_DATOS_IDX_NOMBRE_COMPLETO): "Alumno/a" = NOMBRE seguido de
 *     APELLIDOS en un solo campo, sin coma (el orden contrario al de la hoja).
 *   - Campo 3: fecha de nacimiento -> columna D.
 *   - Campo 6: país -> columna F.
 *   - Campo 7: nacionalidad -> columna G.
 *   - Campos 0 (Nº), 2 (NIF/Pasaporte), 4 (Localidad) y 5 (Provincia): SE IGNORAN.
 *
 * CODIFICACIÓN Y VALIDACIÓN (comunes a los dos CSV)
 *   - Codificación (ver leerCsv_): se detecta automáticamente. UTF-16 LE/BE
 *     (con o sin BOM; es la que genera el sistema de gestión del centro),
 *     UTF-8 (con o sin BOM) y, si tiene caracteres inválidos, windows-1252.
 *     Un CSV UTF-16 leído como UTF-8 se ve como texto con caracteres nulos
 *     ("�") entre cada letra; no cambiar esta detección sin revisarlo.
 *   - VALIDACIÓN DE ESTRUCTURA (ver prepararFilasCsv_, validarEstructuraCsv_ y la
 *     función esLineaValida de cada formato): una línea es válida si tiene campos
 *     suficientes para el índice mayor usado y los campos de nombre no están
 *     vacíos (paso 1: apellidos y nombre; paso 2: Alumno/a). El encabezado no
 *     cuenta. Si menos del PORCENTAJE_MIN_LINEAS_VALIDAS (90%) de las líneas no
 *     vacías son válidas se lanza un error que explica el formato esperado de
 *     ESE paso, con ejemplos de líneas incorrectas, y no se modifica nada. Las
 *     pocas líneas inválidas que queden se ignoran y se cuentan en el resumen.
 *     Si se usa el CSV de un paso en el otro, falla la comprobación del
 *     encabezado y se muestra ese mismo error.
 *
 * FLUJO (SIN HTML: no se usa HtmlService ni google.script.run, para evitar
 * problemas con varias cuentas de Google abiertas en el navegador)
 *   1. Menú "Estudiantes" con dos opciones (onOpen): "Paso 1: Actualizar cursos
 *      desde CSV" y "Paso 2: Completar fecha de nacimiento, país y nacionalidad".
 *      Ambas ejecutan ejecutarPaso_ con su función de proceso.
 *   2. Aviso de confirmación con el nombre de la hoja activa.
 *   3. Cuadro de texto que pide el nombre, la URL o el ID del CSV en Google Drive
 *      (ver buscarArchivoCsv_).
 *   3b. comprobarEsCsv_ verifica (solo con nombre y tipo MIME, sin leer el
 *       contenido) que el archivo es un CSV. Si no lo es, se muestra un aviso y
 *       se detiene el proceso sin leer el archivo ni modificar la hoja.
 *   4. procesarCursos_ (paso 1) o procesarDatosPersonales_ (paso 2) compara y
 *      actualiza la hoja y devuelve un resumen.
 *   5. El resumen (solo totales, sin detalle de estudiantes) se muestra en un
 *      cuadro de alerta y se guarda en el registro de ejecución con Logger.log.
 *      Si había varios archivos con el mismo nombre, el aviso se antepone.
 *
 * REGLAS DE NEGOCIO DEL PASO 1 (cursos)
 *   - Comparación por nombre completo NORMALIZADO: concatenar
 *     apellido1 + apellido2 + nombre (hoja) frente a "apellidos nombre" (CSV),
 *     sin comas, minúsculas, sin tildes ni diacríticos (la ñ pasa a n) y con
 *     espacios múltiples colapsados a uno solo.
 *   - CONFIRMACIÓN MANUAL (CONFIRMAR_MISMOS_APELLIDOS; bloque 2b de
 *     procesarCursos_): para las filas de la hoja sin coincidencia exacta pero
 *     con una entrada del CSV con los MISMOS apellidos (normalizados) que no
 *     coincide ya con otra fila, se muestra un cuadro Sí/No/Cancelar (los
 *     botones de Apps Script no se pueden renombrar, por eso el texto explica
 *     cada uno):
 *        Sí       -> actualiza SOLO el nombre (columna C) y el curso; los
 *                    apellidos (columnas A y B) NO se modifican nunca
 *        No       -> actualiza solo el curso, pero ANTES muestra una segunda
 *                    ventana Sí/No de confirmación ("Se actualizará SOLO el
 *                    curso..."); si en ella se responde No, se vuelve a mostrar
 *                    la pregunta anterior
 *        Cancelar -> no es el mismo; se prueba el siguiente candidato y, si
 *                    no hay más, la fila queda como no encontrada
 *     Las respuestas solo se aplican en el paso 5 (escritura), así que no se
 *     modifica nada antes. Cubre, por ejemplo, erratas en el nombre
 *     (hoja "YALIANA ALTAGRACIA" frente a CSV "YALIANI ALTAGRACIA") y nombres
 *     incompletos en la hoja ("DULCE" frente a "DULCE ZOE"). NO hay
 *     coincidencia automática por nombre parcial: estos casos siempre se
 *     preguntan.
 *   - ICONO DEL AVISO: el cuadro empieza y sigue con ✅ si aparentemente solo
 *     falta parte del nombre en la hoja (el nombre de la hoja es el PRINCIPIO del
 *     nombre del CSV y el CSV tiene más palabras: "DULCE" frente a "DULCE ZOE"),
 *     y con ⚠️ en cualquier otro caso (hermanos con los mismos apellidos,
 *     erratas, otro orden de los nombres...). Con ⚠️ además se deja una línea
 *     en blanco más antes de "Hoja:" y antes de "¿Es el mismo estudiante?". Es solo informativo: no cambia lo que
 *     hacen los botones (ver nombreEsInicioDe_).
 *   - CONTADOR DE PREGUNTAS: el título de cada cuadro muestra "pregunta X/N".
 *     N es el nº de filas de la hoja sin coincidencia exacta que tienen algún
 *     candidato con los mismos apellidos, calculado antes de preguntar; es
 *     aproximado (una respuesta puede emparejar una entrada que era candidata de
 *     otra fila, y entonces se salta alguna pregunta; si se responde "Cancelar"
 *     y la fila tiene otro candidato, se repite el mismo número).
 *   - LÍMITE DE TIEMPO (MINUTOS_AVISO_LIMITE = 5): Apps Script corta cada
 *     ejecución a los 6 minutos y se perderían todas las respuestas, que solo se
 *     escriben al final. Si al ir a mostrar una pregunta han pasado 5 minutos
 *     desde que se lanzó el paso, se muestra en su lugar un aviso (Sí = guardar
 *     los avances y detener, recomendado; No = seguir preguntando bajo el propio
 *     riesgo). El GUARDADO PARCIAL escribe solo los estudiantes ya emparejados
 *     (exactos y confirmados) y NO marca no encontrados ni añade nuevos, porque
 *     las filas sin preguntar todavía no están resueltas. Al volver a lanzar el
 *     paso 1: los confirmados con "Sí" ya coinciden exactamente (no se
 *     preguntan), pero los respondidos con "No" o "Cancelar" se volverán a
 *     preguntar.
 *   - Estudiante de la hoja encontrado en el CSV -> se escribe SIEMPRE el curso
 *     del CSV (el CSV es la fuente más actualizada), sea igual o distinto al actual.
 *   - REGLA GENERAL DE INCIDENCIAS: la columna COL_CURSO tiene validación de
 *     datos en la hoja, así que SOLO puede contener un valor de CURSOS_VALIDOS
 *     o estar en blanco. Ante cualquier incidencia (no encontrado, duplicado,
 *     curso no reconocido...) COL_CURSO se deja EN BLANCO y la incidencia se
 *     escribe en COL_CURSO_EXTRA. No escribir nunca textos de aviso en COL_CURSO.
 *   - Estudiante de la hoja NO encontrado en el CSV -> curso en blanco y
 *     TEXTO_NO_ENCONTRADO en COL_CURSO_EXTRA.
 *   - Estudiante del CSV que no está en la hoja -> se añade como fila nueva
 *     justo después del último estudiante. Se escribe en MAYÚSCULAS conservando
 *     tildes y ñ. Reparto de apellidos (ver separarApellidos_): si son
 *     exactamente 2 palabras, la primera va a COL_APELLIDO1 y la segunda a
 *     COL_APELLIDO2; en cualquier otro caso (1 palabra, o 3 o más, por ejemplo
 *     apellidos compuestos como "DE LA CRUZ") TODOS los apellidos van en
 *     COL_APELLIDO1 y COL_APELLIDO2 queda vacía. Este reparto se aplica SOLO a
 *     los estudiantes nuevos: los apellidos de los estudiantes que ya están en
 *     la hoja no se reescriben nunca.
 *   - Duplicados: si un nombre normalizado aparece más de una vez en la hoja o
 *     en el CSV, todas las filas afectadas quedan con el curso en blanco y
 *     TEXTO_DUPLICADO en COL_CURSO_EXTRA (si el CSV tiene más repeticiones que
 *     la hoja, las que faltan se añaden como filas nuevas también marcadas) y
 *     se cuentan en el resumen.
 *   - TRANSFORMACIÓN DEL CURSO (ver parsearCurso_): el valor del CSV se divide
 *     en "curso" (columna COL_CURSO, siempre un valor de CURSOS_VALIDOS) y
 *     "resto" (columna COL_CURSO_EXTRA). El valor puede traer varios elementos
 *     separados por coma. Ejemplos (CSV -> curso | resto):
 *         ESO1A          -> 1ESO | A
 *         ESO1A (R)      -> 1ESO | A (R)
 *         ESO 4D (R)     -> 4ESO | D (R)
 *         BTO2C (R)      -> 2BTO | C (R)
 *         FPB1A (R)      -> FPB  | 1A (R)   (el número se conserva en el resto)
 *         AE, ESO3A      -> AE   | ESO3A
 *         ESO4C, AE      -> AE   | ESO4C
 *     Si aparece "AE" junto a un nivel, el curso es AE (AE_TIENE_PRIORIDAD).
 *     Si el curso viene VACÍO en el CSV, el curso queda en blanco y en
 *     COL_CURSO_EXTRA se escribe TEXTO_CURSO_VACIO; en el resumen se cuenta
 *     aparte ("Estudiantes con el curso vacío"), NO como curso no reconocido.
 *     Si el curso resultante no está en CURSOS_VALIDOS (o no se reconoce el
 *     valor), el curso queda en blanco, en COL_CURSO_EXTRA se escribe
 *     "TEXTO_CURSO_NO_RECONOCIDO: <valor original completo del CSV>" y se cuenta
 *     en el resumen.
 *     La columna del resto se reescribe entera para los estudiantes
 *     existentes (los valores anteriores se sobrescriben).
 *   - NO se cambia ningún estilo/formato/color de la hoja.
 *   - Se muestra un resumen final solo con totales (estudiantes en el CSV,
 *     actualizados, nuevos, no encontrados, duplicados, líneas ignoradas y
 *     cursos no reconocidos y estudiantes con el curso vacío). NO se listan
 *     nombres de estudiantes.
 *   - CONTROL DE COINCIDENCIAS: en procesarCursos_, si tras procesar
 *     UMBRAL_CONTROL_COINCIDENCIAS estudiantes de la hoja (150 por defecto)
 *     ninguno ha coincidido con el CSV, se muestra un aviso (Sí/No) para
 *     continuar o abortar. Se pregunta una sola vez y siempre antes de
 *     escribir nada en la hoja, así que abortar no modifica nada.
 *
 * REGLAS DE NEGOCIO DEL PASO 2 (datos personales)
 *   - Comparación ESTRICTA (sin búsquedas aproximadas ni preguntas): el campo
 *     "Alumno/a" del CSV se compara con NOMBRE + APELLIDO1 + APELLIDO2 de la hoja
 *     (en ese orden, columnas C + A + B), ambos normalizados con
 *     normalizarNombreEstricto_: sin tildes ni diacríticos (ñ -> n), minúsculas,
 *     cualquier carácter que no sea letra, número o espacio pasa a espacio, y
 *     espacios múltiples colapsados. (Es distinta de la normalización del paso 1,
 *     que NO se debe cambiar: allí la clave es apellido1 + apellido2 + nombre.)
 *   - Nombre idéntico y único -> se copian a la hoja la fecha de nacimiento
 *     (COL_FECHA_NACIMIENTO), el país (COL_PAIS) y la nacionalidad
 *     (COL_NACIONALIDAD): se rellenan si estaban vacías y se actualizan si ya
 *     tenían un valor.
 *   - Nombre no idéntico -> NO se hace nada. Los estudiantes de la hoja sin
 *     coincidencia no se tocan, y los del CSV que no están en la hoja se ignoran
 *     (este paso NUNCA añade filas ni cambia nombres, cursos ni incidencias).
 *   - Nombre repetido en la hoja o en el CSV (ambiguo) -> NO se toca a esos
 *     estudiantes y se cuentan en el resumen.
 *   - Si el valor del CSV viene vacío se conserva lo que haya en la hoja
 *     (SOBRESCRIBIR_CON_VACIOS = false).
 *   - La fecha se convierte a fecha real (CONVERTIR_FECHA); si no se reconoce el
 *     formato se copia el texto tal cual. País y nacionalidad se copian tal cual
 *     vienen (sin cambiar mayúsculas ni tildes).
 *   - Solo se reescriben las columnas en las que algún estudiante ha cambiado.
 *   - Resumen solo con totales. Si hay estudiantes en la hoja pero ninguno
 *     coincide con el CSV, el resumen lo avisa (CSV u hoja equivocados).
 *
 * NOTAS TÉCNICAS PARA FUTURAS MEJORAS
 *   - Las lecturas/escrituras se hacen por bloques (getValues/setValues) por
 *     rendimiento; evitar leer/escribir celda a celda.
 *   - La columna de curso se reescribe completa en el rango de estudiantes
 *     existentes (si tuviera fórmulas, se convertirían en valores). Lo mismo
 *     ocurre en el paso 2 con las columnas de fecha de nacimiento, país y
 *     nacionalidad; cada una se lee y escribe por separado para no tocar las
 *     columnas intermedias.
 *   - Las fechas se escriben como objeto Date a las 12:00 para que una diferencia
 *     entre la zona horaria del script y la de la hoja no desplace el día.
 *   - El último estudiante es la última fila con contenido en las columnas de
 *     apellidos/nombre (no se usa getLastRow() a ciegas, porque otras columnas
 *     pueden tener datos más abajo).
 *   - El acceso a Drive (DriveApp) requiere autorizar el permiso de Drive la
 *     primera vez que se ejecuta.
 *   - Los cuadros de alerta de Apps Script tienen tamaño limitado, por eso el
 *     resumen solo contiene totales.
 *   - La validación del tipo de archivo (comprobarEsCsv_) usa la extensión
 *     .csv o los tipos MIME de MIME_TYPES_CSV; rechaza siempre los archivos
 *     nativos de Google (application/vnd.google-apps.*).
 * ============================================================================
 */


/* ============================================================================
 * CONFIGURACIÓN (constantes editables)
 * ========================================================================== */

/* ----------------------------------------------------------------------------
 * CONFIGURACIÓN DEL SPREADSHEET
 * -------------------------------------------------------------------------- */

/** Número de filas de encabezado que se saltan (el primer estudiante está en HEADER_ROWS + 1). */
const HEADER_ROWS = 2;

/** Letra de la columna del apellido 1 en la hoja. */
const COL_APELLIDO1 = 'A';

/** Letra de la columna del apellido 2 en la hoja. */
const COL_APELLIDO2 = 'B';

/** Letra de la columna del nombre en la hoja. */
const COL_NOMBRE = 'C';

/** Letra de la columna donde está (y se escribe) el curso del estudiante. */
const COL_CURSO = 'Q';

/** Letra de la columna donde se copia el resto de la información del curso del CSV (grupo, "(R)", etc.) y donde se escriben las incidencias. */
const COL_CURSO_EXTRA = 'R';

/** Letra de la columna donde el paso 2 copia la fecha de nacimiento del CSV. */
const COL_FECHA_NACIMIENTO = 'D';

/** Letra de la columna donde el paso 2 copia el país del CSV. */
const COL_PAIS = 'F';

/** Letra de la columna donde el paso 2 copia la nacionalidad del CSV. */
const COL_NACIONALIDAD = 'G';

/** Valores válidos de la columna de curso en la hoja. Lo que no encaje en esta lista se marca como no reconocido. */
const CURSOS_VALIDOS = ['1ESO', '2ESO', '3ESO', '4ESO', '1BTO', '2BTO', 'AE', 'FPB'];

/** Texto que se escribe en la columna COL_CURSO_EXTRA (el curso queda en blanco) de los estudiantes de la hoja que no aparecen en el CSV. */
const TEXTO_NO_ENCONTRADO = 'NO ENCONTRADO';

/** Texto que se escribe en la columna COL_CURSO_EXTRA (el curso queda en blanco) de los estudiantes con nombre duplicado. */
const TEXTO_DUPLICADO = 'EST DUPLICADO';

/** Texto que se escribe al inicio de la columna COL_CURSO_EXTRA (seguido del valor original del CSV; el curso queda en blanco) cuando el valor del CSV no se puede convertir a un curso válido. */
const TEXTO_CURSO_NO_RECONOCIDO = 'CURSO NO RECONOCIDO';

/** Texto que se escribe en la columna COL_CURSO_EXTRA (el curso queda en blanco) cuando el curso viene vacío en el CSV. */
const TEXTO_CURSO_VACIO = 'CURSO VACÍO';

/** Nombre del menú que se añade a la hoja de cálculo. */
const NOMBRE_MENU = 'Estudiantes';

/**
 * Si es true, para cada estudiante de la hoja sin coincidencia pero con los mismos apellidos en el CSV se muestra un aviso
 * para confirmar si es el mismo (actualizar nombre y curso / solo curso / no es el mismo).
 */
const CONFIRMAR_MISMOS_APELLIDOS = true;

/**
 * (Paso 1) Minutos de ejecución tras los cuales, en lugar de la siguiente pregunta de confirmación, se avisa de que Apps Script
 * limita cada ejecución a 6 minutos y se recomienda guardar lo respondido y volver a lanzar el paso 1.
 */
const MINUTOS_AVISO_LIMITE = 5;

/**
 * (Paso 2) Si es true, la fecha de nacimiento se escribe como fecha real de la hoja (acepta dd/mm/aaaa, dd-mm-aaaa, dd.mm.aaaa y aaaa-mm-dd).
 * Si es false, o si el formato no se reconoce, se copia el texto tal cual viene en el CSV.
 */
const CONVERTIR_FECHA = true;

/**
 * (Paso 2) Si es false, cuando el CSV trae vacía la fecha de nacimiento, el país o la nacionalidad se conserva lo que ya hubiera en la hoja.
 * Si es true, se borra el valor de la hoja. (No afecta a la columna de curso del paso 1, que siempre se escribe.)
 */
const SOBRESCRIBIR_CON_VACIOS = false;

/** Si tras procesar este número de estudiantes de la hoja ninguno ha coincidido con el CSV, se avisa y se permite continuar o abortar. */
const UMBRAL_CONTROL_COINCIDENCIAS = 150;


/* ----------------------------------------------------------------------------
 * CONFIGURACIÓN DEL CSV
 * -------------------------------------------------------------------------- */

/** Si es true, la primera línea con contenido de cada CSV es un encabezado: se comprueba y se salta. */
const CSV_TIENE_CABECERA = true;

/* --- CSV del PASO 1 (cursos): "NIF/Pasaporte","Apellidos","Nombre","Sección" --- */

/** Índice (base 0) del campo "Apellidos" en el CSV de cursos. */
const CSV_CURSOS_IDX_APELLIDOS = 1;

/** Índice (base 0) del campo "Nombre" en el CSV de cursos. */
const CSV_CURSOS_IDX_NOMBRE = 2;

/** Índice (base 0) del campo "Sección" (curso) en el CSV de cursos. */
const CSV_CURSOS_IDX_CURSO = 3;

/* --- CSV del PASO 2 (datos personales): "Nº","Alumno/a","NIF/Pasaporte","Fecha nacimiento","Localidad","Provincia","País","NACIONALIDAD" --- */

/** Índice (base 0) del campo "Alumno/a" (nombre seguido de apellidos, en un solo campo) en el CSV de datos personales. */
const CSV_DATOS_IDX_NOMBRE_COMPLETO = 1;

/** Índice (base 0) del campo "Fecha nacimiento" en el CSV de datos personales. */
const CSV_DATOS_IDX_FECHA_NACIMIENTO = 3;

/** Índice (base 0) del campo "País" en el CSV de datos personales. */
const CSV_DATOS_IDX_PAIS = 6;

/** Índice (base 0) del campo "NACIONALIDAD" en el CSV de datos personales. */
const CSV_DATOS_IDX_NACIONALIDAD = 7;

/**
 * Descripción de cada formato de CSV. La usan prepararFilasCsv_ y los mensajes de error.
 *  - cabecerasEsperadas: nombre que debe tener en el encabezado cada columna usada (se compara sin tildes ni mayúsculas).
 *    Protege de que, si cambia el orden de las columnas, se lean o copien datos de la columna equivocada.
 *  - esLineaValida: indica si una línea ya separada en campos tiene la estructura esperada.
 */
const FORMATO_CSV_CURSOS = {
  paso: 'paso 1 (cursos)',
  cabecera: '"NIF/Pasaporte","Apellidos","Nombre","Sección"',
  ejemplo: '"12345678Z","ORBULESCU","ANA MARÍA","ESO1A"',
  descripcion: 'Se usan apellidos, nombre y sección (curso). El NIF/Pasaporte se ignora.',
  cabecerasEsperadas: [
    { idx: CSV_CURSOS_IDX_APELLIDOS, nombre: 'Apellidos' },
    { idx: CSV_CURSOS_IDX_NOMBRE,    nombre: 'Nombre' },
    { idx: CSV_CURSOS_IDX_CURSO,     nombre: 'Sección' }
  ],
  esLineaValida: esLineaCursosValida_
};

const FORMATO_CSV_DATOS = {
  paso: 'paso 2 (datos personales)',
  cabecera: '"Nº","Alumno/a","NIF/Pasaporte","Fecha nacimiento","Localidad","Provincia","País","NACIONALIDAD"',
  ejemplo: '"18","ANA MARÍA ORBULESCU","12345678Z","03/05/2009","MADRID","MADRID","ESPAÑA","RUMANA"',
  descripcion: 'Se usan Alumno/a (nombre y apellidos en ese orden, sin coma), fecha de nacimiento, país y nacionalidad.',
  cabecerasEsperadas: [
    { idx: CSV_DATOS_IDX_NOMBRE_COMPLETO,  nombre: 'Alumno/a' },
    { idx: CSV_DATOS_IDX_FECHA_NACIMIENTO, nombre: 'Fecha nacimiento' },
    { idx: CSV_DATOS_IDX_PAIS,             nombre: 'País' },
    { idx: CSV_DATOS_IDX_NACIONALIDAD,     nombre: 'NACIONALIDAD' }
  ],
  esLineaValida: esLineaDatosValida_
};

/** Proporción mínima (0 a 1) de líneas no vacías del CSV que deben tener la estructura esperada; si hay menos, se muestra un error y no se modifica nada. */
const PORCENTAJE_MIN_LINEAS_VALIDAS = 0.9;

/** Tipos MIME que se aceptan como CSV (Drive asigna uno u otro según cómo se subió el archivo). */
const MIME_TYPES_CSV = [
  'text/csv',
  'text/comma-separated-values',
  'application/csv',
  'application/vnd.ms-excel',   // Drive lo asigna a veces a CSV subidos desde Windows con Excel
  'text/plain'                  // a veces un .csv se guarda como texto plano
];

/** Valor de curso (de CURSOS_VALIDOS) que se usa cuando en el CSV aparece "AE". */
const CURSO_AE = 'AE';

/** Si es true, cuando en el CSV aparece "AE" junto a un nivel (ej. "AE, ESO3A") el curso es AE. Si es false, el curso es el nivel y "AE" pasa al resto. */
const AE_TIENE_PRIORIDAD = true;

/**
 * Reglas para convertir el nivel del CSV en un curso de la hoja.
 *  - prefijo: texto con el que empieza el nivel en el CSV (ESO, BTO, FPB...). Se admite un espacio y un número opcional.
 *  - formato: curso de la hoja; {n} se sustituye por el número del CSV (ESO2A -> 2ESO). Sin {n}, el curso es fijo (FPB).
 *  - numeroEnResto: si es true, el número no está en el curso y se copia al resto (FPB1A -> FPB | 1A).
 */
const REGLAS_NIVEL = [
  { prefijo: 'ESO', formato: '{n}ESO', numeroEnResto: false },
  { prefijo: 'BTO', formato: '{n}BTO', numeroEnResto: false },
  { prefijo: 'FPB', formato: 'FPB',    numeroEnResto: true }
];


/* ============================================================================
 * MENÚ Y FLUJO DE INTERFAZ
 * ========================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(NOMBRE_MENU)
    .addItem('Paso 1: Actualizar cursos desde CSV', 'actualizarCursos')
    .addItem('Paso 2: Completar fecha de nacimiento, país y nacionalidad', 'actualizarDatosPersonales')
    .addToUi();
}

/** Paso 1: actualiza el curso (y añade/marca estudiantes) a partir del CSV de cursos. */
function actualizarCursos() {
  ejecutarPaso_(
    'Paso 1: cursos',
    'el CSV de cursos (NIF/Pasaporte, Apellidos, Nombre, Sección)',
    procesarCursos_
  );
}

/** Paso 2: rellena fecha de nacimiento, país y nacionalidad a partir del CSV de datos personales. */
function actualizarDatosPersonales() {
  ejecutarPaso_(
    'Paso 2: datos personales',
    'el CSV de datos personales (Nº, Alumno/a, NIF/Pasaporte, Fecha nacimiento, ..., NACIONALIDAD)',
    procesarDatosPersonales_
  );
}

/**
 * Flujo común a los dos pasos: avisa de la hoja de destino, pide el CSV, lo valida y ejecuta el proceso.
 * @param {string} titulo Título del paso (para los cuadros de diálogo).
 * @param {string} descripcionCsv Descripción del CSV que se pide.
 * @param {function(GoogleAppsScript.Spreadsheet.Sheet, string, Object, number): string} procesar
 *   Función que procesa el CSV y devuelve el resumen (recibe también la hora de inicio, en ms).
 */
function ejecutarPaso_(titulo, descripcionCsv, procesar) {
  const inicio = Date.now();   // para el aviso del límite de tiempo de Apps Script
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  // 1. Aviso de la hoja de destino
  const confirmacion = ui.alert(
    titulo + ' - Hoja de destino',
    'Los datos se escribirán en la hoja:\n\n"' + sheetName + '"\n\n¿Quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  // 2. Pedir el CSV
  const respuesta = ui.prompt(
    titulo + ' - Archivo CSV',
    'Sube antes ' + descripcionCsv + ' a Google Drive y pega aquí el enlace de ' +
    '𝗖𝗼𝗺𝗽𝗮𝗿𝘁𝗶𝗿 → 𝗖𝗼𝗽𝗶𝗮𝗿 𝗲𝗻𝗹𝗮𝗰𝗲',
    ui.ButtonSet.OK_CANCEL
  );
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;

  const entrada = respuesta.getResponseText().trim();
  if (!entrada) {
    ui.alert('No has pegado ningún enlace. No se ha modificado nada.');
    return;
  }

  try {
    const archivo = buscarArchivoCsv_(entrada);

    const problema = comprobarEsCsv_(archivo.file);
    if (problema) {
      ui.alert(
        'Archivo no válido',
        problema + '\n\nNo se ha leído el archivo ni se ha modificado nada.',
        ui.ButtonSet.OK
      );
      return;
    }

    const csvText = leerCsv_(archivo.file);
    const resumen = (archivo.aviso ? archivo.aviso + '\n\n' : '') + procesar(sheet, csvText, archivo, inicio);
    Logger.log(resumen);
    ui.alert(titulo + ' - Resumen', resumen, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'No se ha completado el proceso:\n\n' + e.message, ui.ButtonSet.OK);
  }
}


/* ============================================================================
 * ACCESO A DRIVE
 * ========================================================================== */

/**
 * Localiza el CSV en Drive a partir de lo que escribe el usuario.
 * Acepta: URL de Drive, ID del archivo o nombre del archivo.
 * @param {string} entrada
 * @return {{file: GoogleAppsScript.Drive.File, aviso: string}}
 */
function buscarArchivoCsv_(entrada) {
  // a) URL de Drive (…/d/ID/… o …?id=ID)
  const matchUrl = entrada.match(/\/d\/([-\w]{25,})/) || entrada.match(/[?&]id=([-\w]{25,})/);
  if (matchUrl) return { file: abrirPorId_(matchUrl[1]), aviso: '' };

  // b) ID suelto
  if (/^[-\w]{25,}$/.test(entrada)) {
    try {
      return { file: abrirPorId_(entrada), aviso: '' };
    } catch (e) {
      // si no es un ID válido, se intenta como nombre
    }
  }

  // c) Nombre del archivo
  const it = DriveApp.getFilesByName(entrada);
  const candidatos = [];
  while (it.hasNext()) candidatos.push(it.next());
  if (candidatos.length === 0) {
    throw new Error('No se ha encontrado en tu Drive ningún archivo llamado "' + entrada + '".');
  }
  candidatos.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });

  const aviso = candidatos.length > 1
    ? 'AVISO: había ' + candidatos.length + ' archivos con ese nombre; se ha usado el más reciente (' +
      Utilities.formatDate(candidatos[0].getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ').'
    : '';
  return { file: candidatos[0], aviso: aviso };
}

function abrirPorId_(id) {
  try {
    return DriveApp.getFileById(id);
  } catch (e) {
    throw new Error('No se puede abrir el archivo con ID "' + id + '". Comprueba que existe y que tienes acceso.');
  }
}

/**
 * Lee el contenido del archivo detectando la codificación:
 *  - UTF-16 LE/BE (con BOM, o sin BOM si los bytes alternan con ceros). Es lo que generan algunos
 *    programas de gestión escolar y Excel ("Unicode"); si se leyera como UTF-8 saldría texto ilegible.
 *  - UTF-8 (con o sin BOM) y, si tiene caracteres inválidos, windows-1252.
 * @param {GoogleAppsScript.Drive.File} file
 * @return {string} Contenido decodificado y sin BOM.
 */
function leerCsv_(file) {
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length === 0) throw new Error('El archivo CSV está vacío.');

  const b0 = bytes[0] & 0xFF;
  const b1 = bytes.length > 1 ? bytes[1] & 0xFF : -1;

  let texto;
  if (b0 === 0xFF && b1 === 0xFE) {
    texto = blob.getDataAsString('UTF-16LE');
  } else if (b0 === 0xFE && b1 === 0xFF) {
    texto = blob.getDataAsString('UTF-16BE');
  } else if (b1 === 0 && b0 !== 0) {
    texto = blob.getDataAsString('UTF-16LE');   // UTF-16 LE sin BOM
  } else if (b0 === 0 && b1 > 0) {
    texto = blob.getDataAsString('UTF-16BE');   // UTF-16 BE sin BOM
  } else {
    texto = blob.getDataAsString('UTF-8');
    if (texto.indexOf('\uFFFD') !== -1) texto = blob.getDataAsString('windows-1252');
  }

  texto = texto.replace(/^\uFEFF/, '');
  if (texto.indexOf('\u0000') !== -1) {
    throw new Error('No se ha podido interpretar la codificación del CSV (contiene caracteres nulos). ' +
                    'Expórtalo de nuevo como CSV UTF-8 o guárdalo en ese formato desde Excel/Hojas de cálculo.');
  }
  return texto;
}

/**
 * Indica si una línea del CSV de cursos (paso 1) tiene la estructura esperada:
 * campos suficientes para el índice mayor usado y apellidos y nombre no vacíos.
 * El curso (sección) puede venir vacío (se marcará como curso vacío).
 * @param {string[]} row Campos de la línea ya separados.
 * @return {boolean}
 */
function esLineaCursosValida_(row) {
  const minCampos = Math.max(CSV_CURSOS_IDX_APELLIDOS, CSV_CURSOS_IDX_NOMBRE, CSV_CURSOS_IDX_CURSO) + 1;
  if (row.length < minCampos) return false;
  return collapse_(row[CSV_CURSOS_IDX_APELLIDOS]) !== '' && collapse_(row[CSV_CURSOS_IDX_NOMBRE]) !== '';
}

/**
 * Indica si una línea del CSV de datos personales (paso 2) tiene la estructura esperada:
 * campos suficientes para el índice mayor usado y "Alumno/a" no vacío.
 * @param {string[]} row Campos de la línea ya separados.
 * @return {boolean}
 */
function esLineaDatosValida_(row) {
  const minCampos = Math.max(
    CSV_DATOS_IDX_NOMBRE_COMPLETO, CSV_DATOS_IDX_FECHA_NACIMIENTO,
    CSV_DATOS_IDX_PAIS, CSV_DATOS_IDX_NACIONALIDAD
  ) + 1;
  if (row.length < minCampos) return false;
  return collapse_(row[CSV_DATOS_IDX_NOMBRE_COMPLETO]) !== '';
}

/**
 * Mensaje de error estándar que explica cómo debe ser el CSV del paso indicado.
 * @param {Object} formato FORMATO_CSV_CURSOS o FORMATO_CSV_DATOS.
 * @param {string} detalle Qué ha fallado.
 */
function mensajeEstructuraCsv_(formato, detalle) {
  return 'El CSV no tiene la estructura esperada para el ' + formato.paso + '. No se ha modificado nada.\n\n' +
    'Debe ser un CSV con una línea de encabezado y después una línea por estudiante, con los campos entre ' +
    'comillas separados por comas:\n\n' +
    formato.cabecera + '\n' +
    formato.ejemplo + '\n\n' +
    formato.descripcion + '\n\n' + detalle;
}

/**
 * Valida la estructura del CSV ya separado en filas. Lanza un Error explicando cómo debe ser
 * si no hay filas o si menos de PORCENTAJE_MIN_LINEAS_VALIDAS de las líneas no vacías son válidas.
 * @param {string[][]} csvRows
 * @param {number} idxCabecera Índice de la fila de encabezado (se excluye del recuento), o -1 si no hay.
 * @param {Object} formato FORMATO_CSV_CURSOS o FORMATO_CSV_DATOS.
 */
function validarEstructuraCsv_(csvRows, idxCabecera, formato) {
  const noVacias = [];
  csvRows.forEach(function (row, i) {
    if (i <= idxCabecera) return;   // encabezado (y posibles líneas vacías previas)
    const vacia = row.every(function (c) { return String(c).trim() === ''; });
    if (!vacia) noVacias.push({ row: row, linea: i + 1 });
  });

  if (noVacias.length === 0) {
    throw new Error(mensajeEstructuraCsv_(formato, 'El archivo no contiene ninguna línea con datos.'));
  }

  const invalidas = noVacias.filter(function (x) { return !formato.esLineaValida(x.row); });
  const validas = noVacias.length - invalidas.length;

  if (validas / noVacias.length < PORCENTAJE_MIN_LINEAS_VALIDAS) {
    const ejemplos = invalidas.slice(0, 3).map(function (x) {
      const texto = x.row.join(' | ');
      return '  Línea ' + x.linea + ': ' + (texto.length > 70 ? texto.substring(0, 70) + '…' : texto);
    });
    throw new Error(mensajeEstructuraCsv_(formato,
      'Líneas válidas: ' + validas + ' de ' + noVacias.length +
      ' (se exige al menos el ' + Math.round(PORCENTAJE_MIN_LINEAS_VALIDAS * 100) + '%).\n\n' +
      'Ejemplos de líneas incorrectas (campos separados por " | "):\n' + ejemplos.join('\n')
    ));
  }
}

/**
 * Separa el CSV en filas, comprueba el encabezado y la estructura (lanza un Error explicativo si algo
 * no es válido) y devuelve las líneas válidas. Es común a los dos pasos.
 * @param {string} csvText Contenido del CSV ya decodificado.
 * @param {Object} formato FORMATO_CSV_CURSOS o FORMATO_CSV_DATOS.
 * @return {{validas: string[][], ignoradas: string[]}} Filas válidas (sin encabezado ni líneas vacías)
 *   y descripción de las líneas inválidas que se ignoran.
 */
function prepararFilasCsv_(csvText, formato) {
  let csvRows;
  try {
    csvRows = Utilities.parseCsv(csvText.replace(/^\uFEFF/, ''));
  } catch (e) {
    throw new Error(mensajeEstructuraCsv_(formato, 'No se ha podido separar el contenido en campos (' + e.message + ').'));
  }

  const idxCabecera = indiceCabecera_(csvRows);
  if (idxCabecera >= 0) {
    const problemaCabecera = comprobarCabecera_(csvRows[idxCabecera], formato);
    if (problemaCabecera) throw new Error(mensajeEstructuraCsv_(formato, problemaCabecera));
  }
  validarEstructuraCsv_(csvRows, idxCabecera, formato);

  const validas = [];
  const ignoradas = [];
  csvRows.forEach(function (row, i) {
    if (i <= idxCabecera) return;   // encabezado
    const vacia = row.every(function (c) { return String(c).trim() === ''; });
    if (vacia) return;
    if (!formato.esLineaValida(row)) {
      ignoradas.push('Línea ' + (i + 1) + ': ' + row.join(','));
      return;
    }
    validas.push(row);
  });
  return { validas: validas, ignoradas: ignoradas };
}

/**
 * Comprueba que el archivo es un CSV sin leer su contenido (solo nombre y tipo MIME).
 * Es válido si la extensión es .csv o si el tipo MIME es de CSV, y nunca si es un
 * archivo nativo de Google (Hojas de cálculo, Docs...) ni un tipo claramente distinto.
 * @param {GoogleAppsScript.Drive.File} file
 * @return {string} Mensaje del problema, o '' si el archivo es válido.
 */
function comprobarEsCsv_(file) {
  const nombre = file.getName();
  const mime = file.getMimeType();
  const tieneExtensionCsv = /\.csv$/i.test(nombre);
  const mimeEsCsv = MIME_TYPES_CSV.indexOf(mime) !== -1;

  if (mime.indexOf('application/vnd.google-apps.') === 0) {
    return 'El archivo "' + nombre + '" es un archivo de Google (' + mime + '), no un CSV.\n' +
           'Descárgalo como CSV (Archivo → Descargar → Valores separados por comas) y súbelo a Drive.';
  }
  if (!tieneExtensionCsv && !mimeEsCsv) {
    return 'El archivo "' + nombre + '" no parece un CSV (tipo: ' + mime + ').';
  }
  if (!tieneExtensionCsv && mime === 'text/plain') {
    return 'El archivo "' + nombre + '" es texto plano y no tiene extensión .csv.';
  }
  return '';
}


/* ============================================================================
 * LÓGICA PRINCIPAL
 * ========================================================================== */

/**
 * PASO 1. Procesa el CSV de cursos y actualiza la hoja indicada.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja de destino.
 * @param {string} csvText Contenido del CSV (ya decodificado).
 * @param {{file: GoogleAppsScript.Drive.File, aviso: string}} archivo Datos del archivo origen (para el resumen).
 * @param {number} [inicio] Hora de inicio de la ejecución en ms (Date.now()); por defecto, la de esta llamada.
 * @return {string} Resumen para mostrar al usuario.
 */
function procesarCursos_(sheet, csvText, archivo, inicio) {
  inicio = inicio || Date.now();
  const sheetName = sheet.getName();
  const colA1 = colIndex_(COL_APELLIDO1);
  const colA2 = colIndex_(COL_APELLIDO2);
  const colNom = colIndex_(COL_NOMBRE);
  const colCurso = colIndex_(COL_CURSO);
  const colExtra = colIndex_(COL_CURSO_EXTRA);
  const firstRow = HEADER_ROWS + 1;

  // ---------- 1. Leer y parsear el CSV ----------
  // Comprueba el encabezado y la estructura (lanza un error explicando el formato esperado si no es válido)
  const preparado = prepararFilasCsv_(csvText, FORMATO_CSV_CURSOS);
  const lineasIgnoradas = preparado.ignoradas;
  const csvEntries = [];                 // en orden del CSV
  const csvMap = new Map();              // key -> [entries]

  preparado.validas.forEach(function (row) {
    const apellidos = collapse_(row[CSV_CURSOS_IDX_APELLIDOS]);
    const nombre = collapse_(row[CSV_CURSOS_IDX_NOMBRE]);

    const cursoInfo = parsearCurso_(row[CSV_CURSOS_IDX_CURSO]);
    const entry = {
      key: normalize_(apellidos + ' ' + nombre),
      apeKey: normalize_(apellidos),
      nomKey: normalize_(nombre),
      apellidos: apellidos.toLocaleUpperCase('es'),
      nombre: nombre.toLocaleUpperCase('es'),
      curso: cursoInfo.curso,
      extra: cursoInfo.extra,
      cursoReconocido: cursoInfo.reconocido,
      cursoVacio: cursoInfo.vacio,
      cursoOriginal: collapse_(row[CSV_CURSOS_IDX_CURSO])
    };
    csvEntries.push(entry);
    if (!csvMap.has(entry.key)) csvMap.set(entry.key, []);
    csvMap.get(entry.key).push(entry);
  });

  if (csvEntries.length === 0) {
    throw new Error('El CSV no contiene ningún estudiante válido. No se ha modificado nada.');
  }

  // ---------- 2. Leer estudiantes de la hoja ----------
  const hoja = leerEstudiantesHoja_(sheet);
  const ap1 = hoja.ap1, ap2 = hoja.ap2, nom = hoja.nom;
  const numStudents = hoja.numStudents;
  const lastStudentRow = hoja.lastStudentRow;     // si no hay estudiantes = firstRow - 1

  const sheetKeys = [];                  // key por fila ('' si la fila está vacía)
  const sheetMap = new Map();            // key -> [índices]
  const sheetNames = [];
  for (let i = 0; i < numStudents; i++) {
    const nombreHoja = [ap1[i][0], ap2[i][0], nom[i][0]].join(' ');
    const key = normalize_(nombreHoja);
    sheetKeys.push(key);
    sheetNames.push(collapse_(nombreHoja));
    if (key === '') continue;
    if (!sheetMap.has(key)) sheetMap.set(key, []);
    sheetMap.get(key).push(i);
  }

  // ---------- 2b. Confirmación manual: mismos apellidos, nombre distinto ----------
  // Para los estudiantes de la hoja sin coincidencia exacta que tienen en el CSV una entrada con los
  // mismos apellidos (y que no coincide ya con otra fila de la hoja), se pregunta al usuario si es el mismo:
  //   Sí (YES)         -> actualizar SOLO el nombre (columna C) y el curso; los apellidos de la hoja no se tocan
  //   No (NO)          -> actualizar solo el curso (el nombre de la hoja no se toca), previa segunda confirmación Sí/No
  //   Cancelar / cerrar -> no es el mismo estudiante; se prueba el siguiente candidato (si lo hay) y se continúa
  // Todavía no se ha escrito nada en la hoja: las respuestas solo se aplican en el paso 5.
  const csvPorApellidos = new Map();   // apeKey -> [entries]
  csvEntries.forEach(function (e) {
    if (!csvPorApellidos.has(e.apeKey)) csvPorApellidos.set(e.apeKey, []);
    csvPorApellidos.get(e.apeKey).push(e);
  });

  /** Entradas del CSV con los mismos apellidos que la fila i (sin duplicados en el CSV); [] si la fila ya está emparejada o vacía. */
  function candidatosDeFila(i) {
    const key = sheetKeys[i];
    if (key === '' || csvMap.has(key)) return [];   // fila vacía o ya emparejada
    const apeKey = normalize_(ap1[i][0] + ' ' + ap2[i][0]);
    if (apeKey === '') return [];
    return (csvPorApellidos.get(apeKey) || []).filter(function (e) {
      return csvMap.get(e.key).length === 1;        // sin duplicados en el CSV
    });
  }

  // Total previsto de preguntas (para el contador "pregunta X/N"): filas con algún candidato libre. Es aproximado.
  let totalPreguntas = 0;
  if (CONFIRMAR_MISMOS_APELLIDOS) {
    for (let i = 0; i < numStudents; i++) {
      const hayCandidatoLibre = candidatosDeFila(i).some(function (e) { return !sheetMap.has(e.key); });
      if (hayCandidatoLibre) totalPreguntas++;
    }
  }

  let confirmadosNombreYCurso = 0;
  let confirmadosSoloCurso = 0;
  let cambioNombres = false;
  let numPregunta = 0;               // filas por las que se ha empezado a preguntar
  let interrumpido = false;          // true si se detiene la confirmación por el límite de tiempo (guardado parcial)
  let preguntasPendientes = 0;
  let avisoTiempoMostrado = false;
  if (CONFIRMAR_MISMOS_APELLIDOS) {
    const ui = SpreadsheetApp.getUi();
    for (let i = 0; i < numStudents; i++) {
      const candidatos = candidatosDeFila(i);
      let filaContada = false;

      for (let c = 0; c < candidatos.length; c++) {
        const e = candidatos[c];
        if (sheetMap.has(e.key)) continue;           // ya emparejada con otra fila (o confirmada antes)

        // Aviso del límite de 6 minutos de Apps Script, en lugar de la siguiente pregunta
        if (!avisoTiempoMostrado && (Date.now() - inicio) >= MINUTOS_AVISO_LIMITE * 60 * 1000) {
          avisoTiempoMostrado = true;
          const pendientes = Math.max(totalPreguntas - numPregunta + (filaContada ? 1 : 0), 1);
          const decision = ui.alert(
            '⌛Límite de tiempo de Apps Script⌛',
            'Llevas más de ' + MINUTOS_AVISO_LIMITE + ' minutos en esta ejecución.\n\n' +
            'Google Apps Script limita cada ejecución a 6 minutos. Si se supera, el proceso se corta y se ' +
            'pierden TODAS las respuestas dadas hasta ahora.\n\n' +
            'Recomendamos guardar los avances y volver a lanzar el Paso 1 desde el menú para continuar con el ' +
            'resto (quedan unas ' + pendientes + ' preguntas de ' + totalPreguntas + ').\n\n' +
            'Sí  =  Guardar los avances y detener (recomendado)\n' +
            'No  =  Seguir preguntando (si se supera el límite se perderá todo)',
            ui.ButtonSet.YES_NO
          );
          if (decision === ui.Button.YES) {
            interrumpido = true;
            preguntasPendientes = pendientes;
            break;
          }
        }

        if (!filaContada) { numPregunta++; filaContada = true; }

        // Icono del aviso: ✅ si aparentemente solo falta parte del nombre en la hoja (el nombre de la hoja es el
        // principio del nombre del CSV, ej. "DULCE" frente a "DULCE ZOE"); ⚠️ en cualquier otro caso (hermanos, erratas...)
        const nomHoja = normalize_(nom[i][0]);
        const icono = (nomHoja !== '' && nombreEsInicioDe_(nomHoja, e.nomKey)) ? '✅' : '⚠️';
        // Con ⚠️ se deja un salto de línea más antes de "Hoja:" y antes de "¿Es el mismo estudiante?" para que el cuadro
        // se vea distinto a simple vista
        const separacion = (icono === '✅') ? '\n\n' : '\n\n\n';

        let respuesta;
        let preguntar = true;
        while (preguntar) {
          preguntar = false;
          respuesta = ui.alert(
          'Estudiante no encontrado - pregunta ' + numPregunta + '/' + totalPreguntas,
          icono + ' No se ha encontrado este estudiante, pero hay uno ' + icono + ' con los mismos apellidos en el CSV:' + separacion +
          'Hoja:  "' + sheetNames[i] + '" (fila ' + (firstRow + i) + ')\n' +
          'CSV:  "' + e.apellidos + ', ' + e.nombre + '" (curso en el CSV: ' + (e.cursoOriginal || 'vacío') + ')' + separacion +
          '¿Es el mismo estudiante?\n' +
          '(Los botones de Apps Script no se pueden renombrar, por eso se explican aquí)\n\n' +
          'Sí  =  Es el mismo: actualizar el nombre y el curso (los apellidos no se tocan)\n' +
          'No  =  Es el mismo, pero NO cambiar el nombre: actualizar solo el curso (te pediré confirmación)\n' +
          'Cancelar  =  No es el mismo: no actualizar nada con esta entrada y continuar',
          ui.ButtonSet.YES_NO_CANCEL
          );

          // "No" = actualizar SOLO el curso: se pide una segunda confirmación (Sí / No).
          // Si en la segunda se responde No, se vuelve a mostrar la pregunta anterior.
          if (respuesta === ui.Button.NO) {
            const confirmaSoloCurso = ui.alert(
              'Confirmar: actualizar solo el curso',
              'Se actualizará SOLO el curso del estudiante "' + sheetNames[i] + '".\n' +
              'El nombre en la hoja NO se modificará.\n\n' +
              '¿Confirmas?\n(Sí = confirmar, No = volver a la pregunta anterior)',
              ui.ButtonSet.YES_NO
            );
            if (confirmaSoloCurso !== ui.Button.YES) preguntar = true;
          }
        }

        if (respuesta === ui.Button.YES || respuesta === ui.Button.NO) {
          sheetKeys[i] = e.key;
          sheetMap.set(e.key, [i]);
          if (respuesta === ui.Button.YES) {
            nom[i][0] = e.nombre;           // solo el nombre: los apellidos de la hoja no se modifican
            cambioNombres = true;
            confirmadosNombreYCurso++;
          } else {
            confirmadosSoloCurso++;
          }
          break;
        }
        // Cancelar o cerrar el cuadro: no es el mismo; se prueba el siguiente candidato
      }
      if (interrumpido) break;
    }
  }

  // ---------- 3. Actualizar cursos de los estudiantes existentes ----------
  let cursos = [];
  let extras = [];
  if (numStudents > 0) {
    cursos = sheet.getRange(firstRow, colCurso, numStudents, 1).getValues();
    extras = sheet.getRange(firstRow, colExtra, numStudents, 1).getValues();
  }

  let actualizados = 0;
  const noEncontrados = [];
  const duplicados = new Set();

  let procesados = 0;
  let coincidencias = 0;
  let controlHecho = false;

  for (let i = 0; i < numStudents; i++) {
    const key = sheetKeys[i];
    if (key === '') continue;

    // Control: tras UMBRAL_CONTROL_COINCIDENCIAS estudiantes sin ninguna coincidencia, avisar.
    // Todavía no se ha escrito nada en la hoja, así que abortar aquí no modifica nada.
    if (!controlHecho && procesados >= UMBRAL_CONTROL_COINCIDENCIAS && coincidencias === 0) {
      controlHecho = true;
      const ui = SpreadsheetApp.getUi();
      const respuesta = ui.alert(
        'Ningún estudiante coincide',
        'Se han revisado ' + procesados + ' estudiantes de la hoja y ninguno coincide con el CSV.\n\n' +
        'Lo normal es que coincidan casi todos, así que puede haber un problema (hoja equivocada, ' +
        'columnas distintas, formato del CSV...).\n\n' +
        'Si continúas, los estudiantes de la hoja que no estén en el CSV se marcarán como no encontrados ' +
        'y los del CSV se añadirán como nuevos.\n\n' +
        '¿Quieres continuar?\n(Sí = continuar, No = abortar sin modificar nada)',
        ui.ButtonSet.YES_NO
      );
      if (respuesta !== ui.Button.YES) {
        return 'PROCESO ABORTADO por el usuario. No se ha modificado nada.';
      }
    }
    procesados++;

    const csvList = csvMap.get(key);
    if (csvList) coincidencias++;
    if (!csvList) {
      if (interrumpido) continue;   // guardado parcial: la fila aún puede tener pregunta pendiente, no se marca
      cursos[i][0] = '';
      extras[i][0] = TEXTO_NO_ENCONTRADO;
      noEncontrados.push(sheetNames[i]);
    } else if (csvList.length > 1 || sheetMap.get(key).length > 1) {
      cursos[i][0] = '';
      extras[i][0] = TEXTO_DUPLICADO;
      duplicados.add(sheetNames[i]);
    } else {
      cursos[i][0] = csvList[0].curso;
      extras[i][0] = csvList[0].extra;
      actualizados++;
    }
  }

  // ---------- 4. Determinar estudiantes nuevos (en orden del CSV) ----------
  const nuevos = [];
  const vistos = new Map();
  // En guardado parcial no se añade ninguno: sus entradas podrían ser el emparejamiento de una pregunta pendiente
  (interrumpido ? [] : csvEntries).forEach(function (entry) {
    const veces = (vistos.get(entry.key) || 0) + 1;
    vistos.set(entry.key, veces);

    const enHoja = sheetMap.has(entry.key) ? sheetMap.get(entry.key).length : 0;
    if (veces <= enHoja) return; // ya tiene su fila en la hoja

    const esDuplicado = csvMap.get(entry.key).length > 1 || enHoja > 1;
    if (esDuplicado) duplicados.add(entry.apellidos + ', ' + entry.nombre);
    nuevos.push({
      apellido1: separarApellidos_(entry.apellidos).apellido1,
      apellido2: separarApellidos_(entry.apellidos).apellido2,
      nombre: entry.nombre,
      curso: esDuplicado ? '' : entry.curso,
      extra: esDuplicado ? TEXTO_DUPLICADO : entry.extra
    });
  });

  // ---------- 5. Escribir en la hoja ----------
  if (numStudents > 0) {
    sheet.getRange(firstRow, colCurso, numStudents, 1).setValues(cursos);
    sheet.getRange(firstRow, colExtra, numStudents, 1).setValues(extras);
  }

  if (cambioNombres) {
    sheet.getRange(firstRow, colNom, numStudents, 1).setValues(nom.slice(0, numStudents));
  }

  if (nuevos.length > 0) {
    const startRow = lastStudentRow + 1;
    const endRow = startRow + nuevos.length - 1;
    const maxRows = sheet.getMaxRows();
    if (endRow > maxRows) sheet.insertRowsAfter(maxRows, endRow - maxRows);

    sheet.getRange(startRow, colA1, nuevos.length, 1)
      .setValues(nuevos.map(function (n) { return [n.apellido1]; }));
    sheet.getRange(startRow, colA2, nuevos.length, 1)
      .setValues(nuevos.map(function (n) { return [n.apellido2]; }));
    sheet.getRange(startRow, colNom, nuevos.length, 1)
      .setValues(nuevos.map(function (n) { return [n.nombre]; }));
    sheet.getRange(startRow, colCurso, nuevos.length, 1)
      .setValues(nuevos.map(function (n) { return [n.curso]; }));
    sheet.getRange(startRow, colExtra, nuevos.length, 1)
      .setValues(nuevos.map(function (n) { return [n.extra]; }));
  }
  SpreadsheetApp.flush();

  // ---------- 6. Resumen ----------
  const cursosNoReconocidos = csvEntries.filter(function (e) { return !e.cursoReconocido && !e.cursoVacio; }).length;
  const cursosVacios = csvEntries.filter(function (e) { return e.cursoVacio; }).length;

  const resumen = [
    'Estudiantes en el CSV: ' + csvEntries.length,
    'Actualizados (curso del CSV): ' + actualizados,
    'Confirmados por el usuario, nombre y curso actualizados (incluidos en las anteriores): ' + confirmadosNombreYCurso,
    'Confirmados por el usuario, solo curso actualizado (incluidos en las anteriores): ' + confirmadosSoloCurso,
    'Nuevos añadidos: ' + nuevos.length,
    'No encontrados en el CSV: ' + noEncontrados.length,
    'Duplicados: ' + duplicados.size,
    'Líneas del CSV ignoradas: ' + lineasIgnoradas.length,
    'Cursos no reconocidos: ' + cursosNoReconocidos,
    'Estudiantes con el curso vacío: ' + cursosVacios
  ];
  if (interrumpido) {
    resumen.unshift(
      'GUARDADO PARCIAL por el límite de tiempo de Apps Script.',
      'Se han guardado los estudiantes ya emparejados y tus respuestas. NO se han marcado estudiantes como no encontrados ' +
        'ni se han añadido nuevos (se hará cuando termines).',
      'Vuelve a lanzar el Paso 1 desde el menú para continuar con el resto (quedaban unas ' + preguntasPendientes + ' preguntas). ' +
        'Los estudiantes a los que respondiste "No" o "Cancelar" se volverán a preguntar.',
      ''
    );
  }
  return resumen.join('\n');
}


/**
 * PASO 2. Procesa el CSV de datos personales: para cada estudiante de la hoja cuyo nombre completo es
 * IDÉNTICO (normalizado) al campo "Alumno/a" del CSV, copia la fecha de nacimiento, el país y la nacionalidad.
 * No hace nada con los que no coinciden ni con los nombres repetidos, y nunca añade filas ni cambia nombres.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja de destino.
 * @param {string} csvText Contenido del CSV (ya decodificado).
 * @param {{file: GoogleAppsScript.Drive.File, aviso: string}} archivo Datos del archivo origen.
 * @return {string} Resumen para mostrar al usuario.
 */
function procesarDatosPersonales_(sheet, csvText, archivo) {
  const colFecha = colIndex_(COL_FECHA_NACIMIENTO);
  const colPais = colIndex_(COL_PAIS);
  const colNac = colIndex_(COL_NACIONALIDAD);

  // ---------- 1. Leer y parsear el CSV ----------
  const preparado = prepararFilasCsv_(csvText, FORMATO_CSV_DATOS);
  const csvEntries = [];
  const csvMap = new Map();              // key -> [entries]

  preparado.validas.forEach(function (row) {
    const entry = {
      key: normalizarNombreEstricto_(row[CSV_DATOS_IDX_NOMBRE_COMPLETO]),
      fechaNacimiento: parsearFecha_(row[CSV_DATOS_IDX_FECHA_NACIMIENTO]),
      pais: collapse_(row[CSV_DATOS_IDX_PAIS]),
      nacionalidad: collapse_(row[CSV_DATOS_IDX_NACIONALIDAD])
    };
    if (entry.key === '') {
      preparado.ignoradas.push('Línea sin nombre utilizable: ' + row.join(','));
      return;
    }
    csvEntries.push(entry);
    if (!csvMap.has(entry.key)) csvMap.set(entry.key, []);
    csvMap.get(entry.key).push(entry);
  });

  if (csvEntries.length === 0) {
    throw new Error('El CSV no contiene ningún estudiante válido. No se ha modificado nada.');
  }

  // ---------- 2. Leer estudiantes de la hoja ----------
  // Clave de la hoja: NOMBRE + APELLIDO1 + APELLIDO2 (el CSV trae nombre y luego apellidos)
  const hoja = leerEstudiantesHoja_(sheet);
  const numStudents = hoja.numStudents;

  const sheetKeys = [];
  const sheetMap = new Map();            // key -> [índices de fila]
  for (let i = 0; i < numStudents; i++) {
    const key = normalizarNombreEstricto_([hoja.nom[i][0], hoja.ap1[i][0], hoja.ap2[i][0]].join(' '));
    sheetKeys.push(key);
    if (key === '') continue;
    if (!sheetMap.has(key)) sheetMap.set(key, []);
    sheetMap.get(key).push(i);
  }

  // ---------- 3. Copiar los datos de los estudiantes con nombre idéntico y único ----------
  let fechas = [], paises = [], nacionalidades = [];
  if (numStudents > 0) {
    fechas = sheet.getRange(hoja.firstRow, colFecha, numStudents, 1).getValues();
    paises = sheet.getRange(hoja.firstRow, colPais, numStudents, 1).getValues();
    nacionalidades = sheet.getRange(hoja.firstRow, colNac, numStudents, 1).getValues();
  }

  let estudiantesHoja = 0;
  let actualizados = 0;
  let sinCoincidencia = 0;
  const duplicados = new Set();
  let cambioFecha = false, cambioPais = false, cambioNac = false;

  for (let i = 0; i < numStudents; i++) {
    const key = sheetKeys[i];
    if (key === '') continue;
    estudiantesHoja++;

    const csvList = csvMap.get(key);
    if (!csvList) {
      sinCoincidencia++;
      continue;
    }
    if (csvList.length > 1 || sheetMap.get(key).length > 1) {
      duplicados.add(key);                // ambiguo: no se toca
      continue;
    }

    const e = csvList[0];
    const f = aplicarDato_(fechas, i, e.fechaNacimiento);
    const p = aplicarDato_(paises, i, e.pais);
    const n = aplicarDato_(nacionalidades, i, e.nacionalidad);
    if (f) cambioFecha = true;
    if (p) cambioPais = true;
    if (n) cambioNac = true;
    if (f || p || n) actualizados++;
  }

  const csvSinHoja = csvEntries.filter(function (e) { return !sheetMap.has(e.key); }).length;
  const emparejados = estudiantesHoja - sinCoincidencia;   // incluye los duplicados (no tocados)

  // ---------- 4. Escribir en la hoja (solo las columnas que han cambiado) ----------
  if (cambioFecha) sheet.getRange(hoja.firstRow, colFecha, numStudents, 1).setValues(fechas);
  if (cambioPais) sheet.getRange(hoja.firstRow, colPais, numStudents, 1).setValues(paises);
  if (cambioNac) sheet.getRange(hoja.firstRow, colNac, numStudents, 1).setValues(nacionalidades);
  SpreadsheetApp.flush();

  // ---------- 5. Resumen ----------
  const lineas = [];
  if (estudiantesHoja > 0 && emparejados === 0) {
    lineas.push('AVISO: ningún estudiante de la hoja coincide con el CSV. ¿Hoja o CSV equivocados?\n');
  }
  lineas.push(
    'Estudiantes en el CSV: ' + csvEntries.length,
    'Estudiantes en la hoja: ' + estudiantesHoja,
    'Con fecha de nacimiento, país o nacionalidad rellenados/actualizados: ' + actualizados,
    'Nombres repetidos (no modificados): ' + duplicados.size,
    'Estudiantes de la hoja sin coincidencia en el CSV (no modificados): ' + sinCoincidencia,
    'Estudiantes del CSV sin coincidencia en la hoja (ignorados): ' + csvSinHoja,
    'Líneas del CSV ignoradas: ' + preparado.ignoradas.length
  );
  return lineas.join('\n');
}

/**
 * Lee los estudiantes de la hoja (apellido 1, apellido 2 y nombre).
 * El último estudiante es la última fila con algo en esas tres columnas (no se usa getLastRow() a ciegas).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {{firstRow: number, ap1: Array[], ap2: Array[], nom: Array[], numStudents: number, lastStudentRow: number}}
 *   ap1/ap2/nom son matrices [[v], ...] con al menos numStudents filas; lastStudentRow = firstRow - 1 si no hay estudiantes.
 */
function leerEstudiantesHoja_(sheet) {
  const firstRow = HEADER_ROWS + 1;
  const lastSheetRow = sheet.getLastRow();
  const numRows = Math.max(lastSheetRow - firstRow + 1, 0);

  let ap1 = [], ap2 = [], nom = [];
  if (numRows > 0) {
    ap1 = sheet.getRange(firstRow, colIndex_(COL_APELLIDO1), numRows, 1).getValues();
    ap2 = sheet.getRange(firstRow, colIndex_(COL_APELLIDO2), numRows, 1).getValues();
    nom = sheet.getRange(firstRow, colIndex_(COL_NOMBRE), numRows, 1).getValues();
  }

  let lastStudentIdx = -1;
  for (let i = 0; i < numRows; i++) {
    if (String(ap1[i][0]).trim() || String(ap2[i][0]).trim() || String(nom[i][0]).trim()) {
      lastStudentIdx = i;
    }
  }
  return {
    firstRow: firstRow,
    ap1: ap1, ap2: ap2, nom: nom,
    numStudents: lastStudentIdx + 1,
    lastStudentRow: firstRow + lastStudentIdx
  };
}


/* ============================================================================
 * UTILIDADES
 * ========================================================================== */

/**
 * Convierte el valor de curso del CSV en un curso válido de la hoja y el resto de información.
 * El valor puede traer varios elementos separados por coma (ej. "AE, ESO3A").
 * @param {string} valorCsv Valor del campo curso del CSV.
 * @return {{curso: string, extra: string, reconocido: boolean, vacio: boolean}}
 *   vacio=true si el curso viene vacío en el CSV (se distingue de "no reconocido").
 */
function parsearCurso_(valorCsv) {
  const original = collapse_(valorCsv);
  if (original === '') {
    return { curso: '', extra: TEXTO_CURSO_VACIO, reconocido: false, vacio: true };
  }
  const tokens = original.split(',').map(collapse_).filter(function (t) { return t !== ''; });

  const aeIdx = tokens.findIndex(function (t) { return t.toUpperCase() === CURSO_AE.toUpperCase(); });
  const otros = tokens.filter(function (t, i) { return i !== aeIdx; });

  let nivel = null;
  for (let i = 0; i < otros.length && !nivel; i++) nivel = interpretarNivel_(otros[i]);

  // AE junto a un nivel (o solo): el curso es AE y el resto es lo demás
  if (aeIdx >= 0 && (AE_TIENE_PRIORIDAD || !nivel)) {
    return { curso: CURSO_AE, extra: otros.join(', '), reconocido: true, vacio: false };
  }

  if (nivel && CURSOS_VALIDOS.indexOf(nivel.curso) !== -1) {
    const extra = aeIdx >= 0
      ? (nivel.resto ? nivel.resto + ', ' : '') + CURSO_AE
      : nivel.resto;
    return { curso: nivel.curso, extra: extra, reconocido: true, vacio: false };
  }

  return { curso: '', extra: TEXTO_CURSO_NO_RECONOCIDO + ': ' + original, reconocido: false, vacio: false };
}

/**
 * Interpreta un nivel del CSV (ej. "ESO 4D (R)") según REGLAS_NIVEL.
 * @param {string} texto
 * @return {{curso: string, resto: string}|null} null si no coincide con ninguna regla.
 */
function interpretarNivel_(texto) {
  for (let i = 0; i < REGLAS_NIVEL.length; i++) {
    const regla = REGLAS_NIVEL[i];
    const m = texto.match(new RegExp('^' + regla.prefijo + '\\s*(\\d?)\\s*(.*)$', 'i'));
    if (!m) continue;
    return {
      curso: regla.formato.replace('{n}', m[1]),
      resto: regla.numeroEnResto ? (m[1] + m[2]).trim() : m[2].trim()
    };
  }
  return null;
}

/**
 * Indica si los nombres de la hoja son el principio de los nombres del CSV (y el CSV tiene más palabras).
 * Ambos ya normalizados. Ej.: ("dulce", "dulce zoe") -> true; ("zoe", "dulce zoe") -> false; ("luis", "pedro") -> false.
 * @param {string} nombreHoja
 * @param {string} nombreCsv
 * @return {boolean}
 */
function nombreEsInicioDe_(nombreHoja, nombreCsv) {
  const wh = nombreHoja.split(' ');
  const wc = nombreCsv.split(' ');
  if (wh.length >= wc.length) return false;   // con las mismas palabras ya habría coincidido exactamente
  for (let k = 0; k < wh.length; k++) {
    if (wh[k] !== wc[k]) return false;
  }
  return true;
}

/**
 * Reparte los apellidos de un estudiante nuevo entre las columnas de apellido 1 y apellido 2.
 * Si son exactamente 2 palabras, una para cada columna; en cualquier otro caso todos van en el apellido 1.
 * @param {string} apellidos Apellidos ya en mayúsculas y con espacios colapsados.
 * @return {{apellido1: string, apellido2: string}}
 */
function separarApellidos_(apellidos) {
  const palabras = collapse_(apellidos).split(' ').filter(function (p) { return p !== ''; });
  if (palabras.length === 2) return { apellido1: palabras[0], apellido2: palabras[1] };
  return { apellido1: palabras.join(' '), apellido2: '' };
}

/**
 * Índice (base 0) de la fila de encabezado del CSV: la primera fila con contenido.
 * @param {string[][]} csvRows
 * @return {number} -1 si CSV_TIENE_CABECERA es false o no hay ninguna fila con contenido.
 */
function indiceCabecera_(csvRows) {
  if (!CSV_TIENE_CABECERA) return -1;
  for (let i = 0; i < csvRows.length; i++) {
    const vacia = csvRows[i].every(function (c) { return String(c).trim() === ''; });
    if (!vacia) return i;
  }
  return -1;
}

/**
 * Comprueba que la fila de encabezado tiene, en las posiciones usadas, los nombres de formato.cabecerasEsperadas.
 * @param {string[]} row Fila de encabezado.
 * @param {Object} formato FORMATO_CSV_CURSOS o FORMATO_CSV_DATOS.
 * @return {string} Descripción del problema, o '' si es correcta.
 */
function comprobarCabecera_(row, formato) {
  const fallos = formato.cabecerasEsperadas.filter(function (c) {
    const valor = row[c.idx] === undefined ? '' : row[c.idx];
    return normalize_(valor) !== normalize_(c.nombre);
  });
  if (fallos.length === 0) return '';
  return 'La primera línea debería ser el encabezado, pero no tiene las columnas esperadas:\n' +
    fallos.map(function (c) {
      const valor = row[c.idx] === undefined ? '' : collapse_(row[c.idx]);
      return '  Campo ' + (c.idx + 1) + ': se esperaba "' + c.nombre + '" y hay "' + valor + '"';
    }).join('\n');
}

/**
 * Convierte la fecha del CSV en una fecha real de la hoja. Se fija a las 12:00 para que una diferencia de
 * zona horaria entre el script y la hoja no cambie el día.
 * @param {string} valorCsv Ej. "03/05/2009" (dd/mm/aaaa) o "2009-05-03".
 * @return {Date|string} Date si se reconoce y es una fecha válida; si no, el texto tal cual ('' si venía vacío).
 */
function parsearFecha_(valorCsv) {
  const texto = collapse_(valorCsv);
  if (texto === '' || !CONVERTIR_FECHA) return texto;

  let d, m, y;
  let mt = texto.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\s.*)?$/);
  if (mt) {
    d = Number(mt[1]); m = Number(mt[2]); y = Number(mt[3]);
  } else {
    mt = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (!mt) return texto;
    y = Number(mt[1]); m = Number(mt[2]); d = Number(mt[3]);
  }
  const fecha = new Date(y, m - 1, d, 12, 0, 0);
  if (fecha.getFullYear() !== y || fecha.getMonth() !== m - 1 || fecha.getDate() !== d) return texto;  // ej. 31/02
  return fecha;
}

/**
 * Escribe un dato personal en el array de valores de una columna (no toca la hoja).
 * @param {Array[]} celdas Valores de la columna ([[v], [v], ...]).
 * @param {number} i Índice de la fila.
 * @param {Date|string} valor Valor del CSV.
 * @return {boolean} true si se ha escrito (no se escribe un valor vacío salvo SOBRESCRIBIR_CON_VACIOS).
 */
function aplicarDato_(celdas, i, valor) {
  if (valor === '' && !SOBRESCRIBIR_CON_VACIOS) return false;
  celdas[i][0] = valor;
  return true;
}

/** Convierte una letra de columna ('A', 'Q', 'AB'...) en índice base 1. */
function colIndex_(letra) {
  let n = 0;
  const l = String(letra).toUpperCase();
  for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
  return n;
}

/** Colapsa espacios múltiples y recorta extremos. */
function collapse_(s) {
  return String(s).replace(/\s+/g, ' ').trim();
}

/**
 * Normaliza un nombre para comparar: sin comas, sin tildes/diacríticos
 * (ñ -> n), en minúsculas y con un único espacio entre palabras.
 */
function normalize_(s) {
  return String(s)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/,/g, ' ')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normalización del paso 2 (comparación estricta de nombres completos): como normalize_, y además cualquier
 * carácter que no sea letra, número o espacio (guiones, apóstrofos, puntos, símbolos raros) pasa a espacio.
 * Se aplica igual a los dos lados de la comparación. No sustituye a normalize_ (paso 1).
 */
function normalizarNombreEstricto_(s) {
  return normalize_(s)
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}