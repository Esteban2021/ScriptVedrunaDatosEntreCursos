/**
 * ============================================================================
 * INSTRUCCIONES PARA LA IA (leer antes de modificar este script)
 * ============================================================================
 * PROPÓSITO
 *   Script de Google Apps Script ligado a un Google Spreadsheet con el listado
 *   de estudiantes de un centro educativo. Una vez al año (o cuando haga falta)
 *   el usuario sube un CSV a Google Drive con el listado oficial actualizado y
 *   el script sincroniza la hoja con ese CSV.
 *
 * ESTRUCTURA DE LA HOJA (configurable en la sección CONFIGURACIÓN)
 *   - Columna A: apellido 1 | Columna B: apellido 2 | Columna C: nombre
 *   - Columna Q: curso del estudiante, con un valor de CURSOS_VALIDOS
 *     (configurable, ver COL_CURSO)
 *   - Columna R: resto de la información del curso del CSV (grupo, "(R)",
 *     etc.) (configurable, ver COL_CURSO_EXTRA)
 *   - Las primeras HEADER_ROWS filas son encabezado (el primer estudiante está
 *     en la fila 3 por defecto).
 *   - Se trabaja siempre sobre la HOJA ACTIVA; antes de empezar se muestra un
 *     aviso con el nombre de la hoja y hay que confirmar.
 *
 * FORMATO DEL CSV (sin fila de encabezado, separador coma, campos entre comillas)
 *     "18","ORBULESCU , ANA MARÍA","ESO1A"
 *   - Campo 0: número de orden -> SE IGNORA.
 *   - Campo 1: "APELLIDOS , NOMBRE" (apellidos y nombre separados por la
 *     PRIMERA coma).
 *   - Campo 2: curso en formato libre del CSV (ver TRANSFORMACIÓN DEL CURSO).
 *   - Codificación (ver leerCsv_): se detecta automáticamente. UTF-16 LE/BE
 *     (con o sin BOM; es la que genera el sistema de gestión del centro),
 *     UTF-8 (con o sin BOM) y, si tiene caracteres inválidos, windows-1252.
 *     Un CSV UTF-16 leído como UTF-8 se ve como texto con caracteres nulos
 *     ("�") entre cada letra; no cambiar esta detección sin revisarlo.
 *   - VALIDACIÓN DE ESTRUCTURA (ver validarEstructuraCsv_ y esLineaCsvValida_):
 *     una línea es válida si tiene 3 campos, el primero son solo dígitos y el
 *     segundo es "APELLIDOS , NOMBRE" (una coma con texto a ambos lados). Si
 *     menos del PORCENTAJE_MIN_LINEAS_VALIDAS (90%) de las líneas no vacías
 *     son válidas se lanza un error que explica el formato esperado, con
 *     ejemplos de líneas incorrectas, y no se modifica nada. Las pocas líneas
 *     inválidas que queden se ignoran y se cuentan en el resumen.
 *
 * FLUJO (SIN HTML: no se usa HtmlService ni google.script.run, para evitar
 * problemas con varias cuentas de Google abiertas en el navegador)
 *   1. Menú "Estudiantes > Actualizar desde CSV" (onOpen -> actualizarEstudiantes).
 *   2. Aviso de confirmación con el nombre de la hoja activa.
 *   3. Cuadro de texto que pide el nombre, la URL o el ID del CSV en Google Drive
 *      (ver buscarArchivoCsv_).
 *   3b. comprobarEsCsv_ verifica (solo con nombre y tipo MIME, sin leer el
 *       contenido) que el archivo es un CSV. Si no lo es, se muestra un aviso y
 *       se detiene el proceso sin leer el archivo ni modificar la hoja.
 *   4. processCsv_ compara y actualiza la hoja y devuelve un resumen.
 *   5. El resumen (solo totales, sin detalle de estudiantes) se muestra en un
 *      cuadro de alerta y se guarda en el registro de ejecución con Logger.log.
 *
 * REGLAS DE NEGOCIO
 *   - Comparación por nombre completo NORMALIZADO: concatenar
 *     apellido1 + apellido2 + nombre (hoja) frente a "apellidos nombre" (CSV),
 *     sin comas, minúsculas, sin tildes ni diacríticos (la ñ pasa a n) y con
 *     espacios múltiples colapsados a uno solo.
 *   - COINCIDENCIA POR NOMBRE PARCIAL (COINCIDIR_POR_NOMBRE_PARCIAL; ver bloque
 *     2b de processCsv_ y nombreEsInicioDe_): el CSV trae el nombre completo
 *     ("PEDRAZA CANDIA, LINDAH CHARLOTTE") pero la hoja a veces solo parte de
 *     él ("PEDRAZA | CANDIA | LINDAH"). Para las filas de la hoja SIN
 *     coincidencia exacta se acepta una entrada del CSV si: los apellidos
 *     (apellido1 + apellido2 de la hoja) coinciden exactamente tras normalizar;
 *     los nombres de la hoja son el PRINCIPIO del nombre del CSV (no al revés
 *     ni en otro orden); el candidato es único; no hay duplicados de esa
 *     entrada en el CSV; no coincide ya exactamente con otra fila; y ninguna
 *     otra fila de la hoja la reclama. Si hay ambigüedad no se empareja y la
 *     fila sigue como no encontrada. Al aceptarla, la fila usa la clave del CSV
 *     y el resto de reglas (curso, duplicados, nuevos) se aplican igual. No se
 *     modifican los nombres de la hoja. Se cuenta en el resumen.
 *   - CONFIRMACIÓN MANUAL (CONFIRMAR_MISMOS_APELLIDOS; bloque 2c de
 *     processCsv_): para las filas de la hoja que siguen sin coincidencia
 *     (tras la exacta y la parcial) pero con una entrada del CSV con los MISMOS
 *     apellidos (normalizados) que no coincide ya con otra fila, se muestra un
 *     cuadro Sí/No/Cancelar (los botones de Apps Script no se pueden
 *     renombrar, por eso el texto explica cada uno):
 *        Sí       -> actualiza nombre (A, B, C con el reparto de
 *                    separarApellidos_) y curso
 *        No       -> actualiza solo el curso
 *        Cancelar -> no es el mismo; se prueba el siguiente candidato y, si
 *                    no hay más, la fila queda como no encontrada
 *     Las respuestas solo se aplican en el paso 5 (escritura), así que no se
 *     modifica nada antes. Cubre, por ejemplo, erratas en el nombre
 *     (hoja "YALIANA ALTAGRACIA" frente a CSV "YALIANI ALTAGRACIA").
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
 *     COL_APELLIDO1 y COL_APELLIDO2 queda vacía.
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
 *   - CONTROL DE COINCIDENCIAS: en processCsv_, si tras procesar
 *     UMBRAL_CONTROL_COINCIDENCIAS estudiantes de la hoja (150 por defecto)
 *     ninguno ha coincidido con el CSV, se muestra un aviso (Sí/No) para
 *     continuar o abortar. Se pregunta una sola vez y siempre antes de
 *     escribir nada en la hoja, así que abortar no modifica nada.
 *
 * NOTAS TÉCNICAS PARA FUTURAS MEJORAS
 *   - Las lecturas/escrituras se hacen por bloques (getValues/setValues) por
 *     rendimiento; evitar leer/escribir celda a celda.
 *   - La columna de curso se reescribe completa en el rango de estudiantes
 *     existentes (si tuviera fórmulas, se convertirían en valores).
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
 * Si es true, un estudiante de la hoja cuyo nombre tiene solo parte de los nombres del CSV (ej. hoja "PEDRAZA CANDIA LINDAH"
 * y CSV "PEDRAZA CANDIA, LINDAH CHARLOTTE") se considera el mismo si los apellidos coinciden exactamente, los nombres de la hoja
 * son el principio del nombre del CSV y la coincidencia es única (sin ambigüedad). Solo se aplica a los que no coinciden exactamente.
 */
const COINCIDIR_POR_NOMBRE_PARCIAL = true;

/**
 * Si es true, para cada estudiante de la hoja sin coincidencia pero con los mismos apellidos en el CSV se muestra un aviso
 * para confirmar si es el mismo (actualizar nombre y curso / solo curso / no es el mismo).
 */
const CONFIRMAR_MISMOS_APELLIDOS = true;

/** Si tras procesar este número de estudiantes de la hoja ninguno ha coincidido con el CSV, se avisa y se permite continuar o abortar. */
const UMBRAL_CONTROL_COINCIDENCIAS = 150;


/* ----------------------------------------------------------------------------
 * CONFIGURACIÓN DEL CSV
 * -------------------------------------------------------------------------- */

/** Índice (base 0) del campo "número de orden" dentro de cada línea del CSV (se ignora, solo se valida que sea numérico). */
const CSV_IDX_NUMERO = 0;

/** Índice (base 0) del campo "APELLIDOS , NOMBRE" dentro de cada línea del CSV. */
const CSV_IDX_NOMBRE_COMPLETO = 1;

/** Índice (base 0) del campo "curso" dentro de cada línea del CSV. */
const CSV_IDX_CURSO = 2;

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
    .addItem('Actualizar desde CSV', 'actualizarEstudiantes')
    .addToUi();
}

/** Punto de entrada: avisa de la hoja de destino, pide el CSV y procesa. */
function actualizarEstudiantes() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  // 1. Aviso de la hoja de destino
  const confirmacion = ui.alert(
    'Hoja de destino',
    'Los datos se escribirán en la hoja:\n\n"' + sheetName + '"\n\n¿Quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  // 2. Pedir el CSV
  const respuesta = ui.prompt(
    'Archivo CSV',
    'Sube antes el CSV a Google Drive y pega aquí el enlace de ' +
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
    const resumen = processCsv_(sheet, csvText, archivo);
    Logger.log(resumen);
    ui.alert('Resumen', resumen, ui.ButtonSet.OK);
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
 * Indica si una línea del CSV tiene la estructura esperada:
 *  - al menos 3 campos,
 *  - campo de número de orden formado solo por dígitos,
 *  - campo de nombre con formato "APELLIDOS , NOMBRE" (una coma con texto a ambos lados).
 * El campo del curso puede venir vacío (se marcará como curso no reconocido).
 * @param {string[]} row Campos de la línea ya separados.
 * @return {boolean}
 */
function esLineaCsvValida_(row) {
  const minCampos = Math.max(CSV_IDX_NUMERO, CSV_IDX_NOMBRE_COMPLETO, CSV_IDX_CURSO) + 1;
  if (row.length < minCampos) return false;
  if (!/^\d+$/.test(String(row[CSV_IDX_NUMERO]).trim())) return false;

  const texto = String(row[CSV_IDX_NOMBRE_COMPLETO]);
  const pos = texto.indexOf(',');
  if (pos < 0) return false;
  return collapse_(texto.substring(0, pos)) !== '' && collapse_(texto.substring(pos + 1)) !== '';
}

/** Mensaje de error estándar que explica cómo debe ser el CSV. */
function mensajeEstructuraCsv_(detalle) {
  return 'El CSV no tiene la estructura esperada. No se ha modificado nada.\n\n' +
    'Debe ser un CSV sin fila de encabezado, con una línea por estudiante y 3 campos entre comillas ' +
    'separados por comas:\n\n' +
    '    "18","ORBULESCU , ANA MARÍA","ESO1A"\n\n' +
    '  1) número de orden (solo dígitos; se ignora)\n' +
    '  2) apellidos y nombre separados por una coma\n' +
    '  3) curso\n\n' + detalle;
}

/**
 * Valida la estructura del CSV ya separado en filas. Lanza un Error explicando cómo debe ser
 * si no hay filas o si menos de PORCENTAJE_MIN_LINEAS_VALIDAS de las líneas no vacías son válidas.
 * @param {string[][]} csvRows
 */
function validarEstructuraCsv_(csvRows) {
  const noVacias = [];
  csvRows.forEach(function (row, i) {
    const vacia = row.every(function (c) { return String(c).trim() === ''; });
    if (!vacia) noVacias.push({ row: row, linea: i + 1 });
  });

  if (noVacias.length === 0) {
    throw new Error(mensajeEstructuraCsv_('El archivo no contiene ninguna línea con datos.'));
  }

  const invalidas = noVacias.filter(function (x) { return !esLineaCsvValida_(x.row); });
  const validas = noVacias.length - invalidas.length;

  if (validas / noVacias.length < PORCENTAJE_MIN_LINEAS_VALIDAS) {
    const ejemplos = invalidas.slice(0, 3).map(function (x) {
      const texto = x.row.join(' | ');
      return '  Línea ' + x.linea + ': ' + (texto.length > 70 ? texto.substring(0, 70) + '…' : texto);
    });
    throw new Error(mensajeEstructuraCsv_(
      'Líneas válidas: ' + validas + ' de ' + noVacias.length +
      ' (se exige al menos el ' + Math.round(PORCENTAJE_MIN_LINEAS_VALIDAS * 100) + '%).\n\n' +
      'Ejemplos de líneas incorrectas (campos separados por " | "):\n' + ejemplos.join('\n')
    ));
  }
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
 * Procesa el CSV y actualiza la hoja indicada.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja de destino.
 * @param {string} csvText Contenido del CSV (ya decodificado).
 * @param {{file: GoogleAppsScript.Drive.File, aviso: string}} archivo Datos del archivo origen (para el resumen).
 * @return {string} Resumen para mostrar al usuario.
 */
function processCsv_(sheet, csvText, archivo) {
  const sheetName = sheet.getName();
  const colA1 = colIndex_(COL_APELLIDO1);
  const colA2 = colIndex_(COL_APELLIDO2);
  const colNom = colIndex_(COL_NOMBRE);
  const colCurso = colIndex_(COL_CURSO);
  const colExtra = colIndex_(COL_CURSO_EXTRA);
  const firstRow = HEADER_ROWS + 1;

  // ---------- 1. Leer y parsear el CSV ----------
  const lineasIgnoradas = [];
  const csvEntries = [];                 // en orden del CSV
  const csvMap = new Map();              // key -> [entries]

  let csvRows;
  try {
    csvRows = Utilities.parseCsv(csvText.replace(/^\uFEFF/, ''));
  } catch (e) {
    throw new Error(mensajeEstructuraCsv_('No se ha podido separar el contenido en campos (' + e.message + ').'));
  }
  validarEstructuraCsv_(csvRows);   // lanza un error explicando el formato esperado si no es válido

  csvRows.forEach(function (row, i) {
    const vacia = row.every(function (c) { return String(c).trim() === ''; });
    if (vacia) return;

    if (!esLineaCsvValida_(row)) {
      lineasIgnoradas.push('Línea ' + (i + 1) + ': ' + row.join(','));
      return;
    }

    const nombreCompleto = row[CSV_IDX_NOMBRE_COMPLETO];

    const texto = String(nombreCompleto);
    const posComa = texto.indexOf(',');
    const apellidos = collapse_(posComa >= 0 ? texto.substring(0, posComa) : texto);
    const nombre = collapse_(posComa >= 0 ? texto.substring(posComa + 1) : '');

    const cursoInfo = parsearCurso_(row[CSV_IDX_CURSO]);
    const entry = {
      key: normalize_(texto),
      apeKey: normalize_(apellidos),
      nomKey: normalize_(nombre),
      apellidos: apellidos.toLocaleUpperCase('es'),
      nombre: nombre.toLocaleUpperCase('es'),
      curso: cursoInfo.curso,
      extra: cursoInfo.extra,
      cursoReconocido: cursoInfo.reconocido,
      cursoVacio: cursoInfo.vacio,
      cursoOriginal: collapse_(row[CSV_IDX_CURSO])
    };
    csvEntries.push(entry);
    if (!csvMap.has(entry.key)) csvMap.set(entry.key, []);
    csvMap.get(entry.key).push(entry);
  });

  if (csvEntries.length === 0) {
    throw new Error('El CSV no contiene ningún estudiante válido. No se ha modificado nada.');
  }

  // ---------- 2. Leer estudiantes de la hoja ----------
  const lastSheetRow = sheet.getLastRow();
  const numRows = Math.max(lastSheetRow - firstRow + 1, 0);

  let ap1 = [], ap2 = [], nom = [];
  if (numRows > 0) {
    ap1 = sheet.getRange(firstRow, colA1, numRows, 1).getValues();
    ap2 = sheet.getRange(firstRow, colA2, numRows, 1).getValues();
    nom = sheet.getRange(firstRow, colNom, numRows, 1).getValues();
  }

  // Último estudiante = última fila con algo en apellidos/nombre
  let lastStudentIdx = -1;
  for (let i = 0; i < numRows; i++) {
    if (String(ap1[i][0]).trim() || String(ap2[i][0]).trim() || String(nom[i][0]).trim()) {
      lastStudentIdx = i;
    }
  }
  const numStudents = lastStudentIdx + 1;
  const lastStudentRow = firstRow + lastStudentIdx;     // si no hay estudiantes = firstRow - 1

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

  // ---------- 2b. Coincidencia por nombre parcial ----------
  // La hoja a veces solo tiene el primer nombre y el CSV el nombre completo. Para los estudiantes de la
  // hoja sin coincidencia exacta se busca en el CSV uno con los mismos apellidos cuyo nombre empiece por
  // los nombres de la hoja. Solo se acepta si el candidato es único y ninguna otra fila de la hoja lo
  // reclama. Al aceptarlo, la fila de la hoja pasa a usar la clave del CSV y el resto del proceso
  // (actualizar, duplicados, nuevos) funciona igual que con una coincidencia exacta.
  const csvPorApellidos = new Map();   // apeKey -> [entries]
  csvEntries.forEach(function (e) {
    if (!csvPorApellidos.has(e.apeKey)) csvPorApellidos.set(e.apeKey, []);
    csvPorApellidos.get(e.apeKey).push(e);
  });

  /** Reconstruye el índice de la hoja (clave -> filas) tras cambiar claves. */
  function reconstruirSheetMap() {
    sheetMap.clear();
    for (let i = 0; i < numStudents; i++) {
      if (sheetKeys[i] === '') continue;
      if (!sheetMap.has(sheetKeys[i])) sheetMap.set(sheetKeys[i], []);
      sheetMap.get(sheetKeys[i]).push(i);
    }
  }

  let coincidenciasParciales = 0;
  if (COINCIDIR_POR_NOMBRE_PARCIAL) {
    const candidatoPorFila = new Map();  // índice de fila -> entry candidata
    const reclamos = new Map();          // entry -> nº de filas de la hoja que la reclaman
    for (let i = 0; i < numStudents; i++) {
      const key = sheetKeys[i];
      if (key === '' || csvMap.has(key)) continue;   // fila vacía o ya coincide exactamente

      const apeKey = normalize_(ap1[i][0] + ' ' + ap2[i][0]);
      const nomKey = normalize_(nom[i][0]);
      if (apeKey === '' || nomKey === '') continue;

      const candidatos = (csvPorApellidos.get(apeKey) || []).filter(function (e) {
        return csvMap.get(e.key).length === 1 &&      // sin duplicados en el CSV
               !sheetMap.has(e.key) &&                // no coincide ya exactamente con otra fila
               nombreEsInicioDe_(nomKey, e.nomKey);
      });
      if (candidatos.length === 1) {
        candidatoPorFila.set(i, candidatos[0]);
        reclamos.set(candidatos[0], (reclamos.get(candidatos[0]) || 0) + 1);
      }
    }

    candidatoPorFila.forEach(function (entry, i) {
      if (reclamos.get(entry) !== 1) return;         // reclamada por varias filas: ambigua
      sheetKeys[i] = entry.key;
      coincidenciasParciales++;
    });
    reconstruirSheetMap();
  }

  // ---------- 2c. Confirmación manual: mismos apellidos, nombre distinto ----------
  // Para los estudiantes de la hoja que siguen sin coincidencia pero tienen en el CSV una entrada con los
  // mismos apellidos (y que no coincide ya con otra fila de la hoja), se pregunta al usuario si es el mismo:
  //   Sí (YES)         -> actualizar el nombre en la hoja con el del CSV y el curso
  //   No (NO)          -> actualizar solo el curso (el nombre de la hoja no se toca)
  //   Cancelar / cerrar -> no es el mismo estudiante; se prueba el siguiente candidato (si lo hay) y se continúa
  // Todavía no se ha escrito nada en la hoja: las respuestas solo se aplican en el paso 5.
  let confirmadosNombreYCurso = 0;
  let confirmadosSoloCurso = 0;
  let cambioNombres = false;
  if (CONFIRMAR_MISMOS_APELLIDOS) {
    const ui = SpreadsheetApp.getUi();
    for (let i = 0; i < numStudents; i++) {
      const key = sheetKeys[i];
      if (key === '' || csvMap.has(key)) continue;   // fila vacía o ya emparejada

      const apeKey = normalize_(ap1[i][0] + ' ' + ap2[i][0]);
      if (apeKey === '') continue;

      const candidatos = (csvPorApellidos.get(apeKey) || []).filter(function (e) {
        return csvMap.get(e.key).length === 1;       // sin duplicados en el CSV
      });

      for (let c = 0; c < candidatos.length; c++) {
        const e = candidatos[c];
        if (sheetMap.has(e.key)) continue;           // ya emparejada con otra fila (o confirmada antes)

        const respuesta = ui.alert(
          'Estudiante no encontrado',
          'No se ha encontrado este estudiante, pero he encontrado este con los mismos apellidos en el CSV\n\n' +
          '"' + sheetNames[i] + '" (fila ' + (firstRow + i) + ')\n' +
          '"' + e.apellidos + ', ' + e.nombre + '" (curso en el CSV: ' + (e.cursoOriginal || 'vacío') + ')\n\n' +
          'Necesito que me confirmes si es el mismo estudiante.\n' +
          '(Por límites de configuración la respuesta NO es rara)\n\n' +
          'Sí  =  Sí, actualiza el nombre y el curso\n' +
          'No  =  Este NO es un NO ACTUALICES EL NOMBRE (lo que haré es actualizar solo el curso)\n' +
          'Cancelar  =  Esto es NO ES EL MISMO, continúa ejecutando el script (no actualizo nada)',
          ui.ButtonSet.YES_NO_CANCEL
        );

        if (respuesta === ui.Button.YES || respuesta === ui.Button.NO) {
          sheetKeys[i] = e.key;
          sheetMap.set(e.key, [i]);
          if (respuesta === ui.Button.YES) {
            const sep = separarApellidos_(e.apellidos);
            ap1[i][0] = sep.apellido1;
            ap2[i][0] = sep.apellido2;
            nom[i][0] = e.nombre;
            cambioNombres = true;
            confirmadosNombreYCurso++;
          } else {
            confirmadosSoloCurso++;
          }
          break;
        }
        // Cancelar o cerrar el cuadro: no es el mismo; se prueba el siguiente candidato
      }
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
  csvEntries.forEach(function (entry) {
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
    sheet.getRange(firstRow, colA1, numStudents, 1).setValues(ap1.slice(0, numStudents));
    sheet.getRange(firstRow, colA2, numStudents, 1).setValues(ap2.slice(0, numStudents));
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

  return [
    'Estudiantes en el CSV: ' + csvEntries.length,
    'Actualizados (curso del CSV): ' + actualizados,
    'Coincidencias por nombre parcial (incluidas en las anteriores): ' + coincidenciasParciales,
    'Confirmados por el usuario, nombre y curso actualizados (incluidos en las anteriores): ' + confirmadosNombreYCurso,
    'Confirmados por el usuario, solo curso actualizado (incluidos en las anteriores): ' + confirmadosSoloCurso,
    'Nuevos añadidos: ' + nuevos.length,
    'No encontrados en el CSV: ' + noEncontrados.length,
    'Duplicados: ' + duplicados.size,
    'Líneas del CSV ignoradas: ' + lineasIgnoradas.length,
    'Cursos no reconocidos: ' + cursosNoReconocidos,
    'Estudiantes con el curso vacío: ' + cursosVacios
  ].join('\n');
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
 * Indica si los nombres de la hoja son el principio de los nombres del CSV (y el CSV tiene más).
 * Ambos ya normalizados. Ej.: ("lindah", "lindah charlotte") -> true; ("charlotte", "lindah charlotte") -> false.
 * @param {string} nombreHoja
 * @param {string} nombreCsv
 * @return {boolean}
 */
function nombreEsInicioDe_(nombreHoja, nombreCsv) {
  const wh = nombreHoja.split(' ');
  const wc = nombreCsv.split(' ');
  if (wh.length >= wc.length) return false;   // si tienen las mismas palabras, ya habría coincidido exactamente
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