// ==========================================
// ⚙️ CONFIGURACIÓN GLOBAL (Modifica aquí fácilmente)
// ==========================================
var NOMBRE_PESTANA_ACTUAL = "2ESO26";     // Cambia por el nombre de tu pestaña de este año
var NOMBRE_PESTANA_ANTERIOR = "1ESO25"; // Cambia por el nombre de tu pestaña del año pasado

// Configuración de columnas (Usa números: A=1, B=2, C=3, D=4, etc.)
var COL_NOMBRE_COMPLETO_ACTUAL = 13;   // Columna del Nombre Completo en la pestaña actual (ej: 1 = Columna A;  13 = Columna M)
var COL_NOMBRE_COMPLETO_ANTERIOR = 3; // Columna del Nombre Completo en la pestaña anterior (ej: 1 = Columna A)

var COLUMNA_INICIO_DESTINO = 22; // Dónde empiezan a pegarse las 7 columnas en la pestaña actual (ej: 2 = Columna B; 22 = Columna a V)
var CANTIDAD_COLUMNAS = 6;      // Cuántas columnas se van a copiar (7 en tu caso)

var COLUMNA_INICIO_ORIGEN = 4;  // Dónde empiezan los datos a copiar en la pestaña anterior (ej: 2 = Columna B)
// ==========================================

// Función auxiliar para convertir letras de columna (ej: "A" -> 1, "B" -> 2, etc.)
function letraANumero(letra) {
  letra = letra.toUpperCase();
  var suma = 0;
  for (var i = 0; i < letra.length; i++) {
    suma *= 26;
    suma += letra.charCodeAt(i) - 64;
  }
  return suma;
}

/**
 * Función independiente para mostrar la advertencia y pedir confirmación al usuario.
 * Retorna true si el usuario acepta, o false si cancela.
 */
function confirmarEjecucion(nombreAnt, nombreAct) {
  // Calculamos las letras de fin de rango para mostrarlas claramente en el mensaje
  var colInicioDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colFinDestNum = colInicioDestNum + CANTIDAD_COLUMNAS - 1;
  var colFinDestLetra = SpreadsheetApp.getActiveSpreadsheet().getRange(1, colFinDestNum).getA1Notation().replace(/[0-9]/g, '');

  var colInicioOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);
  var colFinOrigNum = colInicioOrigNum + CANTIDAD_COLUMNAS - 1;
  var colFinOrigLetra = SpreadsheetApp.getActiveSpreadsheet().getRange(1, colFinOrigNum).getA1Notation().replace(/[0-9]/g, '');

  // 💬 Mensaje de confirmación detallado
  var mensaje = "⚠️ ADVERTENCIA DE SOBRESCRITURA ⚠️\n\n" +
    "Estás a punto de ejecutar la importación de datos:\n\n" +
    "• Pestaña de ORIGEN (Lectura): " + nombreAnt + " (Columna de búsqueda: " + COL_NOMBRE_COMPLETO_ANTERIOR + ", Columnas de datos: " + COLUMNA_INICIO_ORIGEN + " a " + colFinOrigLetra + ")\n" +
    "• Pestaña de DESTINO (Escritura): " + nombreAct + " (Columna de búsqueda: " + COL_NOMBRE_COMPLETO_ACTUAL + ", Columnas de destino: " + COLUMNA_INICIO_DESTINO + " a " + colFinDestLetra + ")\n\n" +
    "❗ ATENCIÓN: Los valores actuales en el rango de destino de la pestaña '" + nombreAct + "' SE PERDERÁN Y SERÁN SOBRESCRITOS.\n\n" +
    "¿Deseas continuar?";

  var respuesta = Browser.msgBox("Confirmación de Proceso", mensaje, Browser.Buttons.YES_NO);

  if (respuesta !== "yes") {
    SpreadsheetApp.getActiveSpreadsheet().toast("Proceso cancelado por el usuario.", "Cancelado", 3);
    Logger.log("Proceso cancelado por el usuario.");
    return false;
  }

  return true;
}

// Función principal que ejecuta la búsqueda y el copiado
function buscarYCopiarAlumnos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var pestanaActual = sheet.getSheetByName(NOMBRE_PESTANA_ACTUAL);
  var pestanaAnterior = sheet.getSheetByName(NOMBRE_PESTANA_ANTERIOR);

  if (!pestanaActual || !pestanaAnterior) {
    var errorMsg = "❌ Error: No se encuentra alguna de las pestañas especificadas. Revisa los nombres en la configuración.";
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  var ultimaFilaActual = pestanaActual.getLastRow();
  if (ultimaFilaActual < 2) {
    Logger.log("⚠️ La pestaña actual no tiene suficientes filas de datos.");
    return;
  }

  // Obtenemos todos los datos de ambas pestañas
  var datosActuales = pestanaActual.getRange(2, 1, ultimaFilaActual - 1, pestanaActual.getLastColumn()).getValues();

  var ultimaFilaAnterior = pestanaAnterior.getLastRow();
  if (ultimaFilaAnterior < 2) {
    Logger.log("⚠️ La pestaña anterior no tiene suficientes filas de datos.");
    return;
  }
  var datosAnteriores = pestanaAnterior.getRange(2, 1, ultimaFilaAnterior - 1, pestanaAnterior.getLastColumn()).getValues();

  Logger.log("----------------------------------------");
  Logger.log("INICIANDO BÚSQUEDA Y COPIA DE ALUMNOS...");
  Logger.log("Total alumnos a procesar: " + datosActuales.length);
  Logger.log("----------------------------------------");

  var encontradosCount = 0;
  var noEncontradosCount = 0;

  // Recorremos cada alumno de este año
  for (var i = 0; i < datosActuales.length; i++) {
    // Normalizamos el texto (pasamos a minúsculas y quitamos espacios sobrantes para evitar fallos por tildes o espacios)
    var nombreCompletoActual = String(datosActuales[i][COL_NOMBRE_COMPLETO_ACTUAL - 1]).trim().toLowerCase();

    // Si la celda del nombre está vacía, saltamos a la siguiente
    if (nombreCompletoActual === "") continue;

    var encontrado = false;
    var datosEncontrados = [];

    // Buscamos en el año pasado
    for (var j = 0; j < datosAnteriores.length; j++) {
      var nombreCompletoAnt = String(datosAnteriores[j][COL_NOMBRE_COMPLETO_ANTERIOR - 1]).trim().toLowerCase();

      if (nombreCompletoActual === nombreCompletoAnt) {
        encontrado = true;
        // Extraemos las 6 columnas correspondientes del año pasado
        var inicioSlice = COLUMNA_INICIO_ORIGEN - 1;
        datosEncontrados = datosAnteriores[j].slice(inicioSlice, inicioSlice + CANTIDAD_COLUMNAS);
        break;
      }
    }

    var filaReal = i + 2; // Fila real en la hoja de cálculo
    var rangoDestino = pestanaActual.getRange(filaReal, COLUMNA_INICIO_DESTINO, 1, CANTIDAD_COLUMNAS);

    // Mostramos el nombre original (sin pasar a minúsculas) en el log para que sea legible
    var nombreOriginal = datosActuales[i][COL_NOMBRE_COMPLETO_ACTUAL - 1];

    if (encontrado) {
      // Si se encuentra: copiamos los valores y limpiamos alertas visuales previas
      rangoDestino.setValues([datosEncontrados]);
      // Esto permite colorear el fondo
      // rangoDestino.setBackground(null);
      Logger.log("✅ [ENCONTRADO] Fila " + filaReal + ": " + nombreOriginal);
      encontradosCount++;
    } else {
      // Si NO se encuentra: limpiamos contenido, marcamos en rojo claro y avisamos en la primera celda
      rangoDestino.clearContent();
      // Esto permite colorear el fondo de rojo, pero luego habría que estar quitando el color
      // rangoDestino.setBackground("#f4cccc"); // Rojo claro de aviso
      pestanaActual.getRange(filaReal, COLUMNA_INICIO_DESTINO).setValue("NO ENCONTRADO");
      Logger.log("❌ [NO ENCONTRADO] Fila " + filaReal + ": " + nombreOriginal + " -> Marcado en rojo.");
      noEncontradosCount++;
    }
  }

  Logger.log("----------------------------------------");
  Logger.log("RESUMEN FINAL:");
  Logger.log("- Alumnos encontrados y copiados: " + encontradosCount);
  Logger.log("- Alumnos no encontrados (marcados): " + noEncontradosCount);
  Logger.log("----------------------------------------");

  SpreadsheetApp.getUi().alert("¡Búsqueda completada!\n\nEncontrados: " + encontradosCount + "\nNo encontrados: " + noEncontradosCount);
}