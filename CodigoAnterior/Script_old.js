// ==========================================
// INFO:
//
// Para usar este script primero tienes que rellenar la CONFIGURACIÓN GLOBAL
// luego das a Guardar (el disquete que aparece un poco más arriba)
// Cuando lo hayas hecho das a Ejecutar, vas a la pestaña de Chrome con la hoja de cálculo
// y verás una ventana de confirmación donde te explica lo que va a pasar
//
//
// ==========================================

// ==========================================
// ⚙️ CONFIGURACIÓN GLOBAL (Modifica aquí fácilmente)
// ==========================================
var NOMBRE_PESTANA_ACTUAL = "2ESO26";     // Pestaña de destino (este año)
var NOMBRE_PESTANA_ANTERIOR = "1ESO25"; // Pestaña de origen (año pasado)

// Configuración usando LETRAS de columnas (ej: "A", "B", "C"...)
var COL_NOMBRE_COMPLETO_ACTUAL = "M";   // Columna del Nombre Completo en la pestaña actual
var COL_NOMBRE_COMPLETO_ANTERIOR = "C"; // Columna del Nombre Completo en la pestaña anterior

var COLUMNA_INICIO_DESTINO = "U"; // Dónde empiezan a pegarse las 7 columnas en la pestaña actual
var CANTIDAD_COLUMNAS = 6;        // Cuántas columnas se van a copiar

var COLUMNA_INICIO_ORIGEN = "D";  // Dónde empiezan los datos a copiar en la pestaña anterior
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
 * Utiliza Ui.alert para respetar correctamente los saltos de línea (\n).
 */
function confirmarEjecucion(nombreAnt, nombreAct, pestanaActual) {
  var ui = SpreadsheetApp.getUi();

  // Calculamos las letras de fin de rango para mostrarlas claramente en el mensaje
  var colInicioDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colFinDestNum = colInicioDestNum + CANTIDAD_COLUMNAS - 1;
  // Usamos la pestaña activa para obtener la anotación A1 correctamente
  var colFinDestLetra = pestanaActual.getRange(1, colFinDestNum).getA1Notation().replace(/[0-9]/g, '');

  var colInicioOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);
  var colFinOrigNum = colInicioOrigNum + CANTIDAD_COLUMNAS - 1;
  var colFinOrigLetra = pestanaActual.getRange(1, colFinOrigNum).getA1Notation().replace(/[0-9]/g, '');

  // 💬 Mensaje de confirmación detallado con saltos de línea limpios
  var mensaje = "⚠️ ADVERTENCIA DE SOBRESCRITURA ⚠️\n\n\n" +
    "Estás a punto de ejecutar la importación de datos:\n\n" +
    "• Pestaña de ORIGEN (Lectura): " + nombreAnt + " (Columna de búsqueda: " + COL_NOMBRE_COMPLETO_ANTERIOR + ", Columnas de datos: " + COLUMNA_INICIO_ORIGEN + " a " + colFinOrigLetra + ")\n" +
    "• Pestaña de DESTINO (Escritura): " + nombreAct + " (Columna de búsqueda: " + COL_NOMBRE_COMPLETO_ACTUAL + ", Columnas de destino: " + COLUMNA_INICIO_DESTINO + " a " + colFinDestLetra + ")\n\n" +
    "❗ ATENCIÓN: Los valores actuales en el rango de destino de la pestaña '" + nombreAct + "' SE PERDERÁN Y SERÁN SOBRESCRITOS.\n\n" +
    "¿Deseas continuar?";

  var respuesta = ui.alert("Confirmación de Proceso", mensaje, ui.ButtonSet.YES_NO);

  if (respuesta !== ui.Button.YES) {
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

  // Llamamos a la función independiente de confirmación pasándole la pestaña actual
  var continuar = confirmarEjecucion(NOMBRE_PESTANA_ANTERIOR, NOMBRE_PESTANA_ACTUAL, pestanaActual);
  if (!continuar) {
    return; // Si el usuario dice que no, el script se detiene aquí
  }

  // Convertimos las letras configuradas a números de índice para el script
  var colNomActNum = letraANumero(COL_NOMBRE_COMPLETO_ACTUAL);
  var colNomAntNum = letraANumero(COL_NOMBRE_COMPLETO_ANTERIOR);
  var colDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);

  var ultimaFilaActual = pestanaActual.getLastRow();
  if (ultimaFilaActual < 2) {
    Logger.log("⚠️ La pestaña actual no tiene suficientes filas de datos.");
    return;
  }

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
    var nombreCompletoActual = String(datosActuales[i][colNomActNum - 1]).trim().toLowerCase();

    // Si la celda del nombre está vacía, saltamos a la siguiente
    if (nombreCompletoActual === "") continue;

    var encontrado = false;
    var datosEncontrados = [];

    // Buscamos en el año pasado
    for (var j = 0; j < datosAnteriores.length; j++) {
      var nombreCompletoAnt = String(datosAnteriores[j][colNomAntNum - 1]).trim().toLowerCase();

      if (nombreCompletoActual === nombreCompletoAnt) {
        encontrado = true;
        // Extraemos las 6 columnas correspondientes del año pasado
        var inicioSlice = colOrigNum - 1;
        datosEncontrados = datosAnteriores[j].slice(inicioSlice, inicioSlice + CANTIDAD_COLUMNAS);
        break;
      }
    }

    var filaReal = i + 2; // Fila real en la hoja de cálculo
    var rangoDestino = pestanaActual.getRange(filaReal, colDestNum, 1, CANTIDAD_COLUMNAS);
    // Mostramos el nombre original (sin pasar a minúsculas) en el log para que sea legible
    var nombreOriginal = datosActuales[i][colNomActNum - 1];

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
      pestanaActual.getRange(filaReal, colDestNum).setValue("NO ENCONTRADO");
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