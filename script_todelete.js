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
var NOMBRE_PESTANA_DESTINO = "3ESO26";     // Pestaña de destino (este año)

// Pestañas de ORIGEN (años anteriores). Hay alumnos que han repetido,
// por eso se busca primero en la primera pestaña y, si no aparece, en la siguiente.
// Puedes añadir o quitar tantas pestañas como necesites.
var PESTANAS_ORIGEN = [
  { nombre: "2ESO25", colNombre: "C" },   // Pestaña de origen 1 (año pasado)
  { nombre: "3ESO25", colNombre: "C" }    // Pestaña de origen 2 (curso anterior al repetir)
];

// Configuración usando LETRAS de columnas (ej: "A", "B", "C"...)
var COL_NOMBRE_COMPLETO_ACTUAL = "L";   // Columna del Nombre Completo en la pestaña actual

var COLUMNA_INICIO_DESTINO = "M"; // Dónde empiezan a pegarse las columnas en la pestaña actual
var CANTIDAD_COLUMNAS = 6;        // Cuántas columnas se van a copiar

var COLUMNA_INICIO_ORIGEN = "D";  // Dónde empiezan los datos a copiar en las pestañas de origen
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
 * Normaliza un nombre para comparaciones: quita tildes, pasa a minusculas
 * y recorta espacios sobrantes.
 * Ejemplo: "LOPEZ SANCHEZ LEIRE" y "LOPEZ SANCHEZ LEIRE" se consideran iguales.
 */
function normalizarNombre(nombre) {
  return nombre
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

/**
 * Devuelve las pestañas de origen configuradas para los mensajes de log y confirmación.
 */
function nombresOrigenes() {
  var nombres = [];
  for (var k = 0; k < PESTANAS_ORIGEN.length; k++) {
    nombres.push(PESTANAS_ORIGEN[k].nombre + " (col. " + PESTANAS_ORIGEN[k].colNombre + ")");
  }
  return nombres.join(" → ");
}

/**
 * Función independiente para mostrar la advertencia y pedir confirmación al usuario.
 * Utiliza Ui.alert para respetar correctamente los saltos de línea (\n).
 */
function confirmarEjecucion(nombreAct, pestanaActual) {
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
    "• Pestañas de ORIGEN (Lectura, se busca en este orden): \n  " + nombresOrigenes() +
    "   [Columnas con datos en ambas: " + COLUMNA_INICIO_ORIGEN + " a " + colFinOrigLetra + "]\n" +
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

//
//
// Función principal que ejecuta la búsqueda y el copiado
//
//
function buscarYCopiarAlumnos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var pestanaActual = sheet.getSheetByName(NOMBRE_PESTANA_DESTINO);

  if (!pestanaActual) {
    var errorMsg = "❌ Error: No se encuentra la pestaña de destino '" + NOMBRE_PESTANA_DESTINO + "'. Revisa los nombres en la configuración.";
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  // Validamos y cargamos TODAS las pestañas de origen configuradas
  var origenes = [];
  for (var k = 0; k < PESTANAS_ORIGEN.length; k++) {
    var cfg = PESTANAS_ORIGEN[k];
    var pestanaOrigen = sheet.getSheetByName(cfg.nombre);
    if (!pestanaOrigen) {
      var msg = "❌ Error: No se encuentra la pestaña de origen '" + cfg.nombre + "'. Revisa los nombres en la configuración.";
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }
    origenes.push({ cfg: cfg, pestana: pestanaOrigen });
  }

  // Llamamos a la función independiente de confirmación pasándole la pestaña actual
  Logger.log("\nAhora necesitas responder en la pestaña del chrome donde está la hoja de cálculo.");
  var continuar = confirmarEjecucion(NOMBRE_PESTANA_DESTINO, pestanaActual);
  if (!continuar) {
    return; // Si el usuario dice que no, el script se detiene aquí
  }

  // Convertimos las letras configuradas a números de índice para el script
  var colNomActNum = letraANumero(COL_NOMBRE_COMPLETO_ACTUAL);
  var colDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);

  var ultimaFilaActual = pestanaActual.getLastRow();
  if (ultimaFilaActual < 2) {
    Logger.log("⚠️ La pestaña actual no tiene suficientes filas de datos.");
    return;
  }

  var datosActuales = pestanaActual.getRange(2, 1, ultimaFilaActual - 1, pestanaActual.getLastColumn()).getValues();

  // Leemos TODOS los datos de cada pestaña de origen en memoria (menos llamadas a la hoja = más rápido)
  for (var k = 0; k < origenes.length; k++) {
    var o = origenes[k];
    o.colNomNum = letraANumero(o.cfg.colNombre);
    var ultimaFilaOrigen = o.pestana.getLastRow();
    if (ultimaFilaOrigen < 2) {
      Logger.log("⚠️ La pestaña de origen '" + o.cfg.nombre + "' no tiene suficientes filas de datos. Se ignorará en la búsqueda.");
      o.datos = [];
      continue;
    }
    o.datos = o.pestana.getRange(2, 1, ultimaFilaOrigen - 1, o.pestana.getLastColumn()).getValues();
  }

  Logger.log("----------------------------------------");
  Logger.log("INICIANDO BÚSQUEDA Y COPIA DE ALUMNOS...");
  Logger.log("Pestañas de origen (en este orden): " + nombresOrigenes());
  Logger.log("Total alumnos a procesar: " + datosActuales.length);
  Logger.log("----------------------------------------");

  var encontradosCount = 0;
  var noEncontradosCount = 0;

  // Recorremos cada alumno de este año
  for (var i = 0; i < datosActuales.length; i++) {
    // Normalizamos el texto (pasamos a minúsculas y quitamos espacios sobrantes para evitar fallos por tildes o espacios)
    var nombreCompletoActual = normalizarNombre(String(datosActuales[i][colNomActNum - 1]));

    // Si la celda del nombre está vacía, saltamos a la siguiente
    if (nombreCompletoActual === "") continue;

    var encontrado = false;
    var datosEncontrados = [];
    var pestanaOrigenEncontrada = "";

    // Buscamos en TODAS las pestañas de origen, en el orden configurado.
    // En cuanto aparece el alumno en una de ellas, se detiene la búsqueda.
    for (var k = 0; k < origenes.length; k++) {
      var o = origenes[k];
      for (var j = 0; j < o.datos.length; j++) {
        var nombreCompletoOrig = normalizarNombre(String(o.datos[j][o.colNomNum - 1]));

        if (nombreCompletoActual === nombreCompletoOrig) {
          encontrado = true;
          pestanaOrigenEncontrada = o.cfg.nombre;
          // Extraemos las columnas correspondientes del origen
          var inicioSlice = colOrigNum - 1;
          datosEncontrados = o.datos[j].slice(inicioSlice, inicioSlice + CANTIDAD_COLUMNAS);
          break;
        }
      }
      if (encontrado) break; // Ya está en una pestaña anterior, no seguimos buscando
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
      Logger.log("✅ [ENCONTRADO] Fila " + filaReal + ": " + nombreOriginal + " (en '" + pestanaOrigenEncontrada + "')");
      encontradosCount++;
    } else {
      // Si NO se encuentra en NINGUNA pestaña: limpiamos contenido, marcamos en rojo claro y avisamos en la primera celda
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