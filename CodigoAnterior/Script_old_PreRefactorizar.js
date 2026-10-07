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

// COLUMNAS QUE COMPONEN EL NOMBRE COMPLETO
// Se indican en una lista, con letras de columna. El script las une EN ESE ORDEN
// para formar el nombre completo y compararlo (ya no hace falta la columna con fórmula).
//   - Si la pestaña tiene 3 columnas (apellido1, apellido2, nombre):  ["A", "B", "C"]
//   - Si la pestaña tiene 2 columnas (apellidos, nombre):             ["A", "B"]
// IMPORTANTE: usa el mismo orden lógico en todas las pestañas (por ejemplo,
// apellidos primero y nombre al final), aunque unas tengan 2 columnas y otras 3.
// Las celdas vacías (por ejemplo, un segundo apellido en blanco) se ignoran.
// ⚠️ AJUSTA estas letras a tus pestañas reales.
var COLS_NOMBRE_DESTINO = ["I", "J", "K"];   // Pestaña actual (destino)

// Pestañas de ORIGEN (años anteriores). Hay alumnos que han repetido,
// por eso se busca primero en la primera pestaña y, si no aparece, en la siguiente.
// Puedes añadir o quitar tantas pestañas como necesites.
// Cada una pude tener su propia lista de columnas de nombre.
var PESTANAS_ORIGEN = [
  { nombre: "2ESO25", colsNombre: ["A", "B"] },        // Pestaña de origen 1 (año pasado): apellidos + nombre
  { nombre: "3ESO25", colsNombre: ["A", "B"] }    // Pestaña de origen 2 (curso anterior al repetir): apellido1 + apellido2 + nombre
];

var COLUMNA_INICIO_DESTINO = "L"; // Dónde empiezan a pegarse las columnas en la pestaña actual
var CANTIDAD_COLUMNAS = 6;        // Cuántas columnas se van a copiar

var COLUMNA_INICIO_ORIGEN = "C";  // Dónde empiezan los datos a copiar en las pestañas de origen
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
 * Convierte una lista de letras de columna (["A","B"]) en una lista de números ([1,2]).
 */
function letrasANumeros(letras) {
  var numeros = [];
  for (var i = 0; i < letras.length; i++) {
    numeros.push(letraANumero(letras[i]));
  }
  return numeros;
}

/**
 * Normaliza un nombre para comparaciones: quita tildes, pasa a minusculas,
 * colapsa espacios repetidos y recorta espacios sobrantes.
 * Ejemplo: "LOPEZ SANCHEZ LEIRE" y "López  Sánchez Leire" se consideran iguales.
 */
function normalizarNombre(nombre) {
  return nombre
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/**
 * Monta el nombre completo de una fila uniendo, en orden, las columnas indicadas.
 * Las celdas vacías se ignoran. Devuelve el resultado ya normalizado.
 * Ejemplo: fila con [LOPEZ, SANCHEZ, LEIRE] y columnas [1,2,3] -> "lopez sanchez leire"
 */
function construirNombre(fila, columnasNum) {
  var partes = [];
  for (var c = 0; c < columnasNum.length; c++) {
    var valor = String(fila[columnasNum[c] - 1]).trim();
    if (valor !== "") partes.push(valor);
  }
  return normalizarNombre(partes.join(" "));
}

/**
 * Devuelve las pestañas de origen configuradas para los mensajes de log y confirmación.
 */
function nombresOrigenes() {
  var nombres = [];
  for (var k = 0; k < PESTANAS_ORIGEN.length; k++) {
    nombres.push(PESTANAS_ORIGEN[k].nombre + " (nombre en col. " + PESTANAS_ORIGEN[k].colsNombre.join("+") + ")");
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
    "• Pestaña de DESTINO (Escritura): " + nombreAct + " (Nombre en col. " + COLS_NOMBRE_DESTINO.join("+") + ", Columnas de destino: " + COLUMNA_INICIO_DESTINO + " a " + colFinDestLetra + ")\n\n" +
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
  var colsNomActNum = letrasANumeros(COLS_NOMBRE_DESTINO);
  var colDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);

  var ultimaFilaActual = pestanaActual.getLastRow();
  if (ultimaFilaActual < 2) {
    Logger.log("⚠️ La pestaña actual no tiene suficientes filas de datos.");
    return;
  }

  var datosActuales = pestanaActual.getRange(2, 1, ultimaFilaActual - 1, pestanaActual.getLastColumn()).getValues();

  // Leemos TODOS los datos de cada pestaña de origen en memoria (menos llamadas a la hoja = más rápido)
  // y montamos un índice "nombre completo -> posición de la fila" para buscar al instante.
  // Si un nombre aparece repetido en la misma pestaña, se queda con la primera aparición.
  for (var k = 0; k < origenes.length; k++) {
    var o = origenes[k];
    o.colsNomNum = letrasANumeros(o.cfg.colsNombre);
    o.indice = {};
    var ultimaFilaOrigen = o.pestana.getLastRow();
    if (ultimaFilaOrigen < 2) {
      Logger.log("⚠️ La pestaña de origen '" + o.cfg.nombre + "' no tiene suficientes filas de datos. Se ignorará en la búsqueda.");
      o.datos = [];
      continue;
    }
    o.datos = o.pestana.getRange(2, 1, ultimaFilaOrigen - 1, o.pestana.getLastColumn()).getValues();

    for (var j = 0; j < o.datos.length; j++) {
      var claveOrigen = construirNombre(o.datos[j], o.colsNomNum);
      if (claveOrigen !== "" && !Object.prototype.hasOwnProperty.call(o.indice, claveOrigen)) {
        o.indice[claveOrigen] = j;
      }
    }
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
    // Montamos el nombre completo uniendo las columnas configuradas
    // (ya normalizado: sin tildes, en minúsculas y sin espacios sobrantes)
    var nombreCompletoActual = construirNombre(datosActuales[i], colsNomActNum);

    // Si el nombre está vacío, saltamos a la siguiente
    if (nombreCompletoActual === "") continue;

    var encontrado = false;
    var datosEncontrados = [];
    var pestanaOrigenEncontrada = "";

    // Buscamos en TODAS las pestañas de origen, en el orden configurado.
    // En cuanto aparece el alumno en una de ellas, se detiene la búsqueda.
    for (var k = 0; k < origenes.length; k++) {
      var o = origenes[k];

      if (Object.prototype.hasOwnProperty.call(o.indice, nombreCompletoActual)) {
        var filaOrigen = o.datos[o.indice[nombreCompletoActual]];
        encontrado = true;
        pestanaOrigenEncontrada = o.cfg.nombre;
        // Extraemos las columnas correspondientes del origen
        var inicioSlice = colOrigNum - 1;
        datosEncontrados = filaOrigen.slice(inicioSlice, inicioSlice + CANTIDAD_COLUMNAS);
        break; // Ya está en una pestaña anterior, no seguimos buscando
      }
    }

    var filaReal = i + 2; // Fila real en la hoja de cálculo
    var rangoDestino = pestanaActual.getRange(filaReal, colDestNum, 1, CANTIDAD_COLUMNAS);
    // Mostramos el nombre completo montado (con los nombres tal como están escritos) en el log
    var nombreOriginal = [];
    for (var c = 0; c < colsNomActNum.length; c++) {
      var parte = String(datosActuales[i][colsNomActNum[c] - 1]).trim();
      if (parte !== "") nombreOriginal.push(parte);
    }
    nombreOriginal = nombreOriginal.join(" ");

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