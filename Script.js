// ==========================================
// 🤖 CONTEXTO PARA ASISTENTES DE IA (leer antes de modificar este archivo)
// ==========================================
// Este es el ÚNICO archivo del proyecto: Google Apps Script vinculado a una
// hoja de cálculo de Google (lenguaje: JavaScript de Apps Script, estilo "var").
// Sin dependencias externas. Se ejecuta a mano desde el editor de Apps Script.
// Idioma de comentarios, logs y mensajes al usuario: ESPAÑOL (mantenerlo).
//
// PROPÓSITO
//   Migrar datos de alumnos de pestañas de cursos anteriores a la pestaña del
//   curso actual. Para cada alumno de la pestaña de DESTINO se busca su nombre
//   completo en las pestañas de ORIGEN y se copia un bloque de columnas.
//
// FLUJO (función principal: buscarYCopiarAlumnos)
//   1. Valida que existan la pestaña de destino y TODAS las de origen.
//   2. Pide confirmación con ui.alert (confirmarEjecucion): avisa de que el rango
//      de destino se SOBRESCRIBE. Si el usuario no pulsa YES, se cancela.
//   3. Lee en memoria todas las filas (getValues, desde la fila 2; la fila 1 es
//      la cabecera). Evitar llamadas a la hoja dentro de bucles: son muy lentas.
//   4. Para cada pestaña de origen crea un índice { nombre normalizado -> fila }
//      (si hay nombres repetidos en una misma pestaña, gana la primera aparición).
//   5. Para cada alumno del destino monta su nombre completo, lo busca en los
//      orígenes EN EL ORDEN CONFIGURADO y se detiene en el primer acierto
//      (hay alumnos repetidores: primero se mira el año pasado, luego el anterior).
//   6. Encontrado     -> copia CANTIDAD_COLUMNAS columnas con setValues en el destino.
//      No encontrado  -> limpia ese rango (clearContent) y escribe "NO ENCONTRADO"
//                        en su primera celda para revisión manual.
//   7. Registra cada alumno con Logger.log y muestra un resumen final con alert.
//
// CÓMO SE COMPARAN LOS NOMBRES (importante)
//   - No existe columna con fórmula: el nombre completo se MONTA en el script
//     uniendo, con espacios y EN EL ORDEN DE LA LISTA, las columnas indicadas en
//     COLS_NOMBRE_DESTINO y en "colsNombre" de cada pestaña de origen.
//   - Cada pestaña puede tener 2 columnas (apellidos, nombre) o 3 (apellido1,
//     apellido2, nombre). El orden lógico debe ser el mismo en todas las pestañas
//     (apellidos primero, nombre al final); si no, no habrá coincidencias.
//   - construirNombre ignora celdas vacías y normalizarNombre quita tildes,
//     pasa a minúsculas y colapsa/recorta espacios, así "LÓPEZ  Sánchez Leire" ==
//     "lopez sanchez leire".
//   - Las filas del destino con nombre vacío se saltan sin tocar nada.
//
// CONFIGURACIÓN (bloque "CONFIGURACIÓN GLOBAL", justo debajo)
//   - Todas las columnas se indican con LETRAS ("A", "B"...) y se convierten a
//     números con letraANumero / letrasANumeros. No escribir índices numéricos.
//   - NOMBRE_PESTANA_DESTINO, COLS_NOMBRE_DESTINO, PESTANAS_ORIGEN (lista de
//     { nombre, colsNombre }), COLUMNA_INICIO_DESTINO, COLUMNA_INICIO_ORIGEN y
//     CANTIDAD_COLUMNAS. El bloque origen y el bloque destino tienen el mismo
//     tamaño (CANTIDAD_COLUMNAS); la columna de inicio de origen es la misma
//     en todas las pestañas de origen.
//   - Para cambiar de curso normalmente solo se tocan los nombres de pestañas.
//
// NOTAS Y DETALLES A TENER EN CUENTA
//   - Las líneas de colores (setBackground con "#f4cccc") están COMENTADAS a
//     propósito: se evita tener que quitar el color a mano después. Los logs
//     dicen "Marcado en rojo", pero ahora lo único visible es el texto
//     "NO ENCONTRADO" en la primera celda del bloque destino.
//   - Convención de nombres: todo lo que lleva "Destino" (pestanaDestino, datosDestino,
//     nombreCompletoDestino...) se refiere a la pestaña del curso actual, que es la que
//     se escribe. Las pestañas de origen están en el array "origenes" (o.pestana).
//   - Al añadir una opción de configuración, hay que reflejarla también en el
//     mensaje de confirmarEjecucion y en el log inicial.
//   - Mantener las comprobaciones previas y la confirmación: el script borra y
//     sobrescribe datos, así que no debe ejecutarse nada destructivo sin avisar.
// ==========================================






// ==========================================
// ==========================================
// ==========================================
// ==========================================
//
// INFO PARA EL HUMANO QUE EJECUTA ESTE SCRIPT:
//
// Para usar este script primero tienes que rellenar la CONFIGURACIÓN GLOBAL
// luego das a Guardar (el disquete que aparece un poco más arriba)
// Cuando lo hayas hecho das a Ejecutar, vas a la pestaña de Chrome con la hoja de cálculo
// y verás una ventana de confirmación donde te explica lo que va a pasar
//
//
// ==========================================
// ==========================================
// ==========================================
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
var COLS_NOMBRE_DESTINO = ["I", "J", "K"];   // Pestaña de destino (este año)

// Pestañas de ORIGEN (años anteriores). Hay alumnos que han repetido,
// por eso se busca primero en la primera pestaña y, si no aparece, en la siguiente.
// Puedes añadir o quitar tantas pestañas como necesites.
// Cada una pude tener su propia lista de columnas de nombre.
var PESTANAS_ORIGEN = [
  { nombre: "2ESO25", colsNombre: ["A", "B"] },        // Pestaña de origen 1 (año pasado): apellidos + nombre
  { nombre: "3ESO25", colsNombre: ["A", "B"] }    // Pestaña de origen 2 (curso anterior al repetir): apellido1 + apellido2 + nombre
];

var COLUMNA_INICIO_DESTINO = "L"; // Dónde empiezan a pegarse las columnas en la pestaña de destino
var CANTIDAD_COLUMNAS = 6;        // Cuántas columnas se van a copiar

var COLUMNA_INICIO_ORIGEN = "C";  // Dónde empiezan los datos a copiar en las pestañas de origen
// ==========================================




// ==========================================
// FUNCIONES AUXILIARES
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
function confirmarEjecucion(nombreDestino, pestanaDestino) {
  var ui = SpreadsheetApp.getUi();

  // Calculamos las letras de fin de rango para mostrarlas claramente en el mensaje
  var colInicioDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colFinDestNum = colInicioDestNum + CANTIDAD_COLUMNAS - 1;
  // Usamos la pestaña activa para obtener la anotación A1 correctamente
  var colFinDestLetra = pestanaDestino.getRange(1, colFinDestNum).getA1Notation().replace(/[0-9]/g, '');

  var colInicioOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);
  var colFinOrigNum = colInicioOrigNum + CANTIDAD_COLUMNAS - 1;
  var colFinOrigLetra = pestanaDestino.getRange(1, colFinOrigNum).getA1Notation().replace(/[0-9]/g, '');

  // 💬 Mensaje de confirmación detallado con saltos de línea limpios
  var mensaje = "⚠️ ADVERTENCIA DE SOBRESCRITURA ⚠️\n\n\n" +
    "Estás a punto de ejecutar la importación de datos:\n\n" +
    "• Pestañas de ORIGEN (Lectura, se busca en este orden): \n  " + nombresOrigenes() +
    "   [Columnas con datos en ambas: " + COLUMNA_INICIO_ORIGEN + " a " + colFinOrigLetra + "]\n" +
    "• Pestaña de DESTINO (Escritura): " + nombreDestino + " (Nombre en col. " + COLS_NOMBRE_DESTINO.join("+") + ", Columnas de destino: " + COLUMNA_INICIO_DESTINO + " a " + colFinDestLetra + ")\n\n" +
    "❗ ATENCIÓN: Los valores actuales en el rango de destino de la pestaña '" + nombreDestino + "' SE PERDERÁN Y SERÁN SOBRESCRITOS.\n\n" +
    "¿Deseas continuar?" +
    "\n\n (NOTA: si das a Sí, ten paciencia, esta ventana se cierra y en unos segundos se ve el resultado)";

  var respuesta = ui.alert("Confirmación de Proceso", mensaje, ui.ButtonSet.YES_NO);

  if (respuesta !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Proceso cancelado por el usuario.", "Cancelado", 3);
    Logger.log("Proceso cancelado por el usuario.");
    return false;
  }

  return true;
}






// ==========================================
//
//
// Función principal que ejecuta la búsqueda y el copiado
//
//
//
// ==========================================
function buscarYCopiarAlumnos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var pestanaDestino = sheet.getSheetByName(NOMBRE_PESTANA_DESTINO);

  if (!pestanaDestino) {
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

  // Llamamos a la función independiente de confirmación pasándole la pestaña de destino
  Logger.log("\nAhora necesitas responder en la pestaña del chrome donde está la hoja de cálculo.");
  var continuar = confirmarEjecucion(NOMBRE_PESTANA_DESTINO, pestanaDestino);
  if (!continuar) {
    return; // Si el usuario dice que no, el script se detiene aquí
  }

  // Convertimos las letras configuradas a números de índice para el script
  var colsNomDestinoNum = letrasANumeros(COLS_NOMBRE_DESTINO);
  var colDestNum = letraANumero(COLUMNA_INICIO_DESTINO);
  var colOrigNum = letraANumero(COLUMNA_INICIO_ORIGEN);

  var ultimaFilaDestino = pestanaDestino.getLastRow();
  if (ultimaFilaDestino < 2) {
    Logger.log("⚠️ La pestaña de destino no tiene suficientes filas de datos.");
    return;
  }

  var datosDestino = pestanaDestino.getRange(2, 1, ultimaFilaDestino - 1, pestanaDestino.getLastColumn()).getValues();

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
  Logger.log("Total alumnos a procesar: " + datosDestino.length);
  Logger.log("----------------------------------------");

  var encontradosCount = 0;
  var noEncontradosCount = 0;

  // Recorremos cada alumno de este año
  for (var i = 0; i < datosDestino.length; i++) {
    // Montamos el nombre completo uniendo las columnas configuradas
    // (ya normalizado: sin tildes, en minúsculas y sin espacios sobrantes)
    var nombreCompletoDestino = construirNombre(datosDestino[i], colsNomDestinoNum);

    // Si el nombre está vacío, saltamos a la siguiente
    if (nombreCompletoDestino === "") continue;

    var encontrado = false;
    var datosEncontrados = [];
    var pestanaOrigenEncontrada = "";

    // Buscamos en TODAS las pestañas de origen, en el orden configurado.
    // En cuanto aparece el alumno en una de ellas, se detiene la búsqueda.
    for (var k = 0; k < origenes.length; k++) {
      var o = origenes[k];

      if (Object.prototype.hasOwnProperty.call(o.indice, nombreCompletoDestino)) {
        var filaOrigen = o.datos[o.indice[nombreCompletoDestino]];
        encontrado = true;
        pestanaOrigenEncontrada = o.cfg.nombre;
        // Extraemos las columnas correspondientes del origen
        var inicioSlice = colOrigNum - 1;
        datosEncontrados = filaOrigen.slice(inicioSlice, inicioSlice + CANTIDAD_COLUMNAS);
        break; // Ya está en una pestaña anterior, no seguimos buscando
      }
    }

    var filaReal = i + 2; // Fila real en la hoja de cálculo
    var rangoDestino = pestanaDestino.getRange(filaReal, colDestNum, 1, CANTIDAD_COLUMNAS);
    // Mostramos el nombre completo montado (con los nombres tal como están escritos) en el log
    var nombreOriginal = [];
    for (var c = 0; c < colsNomDestinoNum.length; c++) {
      var parte = String(datosDestino[i][colsNomDestinoNum[c] - 1]).trim();
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
      pestanaDestino.getRange(filaReal, colDestNum).setValue("NO ENCONTRADO");
      Logger.log("❌ [NO ENCONTRADO] Fila " + filaReal + ": " + nombreOriginal + " -> Marcado en rojo.");
      noEncontradosCount++;
    }
  }

  Logger.log("----------------------------------------");
  Logger.log("RESUMEN FINAL:");
  Logger.log("- Alumnos encontrados y copiados: " + encontradosCount);
  Logger.log("- Alumnos no encontrados (marcados): " + noEncontradosCount);
  Logger.log("----------------------------------------");

  SpreadsheetApp.getUi().alert("¡Búsqueda completada!\n\nEncontrados: " + encontradosCount + "\nNo encontrados: " + noEncontradosCount + "\n\n (NOTA: Ten paciencia, una vez cerrada esta ventana tarda un par de segundos en verse los valores)");
}