Esto es un script para Google Spreadsheets que me pidió Eva para copiar datos de estudiantes de un curso a otro.
Los estudiantes están en diferentes pestañas.
<br>
<br>

Para poder usarlo:
- Abre el documento de excel donde quieras ejecutarlo.
- Ve al menu Extensiones y selecciona Apps Scripts
  - Esto abre una pestaña nueva
- Sustituye el codigo que viene en el archivo que se ha abierto por el contenido de `Script.js`
- Cambia el nombre del *Proyecto sin titulo* por algo más descriptivo tipo *Copiar datos de estudiantes*
- Da a Guardar (Disquete que sale)
- Luego das a Ejecutar y eso abre una ventana de confirmación en la hoja de cálculo.\
- Te dará un aviso de permisos, tienes que permitirlo para poder ejecutarlo.
<br>


> **Importante**\
> Para que funcione a la primera debes configurar la seccion `CONFIGURACIÓN GLOBAL` dentro del codigo, buscalo por debajo de las instrucciones de IA.\
> \
> Hay comentarios de que rellenar, pero dejo aqui unas pistas
> **NOMBRE_PESTANA_DESTINO** = "3ESO26";     // Pestaña donde se copiaran los datos\
> **COLS_NOMBRE_DESTINO** = ["I", "J", "K"];   // Columnas que contienen los apellidos y el nombre\
> **PESTANAS_ORIGEN**: aqui ponemos el nombre de las pestañas donde buscará los estudiantes y las columnas con los apellidos y nombre de cada pestaña. (Se pueden poner 2 pestañas porque hay estudiantes repetidores)\
> **COLUMNA_INICIO_DESTINO** = "L"; // Columnas donde empiezan a pegarse los datos\
> **CANTIDAD_COLUMNAS** = 6; // Cuántas columnas se van a copiar\
> **COLUMNA_INICIO_ORIGEN** = "C"; // Primera columnas con los datos a copiar