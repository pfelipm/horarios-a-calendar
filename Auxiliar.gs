/************************
 * Funciones auxiliares *
 ************************/


/**
 * Crea el menú de la aplicación
 */
function onOpen() {

  SpreadsheetApp.getUi().createMenu('🧑‍🏫 Horarios a Calendar')
    .addItem('↗️ Generar clases en Calendar', 'm_CrearEventos')
    .addSeparator()
    .addItem('🗓️ Buscar calendarios instructores', 'm_ObtenerCalInstructores')
    .addItem('🏫 Buscar salas', 'm_ObtenerSalas')
    .addToUi();

}

/**
 * Muestra un toast informativo con algunos valores por defecto.
 * 
 * @param {string}  mensaje
 * @param {number}  [tiempoSeg]
 * @param {string}  [titulo]
 */
function mostrarMensaje(mensaje, tiempoSeg = -1, titulo = PARAM.nombreAp) {

  SpreadsheetApp.getActive().toast(mensaje, titulo, tiempoSeg);

}

/**
 * Muestra un cuadro de diálogo con mensaje y botones personalizados, usando
 * ciertos valores por defecto y devuelve el botón pulsado por el usuario
 * como una [enumeración `Ui.Button`](https://developers.google.com/apps-script/reference/base/button).
 * 
 * @param   {string}    mensaje
 * @param   {ButtonSet} [botones] [Enumeración `Ui.ButtonSet`](https://developers.google.com/apps-script/reference/base/button-set)
 * @param   {string}    [titulo]
 * 
 * @return  {Button}    Botón sobre el que se ha hecho clic    
 */
function alerta(mensaje, botones = SpreadsheetApp.getUi().ButtonSet.OK_CANCEL, titulo = PARAM.nombreApp) {

  return SpreadsheetApp.getUi().alert(titulo, mensaje, botones);

}

/**
 * Devuelve un array con el contenido de todas las celdas de una hoja que contienen datos
 * que se encuentran a partir de una fila y columnas iniciales dadas (se asumen 1 si se omiten).
 * 
 * @param   {SpreadsheetApp.Sheet}  hoja   
 * @param   {number}                [numFila]     Nº de fila, comenzando por 1 (se asume 1 si no se facilita)
 * @param   {number}                [numColumna]  Nº de columna, comenzando por 1 (se asume 1 si no se facilita)
 * 
 * @return  {Array}                 Valores de las celdas de la tabla
 */
function leerDatosHoja(hoja, numFila = 1, numColumna = 1) {

  return hoja.getDataRange().getValues().slice(numFila - 1).map(fila => fila.slice(numColumna - 1));

}

/**
 * Escribe los valores de una matriz en las celdas de una hoja a partir de una fila y columnas dadas
 * (se asumen 1 si se omiten), borrando previamente de manera opcional los valores y/o formato existentes.
 * Si no se establecen los parámetros opcionales `borrarDatos` y `borrarFormato` se borrarán únicamente
 * los valores previos. Si la matriz de datos a escribir es 'undefined' solo se borrarán las celdas, en su caso.
 * 
 * @param {SpreadsheetApp.Sheet}  hoja
 * @param {Array}                 matriz            `undefined` si no se desean escribir datos en la tabla
 * @param {number}                [numFila]         Nº de fila, comenzado por 1 (se asume 1 si no se facilita)
 * @param {number}                [numColumna]      Nº de columna, comenzado por 1 (se asume 1 si no se facilita)
 * @param {boolean}               [borrarDatos]     VERDADERO si se desean borrar los valores
 * @param {boolean}               [borrarFormato]   VERDADERO si se desean borrar formato y reglas de validación
 */
function actualizarDatosTabla(hoja, matriz, numFila = 1, numColumna = 1, borrarDatos = true, borrarFormato = false) {

  // Borrar datos previos de la tabla
  if ((borrarDatos || borrarFormato) && hoja.getLastRow() >= numFila) {
    hoja.getDataRange()
      .offset(numFila - 1, numColumna - 1, hoja.getLastRow() - numFila + 1, hoja.getLastColumn() - numColumna + 1)
      .clear({ contentsOnly: borrarDatos, formatOnly: borrarFormato });
  }

  // Escribir datos en la tabla, si se han facilitado
  if (matriz && matriz.map) hoja.getRange(numFila, numColumna, matriz.length, matriz[0].length).setValues(matriz);

}

/**
 * Elimina las filas y/o columnas sobrantes de la hoja indicada,
 * devuelve el número de filas y/o columnas eliminadas dentro de
 * un objeto como `{ filas: number, columnas: number }`.
 * 
 * @param {SpreadsheetApp.Sheet}        hoja                  Hoja a reducir
 * @param {Object}                      [reducir]             Elementos a eliminar como `{ filas: boolean, columnas: boolean }`
 * @param {Boolean}                     reducir.filas
 * @param {Boolean}                     reducir.columnas
 *
 * @return {Object} eliminadas          Nº de F/C eliminadas
 * @return {number} eliminadas.filas
 * @return {number} eliminadas.columnas
 */
function reducirHoja(hoja, reducir = { filas: true, columnas: false }) {

  const nFilas = hoja.getLastRow();
  const nMaxFilas = hoja.getMaxRows();
  const nColumnas = hoja.getLastColumn();
  const nMaxColumnas = hoja.getMaxColumns();

  if (reducir.filas && nMaxFilas > nFilas) hoja.deleteRows(nFilas + 1, nMaxFilas - nFilas);
  if (reducir.columnas && nMaxColumnas > nColumnas ) hoja.deleteColumns(nColumnas + 1, nMaxColumnas - nColumnas);

  return { filas: nMaxFilas - nFilas, columnas: nMaxColumnas - nColumnas };

}




function botonCheckEventos() {
  console.info(reducirHoja(SpreadsheetApp.getActive().getSheetByName('Instructores'), {}));
}

