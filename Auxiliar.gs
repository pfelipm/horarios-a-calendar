/************************
 * Funciones auxiliares *
 ************************/


/**
 * Crea el menú de la aplicación
 */
function onOpen() {

  SpreadsheetApp.getUi().createMenu('🗓️ Horarios a Calendar')
    .addItem('➕ Generar clases en Calendar', 'm_CrearEventos')
    .addItem('🟰 Actualizar clases en Calendar', 'm_ActualizarEventos')
    .addItem('✖️ Eliminar clases en Calendar', 'm_EliminarEventos')
    .addSeparator()
    .addItem('🧑‍🏫 Buscar calendarios instructores', 'm_ObtenerCalInstructores')
    .addItem('🏫 Buscar salas', 'm_ObtenerSalas')
    .addSeparator()
    .addItem(`💡 Acerca de ${PARAM.nombreApp}`, 'acercaDe')
    .addToUi();

}

/**
 * Muestra la ventana de información de la aplicación
 */
function acercaDe() {

  let panel = HtmlService.createTemplateFromFile('acercaDe');
  panel.nombre = PARAM.nombre;
  panel.version = PARAM.version;
  panel.urlRepoGitHub = PARAM.urlRepoGitHub;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(425), `${PARAM.icono} ${PARAM.nombre}`);

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
 * @param   {ButtonSet} [botones] [Enumeración `Ui.ButtonSet`](https://developers.google.com/apps-script/reference/base/button-set).
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
 * @param   {number}                [numFila]     Nº de fila, comenzando por 1 (se asume 1 si no se facilita).
 * @param   {number}                [numColumna]  Nº de columna, comenzando por 1 (se asume 1 si no se facilita).
 * 
 * @return  {Array}                 Valores de las celdas de la tabla
 */
function leerDatosHoja(hoja, numFila = 1, numColumna = 1) {

  return hoja.getDataRange().getValues().slice(numFila - 1).map(fila => fila.slice(numColumna - 1));

}

/**
 * Escribe los valores de una matriz en las celdas de una hoja a partir de una fila y columnas dadas
 * (se asumen 1 si se omiten), borrando previamente de manera opcional los valores y/o formato existentes
 * en la hoja. Si no se establecen los parámetros opcionales `borrarDatos` y `borrarFormato` se borrarán
 * únicamente los valores previos. Si la matriz de datos a escribir es 'undefined' solo se efectuará el
 * borrado de datos.
 * 
 * @param {SpreadsheetApp.Sheet}  hoja
 * @param {Array}                 matriz            Pasar un valor `undefined` si solo se desea borrar la hoja.
 * @param {number}                [numFila]         Nº de fila, comenzado por 1 (se asume 1 si no se facilita).
 * @param {number}                [numColumna]      Nº de columna, comenzado por 1 (se asume 1 si no se facilita).
 * @param {boolean}               [borrarDatos]     VERDADERO si se desean borrar los valores.
 * @param {boolean}               [borrarFormato]   VERDADERO si se desean borrar formato y reglas de validación.
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
 * @param   {SpreadsheetApp.Sheet}        hoja                  Hoja a reducir.
 * @param   {Object}                      [reducir]             Elementos a eliminar como `{ filas: boolean, columnas: boolean }`.
 * @param   {Boolean}                     reducir.filas
 * @param   {Boolean}                     reducir.columnas
 *
 * @return  {Object} eliminadas          Nº de F/C eliminadas
 * @return  {number} eliminadas.filas
 * @return  {number} eliminadas.columnas
 */
function reducirHoja(hoja, reducir = { filas: true, columnas: false }) {

  const numFilas = hoja.getLastRow();
  const numMaxFilas = hoja.getMaxRows();
  const numColumnas = hoja.getLastColumn();
  const numMaxColumnas = hoja.getMaxColumns();

  if (reducir.filas && numMaxFilas > numFilas) hoja.deleteRows(numFilas + 1, numMaxFilas - numFilas);
  if (reducir.columnas && numMaxColumnas > numColumnas) hoja.deleteColumns(numColumnas + 1, numMaxColumnas - numColumnas);

  return { filas: numMaxFilas - numFilas, columnas: numMaxColumnas - numColumnas };

}

/**
 * Envoltorio para la función `conmutarChecks()`, a la que invoca con los
 * parámetros específicos correspondientes a un botón o comando de menú
 * determinado para conmutar el estado de un intervalo de casillas de
 * verificación.
 */
function botonCheckEventos() {

  conmutarChecks(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARAM.eventos.hoja),
    PARAM.eventos.filEncabezado + 1,
    PARAM.eventos.colCheck,
    2);

}

/**
 * Conmuta el estado un conjunto de casillas de verificación a partir de la fila indicada.
 * Devuelve el nº de casillas han sido actualizadas.
 * 
 * @param   {SpreadsheetApp.Sheet}  hoja            Hoja en la que se encuentra el intervalo con casillas de verificación.
 * @param   {number}                filCheck        Nº de la fila en la que se encuentra la primera casilla de verificación.
 * @param   {number}                colCheck        Nº de la columna donde se encuentran las casillas de verificación.
 * @param   {number}                colDatos        Nº de la columna que se usa para determinar si hay datos en cada fila.
 * @param   {numFilas}              numFilas        Nº de casillas de verificación o '0' si se extienden hasta `lastRow()`.
 * @param   {string}                propiedadEstado Clave de las `ScriptProperties` en la que se guardará el estado actual de las casillas.
 * 
 * @return  {number}                                Número de casillas de verificación actualizadas.
 */
function conmutarChecks(hoja, filCheck, colCheck, colDatos = 1, numFilas = 0, propiedadEstado = 'estadoCheck01') {

  numFilas = !numFilas ? hoja.getLastRow() - filCheck + 1 : numFilas;

  // Uso el almacén del script porque puede administrarse desde el editor,
  // pero lo adecuado es emplear el del documento (⚠️ imprescindible en un complemento).
  const propiedadesDoc = PropertiesService.getScriptProperties();
  const estado = JSON.parse(propiedadesDoc.getProperty(propiedadEstado));
  let numCheckActivos;

  if (colDatos > 1) {

    const rangoExisten = hoja.getRange(filCheck, colCheck + 1, numFilas);
    const existen = rangoExisten.getValues();
    numCheckActivos = existen.length - existen.reverse().findIndex(el => el[0] != '');

  } else numCheckActivos = numFilas;

  if (numCheckActivos > 0) {

    if (estado) {
      hoja.getRange(filCheck, colCheck, numCheckActivos).setValue(false);
      propiedadesDoc.setProperty(propiedadEstado, false);
    } else {
      hoja.getRange(filCheck, colCheck, numCheckActivos).setValue(true);
      propiedadesDoc.setProperty(propiedadEstado, true)
    }

  }

  return numCheckActivos;

}
