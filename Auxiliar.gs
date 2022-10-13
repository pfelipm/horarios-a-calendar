/************************
 * Funciones auxiliares *
 ************************/


/**
 * Crea el menú de la aplicación
 */
function onOpen() {

  SpreadsheetApp.getUi().createMenu('🗓️ Horarios a Calendar')
    .addItem('➕ Generar/actualizar clases en Calendar', 'm_CrearEventos')
    .addItem('✖️ Eliminar clases en Calendar', 'm_EliminarEventos')
    .addSeparator()
    .addItem('🧑‍🏫 Buscar calendarios instructores', 'm_ObtenerCalInstructores')
    .addItem('🏫 Buscar salas', 'm_ObtenerSalas')
    .addSeparator()
    .addItem(`💡 Acerca de ${PARAM.nombre}`, 'acercaDe')
    .addToUi();

}

/**
 * Muestra la ventana de información de la aplicación
 */
function acercaDe() {

  let panel = HtmlService.createTemplateFromFile('Acerca de');
  panel.nombre = PARAM.nombre;
  panel.version = PARAM.version;
  panel.urlRepoGitHub = PARAM.urlRepoGitHub;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(425), `${PARAM.icono} ${PARAM.nombre}`);

}

/**
 * Muestra un toast informativo con algunos valores por defecto.
 * 
 * @param {string}  mensaje
 * @param {number}  [tiempoSeg] Segundos en pantalla [hasta clic].
 * @param {string}  [titulo]    Título del toast [`PARAM.icono` `PARAM.nombre`].
 */
function mostrarMensaje(mensaje, tiempoSeg = -1, titulo = `${PARAM.icono} ${PARAM.nombre}`) {

  SpreadsheetApp.getActive().toast(mensaje, titulo, tiempoSeg);

}

/**
 * Muestra un cuadro de diálogo con mensaje y botones personalizados, usando
 * ciertos valores por defecto y devuelve el botón pulsado por el usuario
 * como una [enumeración `Ui.Button`](https://developers.google.com/apps-script/reference/base/button).
 * 
 * @param   {string}    mensaje
 * @param   {ButtonSet} [botones] [Enumeración `Ui.ButtonSet`](https://developers.google.com/apps-script/reference/base/button-set) [ButtonSet.OK_CANCEL].
 * @param   {string}    [titulo]  Título de la alerta [`PARAM.icono` `PARAM.nombre`].
 * 
 * @return  {Button}    Botón sobre el que se ha hecho clic    
 */
function alerta(mensaje, botones = SpreadsheetApp.getUi().ButtonSet.OK_CANCEL, titulo = `${PARAM.icono} ${PARAM.nombre}`) {

  return SpreadsheetApp.getUi().alert(titulo, mensaje, botones);

}

/**
 * Devuelve un array con el contenido de todas las celdas de una hoja que contienen datos
 * que se encuentran a partir de una fila y columnas iniciales dadas (se asumen 1 si se omiten).
 * 
 * @param   {SpreadsheetApp.Sheet}  hoja   
 * @param   {number}                [numFila]     Nº de fila, comenzando por 1 [1].
 * @param   {number}                [numColumna]  Nº de columna, comenzando por 1 [1].
 * 
 * @return  {Array} Valores de las celdas de la tabla
 */
function leerDatosHoja(hoja, numFila = 1, numColumna = 1) {

  return hoja.getDataRange().getValues().slice(numFila - 1).map(fila => fila.slice(numColumna - 1));

}

/**
 * Escribe los valores de una matriz en las celdas de una hoja a partir de una fila y columnas dadas
 * (se asumen 1 si se omiten), borrando previamente de manera opcional los valores y/o formato existentes
 * en la hoja. Si no se establecen los parámetros opcionales `borrarDatos` y `borrarFormato` se borrarán
 * únicamente los valores previos. Si la matriz de datos a escribir es `null`o 'undefined' solo se efectuará
 * el borrado de datos.
 * 
 * @param {SpreadsheetApp.Sheet}  hoja
 * @param {Array}                 matriz            Pasar un valor `null' o `undefined` si solo se desea borrar la hoja.
 * @param {number}                [numFila]         Nº de fila, comenzado por 1 [1].
 * @param {number}                [numColumna]      Nº de columna, comenzado por 1 [1].
 * @param {boolean}               [borrarDatos]     VERDADERO si se desean borrar los valores [`true`].
 * @param {boolean}               [borrarFormato]   VERDADERO si se desean borrar formato y reglas de validación [`false`].
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
 * @param   {Object}                      [reducir]             Elementos a eliminar [`{ filas: true, columnas: false }`].
 * @param   {Boolean}                     reducir.filas
 * @param   {Boolean}                     reducir.columnas
 *
 * @return  {Object} eliminadas           Nº de F/C eliminadas
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
    SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja),
    PARAM.eventos.filEncabezado + 1,
    PARAM.eventos.colCheck,
    2); // 2 == Columna "Grupo"

}

/**
 * Conmuta el estado un conjunto de casillas de verificación a partir de la fila indicada de una tabla (hoja).
 * Devuelve el nº de casillas han sido actualizadas.
 * 
 * @param   {SpreadsheetApp.Sheet}  hoja              Hoja en la que se encuentra el intervalo con casillas de verificación.
 * @param   {number}                filCheck          Nº de la fila en la que se encuentra la primera casilla de verificación.
 * @param   {number}                colCheck          Nº de la columna donde se encuentran las casillas de verificación.
 * @param   {number}                [colDatos]        Nº de la columna que se usa para determinar si hay datos en cada fila [1].
 * @param   {numFilas}              [numFilas]        Nº de casillas de verificación o '0' si se extienden hasta `lastRow()` [0].
 * @param   {string}                [propiedadEstado] Clave de las `ScriptProperties` que guardará el estado actual de las casillas [`'estadoCheck01'`].
 * 
 * @return  {number} Número de casillas de verificación actualizadas.
 */
function conmutarChecks(hoja, filCheck, colCheck, colDatos = 1, numFilas = 0, propiedadEstado = 'estadoCheck01') {

  numFilas = !numFilas ? hoja.getLastRow() - filCheck + 1 : numFilas;

  // Uso el almacén del script porque puede administrarse desde el editor,
  // pero lo adecuado es emplear el del documento (⚠️ imprescindible en un complemento).
  const propiedadesDoc = PropertiesService.getScriptProperties();
  const estado = JSON.parse(propiedadesDoc.getProperty(propiedadEstado));
  let numCheckActivos;

  if (colDatos > 1) {

    const existen = hoja.getRange(filCheck, colDatos, numFilas).getValues();
    const indiceUltimoValor = existen.reverse().findIndex(el => el[0] != '');
    if (indiceUltimoValor == -1) numCheckActivos = 0;
    else numCheckActivos = existen.length - indiceUltimoValor;

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

/**
 * Busca cada ocurrencia del evento caracterizado por grupo y clase
 * en la tabla de la hoja de registro de eventos `PARAM.registro.hoja` y
 * si existen:
 * 
 *  1. Elimina el evento de calendario asociado
 *  2. Elimina la fila de la tabla en la que se ha hallado.
 * 
 * (!) Si la tabla no ha sido manipulada y no se han producido errores, solo debería
 * darse una única coincidencia, no obstante se comprueban múltiples.
 * 
 * @param   {string}  grupo Código de grupo de la clase.
 * @param   {string}  clase Código de la clase.
 * 
 * @return  {number}  Número de filas / eventos eliminados
 */
function eliminarEventosPreviosRegistro(grupo, clase) {

  const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(PARAM.registro.hoja);

  const eventosEliminar = leerDatosHoja(hojaRegistro, PARAM.registro.filEncabezado + 1)
    .reduce((listaEventos, evento, indice) => {

      if (evento[PARAM.registro.colGrupo - 1] == grupo && evento[PARAM.registro.colClase - 1] == clase) {
        return [...listaEventos,
          {
            fila: PARAM.registro.filEncabezado + indice + 1,
            idEvento: evento[PARAM.registro.colIdEv - 1],
            idCalendario: evento[PARAM.registro.colIdCal - 1]
          }];
      } else return listaEventos;

    },[]);
    
  //console.info(eventosEliminar);

  let eventosEliminados = 0;
  if (eventosEliminar.length > 0) {

    // Se comienza desde el final de la hoja para que los números de filas sigan siendo válidos tras cada eliminación
    eventosEliminar.reverse().forEach(evento => {

      try {

        // Refinar esto: si existe fila siempre debe borrarse, aunque el evento ya no exista en Calendar
        // diferenciar ambos posibles resultados para mostrar un mensaje de resultado significativo
        // CalendarApp.getCalendarById(evento.idCalendario).getEventSeriesById(evento.idEvento).deleteEventSeries();
        hojaRegistro.deleteRow(evento.fila);
        eventosEliminados++;

      } catch(e) {

        // Sin tratamiento, simplemente se registra el error
        console.info('Excepción: ' + evento + '\n' + e.message);

      }
    
    });
  }

  return eventosEliminados;

}