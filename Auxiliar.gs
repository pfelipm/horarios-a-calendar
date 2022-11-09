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
    .addItem('♻️ Borrar resultados proceso', 'm_EliminarResultados')
    .addSeparator()
    .addItem('👥 Buscar calendarios instructores', 'm_ObtenerCalInstructores')
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
 * @param   {boolean}               [estado]          Establece opcionalmente el estado de las casillas al valor indicado, se ignora si `null`o `u
 * @param   {string}                [propiedadEstado] Clave de las `ScriptProperties` que guardará el estado actual de las casillas [`'estadoCheck01'`].
 * 
 * @return  {number}  Número de casillas de verificación actualizadas.
 */
function conmutarChecks(hoja, filCheck, colCheck, colDatos = 1, numFilas = 0, estado, propiedadEstado = PARAM.propiedadEstadoCheck) {

  numFilas = !numFilas ? hoja.getLastRow() - filCheck + 1 : numFilas;

  // Uso el almacén del script porque puede administrarse desde el editor,
  // pero lo adecuado es emplear el del documento (⚠️ imprescindible en un complemento).
  const propiedadesDoc = PropertiesService.getScriptProperties();
  let numCheckActivos;
  if (estado == undefined || estado == null) estado = JSON.parse(propiedadesDoc.getProperty(propiedadEstado));
  else estado = !estado;

  if (colDatos > 1) {

    const existen = hoja.getRange(filCheck, colDatos, numFilas).getValues();
    const indiceUltimoValor = existen.reverse().findIndex(el => el[0] != '');
    if (indiceUltimoValor == -1) numCheckActivos = 0;
    else numCheckActivos = existen.length - indiceUltimoValor;

  } else numCheckActivos = numFilas;

  if (numCheckActivos > 0) {

    hoja.getRange(filCheck, colCheck, numCheckActivos).setValue(!estado);
    propiedadesDoc.setProperty(propiedadEstado, !estado);

  }

  return numCheckActivos;

}

/**
 * Busca cada ocurrencia del evento caracterizado por grupo y clase
 * en la tabla de la hoja de registro de eventos `PARAM.registro.hoja` y
 * si existe y su fecha de creación es anterior a `selloTiempoProceso`: 
 * 
 *  1. Elimina el evento de calendario asociado.
 *  2. Elimina la fila de la hoja de registros en la que se encontraba.
 * 
 * (!) Cada par (Grupo, Clase) puede tener múltiples coincidencias en la tabla
 * de registro de eventos, correspondientes a distintas sesiones semanales.
 * 
 * ¡Cuidado! Esto podría eliminar de Calendar las clases del curso anterior,
 * si es que aún se mantuvieran en la tabla de registro de eventos (poco probable),
 * a menos que se diferenciaran de algún modo, por ejemplo sufijo curso en grupo (Ej: DAM2 22/23).
 * 
 * @param   {string}  grupo               Código de grupo de la clase.
 * @param   {string}  clase               Código de la clase.
 * @param   {Object}  selloTiempoProceso  Objeto de la clase JS Date.
 * 
 * @return  {number}  Número de filas / eventos eliminados
 */
function eliminarEventosPreviosRegistro(grupo, clase, selloTiempoProceso) {

  const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(PARAM.registro.hoja);
  const eventosRegistroEliminar = leerDatosHoja(hojaRegistro, PARAM.registro.filEncabezado + 1)
    .reduce((listaEventos, evento, indice) => {

      if (
        evento[PARAM.registro.colGrupo - 1] == grupo &&
        evento[PARAM.registro.colClase - 1] == clase &&
        evento[PARAM.registro.colFechaProceso - 1] < selloTiempoProceso
      ) {
        return [...listaEventos,
        {
          fila: PARAM.registro.filEncabezado + indice + 1,
          idEvento: evento[PARAM.registro.colIdEv - 1],
          idCalendario: evento[PARAM.registro.colIdCal - 1]
        }];
      } else return listaEventos;

    }, []);

  //console.info(eventosRegistroEliminar);

  let eventosEliminados = 0;
  if (eventosRegistroEliminar && eventosRegistroEliminar.length > 0) {

    // Se comienza desde el final de la hoja para que los números de fila sigan siendo válidos tras cada eliminación
    eventosRegistroEliminar.reverse().forEach(evento => {

      try {

        // El evento asociado a la clase puede existir en la tabla de registro de eventos pero no en
        // Calendar. Perseguimos sincronicidad, por tanto siempre se borrará en la tabla, aunque solo
        // se incrementará la cuenta de eliminados si realmente se ha eliminado un evento del calendario.
        // ➕ Mejora: diferenciar ambos casos a la hora de devolver información de resultado.
        hojaRegistro.deleteRow(evento.fila);
        CalendarApp.getCalendarById(evento.idCalendario).getEventSeriesById(evento.idEvento).deleteEventSeries();
        eventosEliminados++;

      } catch (e) {
        // Sin tratamiento, simplemente se registra el error
        console.info('Excepción: ' + evento + '\n' + e.message);
      }

    });
  }

  return eventosEliminados;

}

function foo() {

  const hdc = SpreadsheetApp.getActive();
  const hojaEventos = hdc.getSheetByName(PARAM.eventos.hoja);
  const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
    .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
    .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true
      && eventoFila.ajustes[PARAM.eventos.colGrupo - 1] != ''
      && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

  // console.info(eventosFilas);
  console.info(eliminarEventosPreviosRegistroMultiple(eventosFilas));

}

/**
 * Busca cada ocurrencia de los eventos caracterizado por grupo y clase
 * que figuran en la tabla de clases a generar con los que se encuentran
 * en la hoja de registro de eventos `PARAM.registro.hoja` y si existe: 
 * 
 *  1. Elimina el evento de calendario asociado.
 *  2. Elimina la fila de la hoja de registros en la que se encontraba.
 * 
 * (!) Cada par (Grupo, Clase) puede tener múltiples coincidencias en la tabla
 * de registro de eventos, correspondientes a distintas sesiones semanales.
 *
 * Se diferencia de `eliminarEventosPreviosRegistro`en que en este caso se
 * realiza el borrado de los eventos asociados a todas las clases de una vez,
 * en lugar de una a una a medida que se procesa cada clase en la tabla de
 * generación.
 *  
 * ¡Cuidado! Esto podría eliminar de Calendar las clases del curso anterior,
 * si es que aún se mantuvieran en la tabla de registro de eventos (poco probable),
 * a menos que se diferenciaran de algún modo, por ejemplo sufijo curso en grupo (Ej: DAM2 22/23).
 *  
 * @param   {Array<Object>}  eventosFilas  Tabla de eventos a procesar, [{ ajustes }, { fila:[] }]
 * 
 * @return  {number}  Número de filas / eventos eliminados
 */
function eliminarEventosPreviosRegistroMultiple(eventosFilas) {

  const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(PARAM.registro.hoja);

  // Leer eventos registrados e indexar para no perder referencia de fila, dado que se eliminarán
  // de la matriz las marcadas para eliminar en Calendar
  const eventosRegistro = leerDatosHoja(hojaRegistro, PARAM.registro.filEncabezado + 1)
    .map((eventoRegistro, indice) => ({ ajustes: eventoRegistro, fila: indice }));

  const eventosRegistroEliminar = [];
  const clasesYaProcesadas = new Set();

  // Recorremos la lista de clases (Grupo, Clase) cuyos eventos previos deben eliminarse
  eventosFilas.forEach(eventoFila => {

    // Simplemente concatenar grupo + clase o grupo + separador + clase NO es seguro como clave única
    // Revisar código de acoplar() para una explicación más detallada en un caso similar
    const claveUnica = JSON.stringify([eventoFila.ajustes[PARAM.eventos.colGrupo - 1], eventoFila.ajustes[PARAM.eventos.colClase - 1]]);

    // ¿Se han buscado ya eventos registrados de la clase caracterizada por (Grupo, Clase)?
    if (!clasesYaProcesadas.has(claveUnica)) {

      // Recorremos la tabla de eventos registrados buscando todos los del mismo grupo y clase que el que se desea eliminar
      const filasEliminar = eventosRegistro.reduce((filasEliminar, evento) => {

        if (
          evento.ajustes[PARAM.registro.colGrupo - 1] == eventoFila.ajustes[PARAM.eventos.colGrupo - 1] &&
          evento.ajustes[PARAM.registro.colClase - 1] == eventoFila.ajustes[PARAM.eventos.colClase - 1]
        ) {
          // Incluir evento en la lista de eliminación de la hoja de registro de eventos y de Calendar
          eventosRegistroEliminar.push(
            {
              fila: PARAM.registro.filEncabezado + evento.fila + 1,
              idEvento: evento.ajustes[PARAM.registro.colIdEv - 1],
              idCalendario: evento.ajustes[PARAM.registro.colIdCal - 1]
            }
          );
          return [...filasEliminar, evento.fila];
        }
        return filasEliminar;

      }, []);

      // Eliminamos las filas con eventos ya marcados para eliminar de la matriz
      filasEliminar.forEach(indiceFila => eventosRegistro.splice(indiceFila, 1));

      // Clase ya procesada, no volveremos a intentar encontrar eventos relacionados
      clasesYaProcesadas.add(claveUnica);

    }

  });

  console.info(eventosRegistroEliminar);

  let eventosEliminados = 0;
  if (eventosRegistroEliminar.length > 0) {

    console.info(eventosRegistroEliminar.sort((evento1, evento2) => evento2.fila - evento1.fila));

    // Borrar eventos en orden decreciente de fila para que los números de fila sigan siendo válidos tras cada eliminación
    eventosRegistroEliminar.sort((evento1, evento2) => evento2.fila - evento1.fila).forEach(evento => {

      try {

        // El evento asociado a la clase puede existir en la tabla de registro de eventos pero no en
        // Calendar. Perseguimos sincronicidad, por tanto siempre se borrará en la tabla, aunque solo
        // se incrementará la cuenta de eliminados si realmente se ha eliminado un evento del calendario.
        // ➕ Mejora: diferenciar ambos casos a la hora de devolver información de resultado.
        hojaRegistro.deleteRow(evento.fila);
        CalendarApp.getCalendarById(evento.idCalendario).getEventSeriesById(evento.idEvento).deleteEventSeries();
        eventosEliminados++;

      } catch (e) {
        // Sin tratamiento, simplemente se registra el error
        console.info('Excepción: ' + evento + '\n' + e.message);
      }

    });
  }

  return eventosEliminados;

}