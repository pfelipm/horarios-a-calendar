/****************************************************
 * Funciones para gestionar la hoja de instructores *
 ****************************************************/

/**
 * Función invocada desde el menú del script.
 * Escribe en la tabla de calendario de instructores los ID (emails)
 * de todos aquellos que comienzan por el prefijo indicado en la 
 * celda PARAM.instructores.prefijo.
 */
function m_ObtenerCalInstructores() {

  if (alerta('Se sobreescribirán los calendarios existentes') == SpreadsheetApp.getUi().Button.OK) {

    mostrarMensaje('Buscando calendarios de instructores...');
    
    hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARAM.instructores.hoja).activate();
    const prefijo = hoja.getRange(PARAM.instructores.prefijo).getValue();

    const calendarios = CalendarApp.getAllCalendars().reduce((lista, calendario) => {
      if (calendario.getName().startsWith(prefijo)) return [...lista, [calendario.getName(), calendario.getId()]];
      else return lista;
    },[]);

    if (calendarios && calendarios.length > 0) {

      // Escribe datos en la tabla (hoja), solo se borra la columna con los ID obtenidos previamente,
      // las iniciales e emails se mantienen
      actualizarDatosTabla(
        hoja,
        calendarios.sort(([nombre1, id1], [nombre2, id2]) => nombre1.localeCompare(nombre2)),
        PARAM.instructores.filEncabezado + 1,
        PARAM.instructores.colNombreCal);
      
      // Eliminar filas sobrantes y mostrar mensajes de resultado
      hoja.getRange(PARAM.instructores.ultEjecucion).setValue(new Date());
      reducirHoja(hoja);
      mostrarMensaje(`Se han obtenido ${calendarios.length} calendarios.`,5);
        
    } else mostrarMensaje('No se han encontrado calendarios.',5);
  
  }

}