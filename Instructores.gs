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

  // Nos movemos a la hoja de gestión de instructores antes de solicitar confirmación
  const hojaActual = SpreadsheetApp.getActiveSheet();
  const hojaInstructores = SpreadsheetApp.getActive().getSheetByName(PARAM.instructores.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se sobreescribirán los calendarios existentes.') == SpreadsheetApp.getUi().Button.OK) {

    const prefijo = hojaInstructores.getRange(PARAM.instructores.prefijo).getValue();
    if (prefijo == '') alerta(`¡No se ha introducido un prefijo en la celda ${PARAM.instructores.prefijo}!`, SpreadsheetApp.getUi().ButtonSet.OK);
    else {

      mostrarMensaje('Buscando calendarios de instructores...');

      const calendarios = CalendarApp.getAllCalendars().reduce((lista, calendario) => {
        if (calendario.getName().startsWith(prefijo)) return [...lista, [calendario.getName(), calendario.getId()]];
        else return lista;
      }, []);

      if (calendarios && calendarios.length > 0) {

        // Escribe datos en la tabla (hoja), solo se borra la columna con los ID obtenidos previamente,
        // las iniciales e emails se mantienen para facilitar su reutilización.
        // Mejora: Guardar lista de calendarios en otra hoja y usar un desplegable para asignar a cada instructor
        actualizarDatosTabla(
          hojaInstructores,
          calendarios.sort(([nombre1, id1], [nombre2, id2]) => nombre1.localeCompare(nombre2)),
          PARAM.instructores.filEncabezado + 1,
          PARAM.instructores.colNombreCalObtenido);

        // Eliminar filas sobrantes y mostrar mensajes de resultado
        hojaInstructores.getRange(PARAM.instructores.ultEjecucion).setValue(new Date());
        reducirHoja(hojaInstructores);
        mostrarMensaje(`Se han obtenido ${calendarios.length} calendarios.`, 5);

      } else mostrarMensaje('No se han encontrado calendarios.', 5);
    
    }

  } else hojaActual.activate();

}