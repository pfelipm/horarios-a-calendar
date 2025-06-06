/********************************************************
 * Funciones para gestionar la hoja de salas (recursos) *
 ********************************************************/


/**
 * Función invocada desde el menú del script.
 * escribe en la hoja de cálculo la lista de salas del dominio
 */
function m_ObtenerSalas() {

  // Nos movemos a la hoja de gestión de salas antes de solicitar confirmación
  const hojaActual = SpreadsheetApp.getActiveSheet();
  const hojaSalas = SpreadsheetApp.getActive().getSheetByName(PARAM.salas.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se sobreescribirán las salas existentes.') == SpreadsheetApp.getUi().Button.OK) {

    mostrarMensaje('Buscando salas en el dominio...');

    const salas = obtenerSalas().map(recurso =>
      [
        recurso.resourceName,
        recurso.generatedResourceName,
        recurso.resourceType,
        recurso.userVisibleDescription,
        recurso.resourceEmail
      ]
    );

    if (salas && salas.length > 0) {

      // Escribe datos en la tabla (hoja)
      actualizarDatosTabla(hojaSalas, salas, PARAM.salas.filEncabezado + 1, PARAM.salas.colDatos);

      // Eliminar filas sobrantes y mostrar mensajes de resultado
      reducirHoja(hojaSalas);
      hojaSalas.getRange(PARAM.salas.ultEjecucion).setValue(new Date());
      mostrarMensaje(`Se han obtenido ${salas.length} salas.`, 5);

    } else mostrarMensaje('No se han encontrado salas.', 5);

  } else hojaActual.activate();

}

/**
 * Devuelve una lista de los recursos del dominio de tipo `CONFERENCE_ROOM`.
 * Puede ser invocada por un usuario sin previlegios administrativos, dadoque
 * la información de recursos de calendario es pública en el el dominio.
 * 
 * @return {Array<Admin_directory_v1.Admin.Directory_v1.Schema.CalendarResource>} Lista de salas
 */
function obtenerSalas() {

  let recursos = [];
  let pageToken;

  // Dentro de bloque try por si se ejecuta en una cuenta personal
  try {

    do {

      const respuesta = AdminDirectory.Resources.Calendars.list('my_customer',
        {
          maxResults: 100,
          orderBy: 'resourceName asc',
          pageToken: pageToken,
          query: 'resourceCategory=CONFERENCE_ROOM'
        });
      if (respuesta) {
        recursos = recursos.concat(respuesta.items);
        pageToken = respuesta.nextPageToken;
      }

    } while (pageToken);

  } catch (e) {
    alerta('⚠️ Error al obtener la lista de salas del dominio.', SpreadsheetApp.getUi().ButtonSet.OK);
  }

  return recursos;

}