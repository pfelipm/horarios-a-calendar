/******************************************************
 * Funciones para gestionar la hoja de salas (recursos)
 ******************************************************/

/**
 * Función de menú, escribe en la hoja de cálculo la lista
 * de salas del dominio
 */
function m_ObtenerSalas() {
  
  mostrarMensaje('Buscando salas en el dominio...');

  hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARAM.salas.hoja);
  const salas = obtenerSalas().map(recurso =>
    [
      recurso.resourceName,
      recurso.generatedResourceName,
      recurso.resourceType,
      recurso.userVisibleDescription,
      recurso.resourceEmail
    ]
  );

  if (salas.length > 0) {

    hoja.getDataRange().offset(PARAM.salas.filaEncabezado,0).clearContent();
    hoja.getRange(PARAM.salas.filaEncabezado + 1,PARAM.salas.colDatos, salas.length, salas[0].length)
      .setValues(salas);

  }

  mostrarMensaje(`Se han obtenido ${salas.length} salas.`,5);

}

/**
 * Devuelve una lista de los recursos del dominio de tipo 'CONFERENCE_ROOM'.
 * @return  {Array<CalendarResource>}
 */
function obtenerSalas() {

  let recursos = [];
  let pageToken;
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

  return recursos;

}