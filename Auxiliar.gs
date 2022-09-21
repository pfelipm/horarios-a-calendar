/**
 * Funciones auxiliares
 */

/**
 * Crea el menú de la aplicación
 */
function onOpen() {

  SpreadsheetApp.getUi().createMenu('🧑‍🏫 Horarios a Calendar')
    .addItem('🗓️ Volcar clases en Calendar', 'm_CrearEventos')
    .addSeparator()
    .addItem('🏫 Actualizar lista de salas', 'm_ObtenerSalas')
    .addToUi();
  
}

/**
 * Muestra un toast informativo
 * @param {String}  mensaje
 * @param {Number}  tiempoSeg
 * @param {String}  titulo
 */
function mostrarMensaje(mensaje, tiempoSeg = -1, titulo) {
  
  const hdc = SpreadsheetApp.getActive();
  if (!mensaje) mensaje = hdc.getName();
  if (!titulo) titulo = PARAM.nombreApp;
  hdc.toast(mensaje, titulo, tiempoSeg);

}

