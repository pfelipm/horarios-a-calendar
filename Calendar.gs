/**
 * Crea eventos en calendario de acuerdo con datos de tabla
 * 
 * Cosas & mejoras:
 *   - Buscar primer día de la semana a partir del de inicio general del grupo
 *     (!) no es necesario, Calendar busca él solito el 1er día de la semana tras la 
 *         fecha de inicio de la repetición
 * 
 * @OnlyCurrentDoc
 */

function m_CrearEventos() {

  const hoja = SpreadsheetApp.getActive().getSheetByName('Eventos')

  const docentes = hoja.getRange(T_CAL_PROF).getValues();
  const aulas = hoja.getRange(T_CAL_AULA).getValues();
  const eventos = hoja.getRange(T_EVENTOS).getValues();
  let creados = 0;
  let saltados = 0;

  const resultado = eventos.map(evento => {

    const title = evento[0] + ' ' + evento[1] + ' (' + evento[2] + ')';
    const endDate = evento[5];
    const startTime = evento[7];
    const endTime = evento[9];
    const dias = evento[3].split('-');
    const calendario = CalendarApp.getCalendarById(docentes.find(docente => docente[0] == evento[2])[2]);
    const guests = docentes.find(docente => docente[0] == evento[2])[1] + ',' + aulas.find(aula => aula[0] == evento[10])[1];
    const idPrevio = evento[11];
    if (title && endDate && startTime && endTime && guests && dias && calendario && !idPrevio) {

      const recurrence = CalendarApp.newRecurrence()
        .addWeeklyRule()
        .onlyOnWeekdays(dias.map(dia => {
          switch (dia) {
            case 'L': return CalendarApp.Weekday.MONDAY; break;
            case 'M': return CalendarApp.Weekday.TUESDAY; break;
            case 'X': return CalendarApp.Weekday.WEDNESDAY; break;
            case 'J': return CalendarApp.Weekday.THURSDAY; break;
            case 'V': return CalendarApp.Weekday.FRIDAY; break;
            case 'S': return CalendarApp.Weekday.SATURDAY; break;
            case 'D': return CalendarApp.Weekday.SUNDAY; break;
          }
        })).until(endDate);

      creados++;
      return [calendario.createEventSeries(title, startTime, endTime, recurrence, {
        guests: guests,
        sendInvites: true
      }).getId(), new Date()];

    } else {
      // Si no se ha creado un nuevo evento se mantienen el ID / fecha previos en columnas resultados
      saltados++
      return [idPrevio, evento[12]];
    }
  });

  hoja.getRange(T_RES).setValues(resultado);
  SpreadsheetApp.getUi().alert('Eventos procesados', 'Creados: ' + creados + '\nSaltados: ' + saltados, SpreadsheetApp.getUi().ButtonSet.OK);

}






// Test marcianada al leer celdas con valores de solo hora
function foo() {
  const date = SpreadsheetApp.getActive.getSheetByName('Eventos').getRange('E3').getDisplayValue();
  const time = SpreadsheetApp.getActive.getSheetByName('Eventos').getRange('G3').getDisplayValue();
  console.info(date, time)
  const timeAsText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Eventos').getRange('G3').getDisplayValue();
  // Celda: 17:00:00 (formateada como hora) >> Objeto Date:Sat Dec 30 1899 17:24:05 GMT+0009 (Central European Standard Time) 😵‍💫
  // https://twitter.com/pfelipm/status/1572253070930878464
  console.info(timeAsText)
  const hours = timeAsText.split(':')[0];
  const minutes = timeAsText.split(':')[1];
  const eventStart = Date.parse(date + ' ' + time)
  console.info(eventStart)
}