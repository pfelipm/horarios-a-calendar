/**
 * Creación de eventos en Google Calendar
 * 
 * @OnlyCurrentDoc
 */


function m_CrearEventos() {

  // Hoja de eventos
  const hdc = SpreadsheetApp.getActive();
  const hojaEventos = hdc.getSheetByName(PARAM.eventos.hoja);

  // Leer datos necesarios para la geneneración de los eventos en Calendar.
  const instructores = leerDatosHoja(hdc.getSheetByName(PARAM.instructores.hoja), PARAM.instructores.filaEncabezado + 1);
  const salas = leerDatosHoja(hdc.getSheetByName(PARAM.salas.hoja), PARAM.salas.filaEncabezado + 1);
  const eventos = leerDatosHoja(hojaEventos, PARAM.eventos.filaEncabezado);

  let creados = 0;
  let saltados = 0;

  const resultado = eventos.map(evento => {

    const title = evento[0] + ' ' + evento[1] + ' (' + evento[2] + ')';
    const endDate = evento[5];
    // ⚠️ Marcianada al leer celdas con datos de hora sin fecha:
    // Ej, celda: 17:00:00 (formateada como hora) >> Objeto Date:Sat Dec 30 1899 17:24:05 GMT+0009 (Central European Standard Time) 😵‍💫
    // https://twitter.com/pfelipm/status/1572253070930878464
    // Posible solución:
    //   - Leer valor de fecha normalmente con getValue() y de hora con getDisplayValue()
    //   - Combinar haciendo eventStart = new Date(fecha.setHours(fechaComoTexto.split(':')[0], fechaComoTexto.split(':')[1]))
    // Paso de movidas y "monto" fecha + hora mediante fórmulas en la hoja de cálculo, eso me evita leer la tabla de dos modos distintos
    const startTime = evento[7]; 
    const endTime = evento[9];
    const dias = evento[3].split('-');
    const calendario = CalendarApp.getCalendarById(instructores.find(instructor => instructor[0] == evento[2])[2]);
    const guests = instructores.find(instructor => instructor[0] == evento[2])[1] + ',' + salas.find(sala => sala[0] == evento[10])[1];
    const idPrevio = evento[11];
    if (title && endDate && startTime && endTime && guests && dias && calendario && !idPrevio) {

      // No es necesario que el día de la semana de startTime coincida con el 1º en la serie según la recurrencia, Calendar ajusta internamente
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

  hojaEventos.getRange(T_RES).setValues(resultado);
  SpreadsheetApp.getUi().alert('Eventos procesados', 'Creados: ' + creados + '\nSaltados: ' + saltados, SpreadsheetApp.getUi().ButtonSet.OK);

}
