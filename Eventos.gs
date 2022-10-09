/*****************************************************
 * Funciones para gestionar la generación de eventos *
 *****************************************************/

/**
 * Función invocada desde el menú del script.
 * Inserta eventos recurrentes en los calendarios públicos de los instructores.
 * Opcionalmente: a) Reserva salas b) Envía invitaciones a los instructores asignadas a cada sesión.
 * Guarda la información de cada evento creado en la tabla `📦 Registro eventos`. Antes de crear uno
 * nuevo se intenta verificar si ya existe otro ya creado para esa clase. Para determinarlo se emplea
 * el par (Grupo, Clase), si existen varias sesiones semanales de una misma clase se eliminarán todas
 * antes de crear nuevos eventos para ellas.
 * 
 * **Nota:** Esta estrategia me parece más eficiente que etiquetar con el método
 * [Event.setTag()](https://developers.google.com/apps-script/reference/calendar/calendar-event#settagkey,-value)
 * los eventos generados y tener que hacer una búsqueda por fuerza bruta sobre ellos, lo que podría ser
 * terrible dado que podrían producirse cambios de instructor en una asignatura determinada.
 */
function m_CrearEventos() {

  // Hoja de eventos
  const hdc = SpreadsheetApp.getActive();
  const hojaEventos = hdc.getSheetByName(PARAM.eventos.hoja);

  // Leer ajustes de generación
  const checkInvitarInstructores = hojaEventos.getRange(PARAM.eventos.checkInvitarInstructores).getValue();
  const checkReservarEspacios = hojaEventos.getRange(PARAM.eventos.checkReservarEspacios).getValue();
  
  // Leer datos necesarios para la geneneración de los eventos en Calendar.
  const instructores = leerDatosHoja(hdc.getSheetByName(PARAM.instructores.hoja), PARAM.instructores.filEncabezado + 1);
  const salas = leerDatosHoja(hdc.getSheetByName(PARAM.salas.hoja), PARAM.salas.filEncabezado + 1);
  // Como hay casillas de verificación en la columna 1 es necesario descartar las filas en las que no hay eventos (sin Clase)
  const eventos = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
    .filter(evento => evento[PARAM.eventos.colCheck - 1] == true && evento[PARAM.eventos.colClase -1] != '');

  let creados = 0;
  let saltados = 0;

  // Vamos a crear eventos...
  const resultado = eventos.map(evento => {

    // Título del evento: Grupo + Clase + (iniciales instructor)
    const title = `${evento[PARAM.eventos.colGrupo - 1]} ${evento[PARAM.eventos.colClase]} (${evento[PARAM.eventos.colInstructor]})`;
    const endDate = evento[PARAM.eventos.colDiaFinRep - 1];

    // ⚠️ Marcianada al leer celdas con datos de hora sin fecha:
    // Ej, celda: 17:00:00 (formateada como hora) >> Objeto Date:Sat Dec 30 1899 17:24:05 GMT+0009 (Central European Standard Time) 😵‍💫
    // https://twitter.com/pfelipm/status/1572253070930878464
    // Posible solución:
    //   - Leer valor de fecha normalmente con getValue() y de hora con getDisplayValue()
    //   - Combinar haciendo eventStart = new Date(fecha.setHours(fechaComoTexto.split(':')[0], fechaComoTexto.split(':')[1]))
    // Paso de movidas y "monto" fecha + hora mediante fórmulas en la hoja de cálculo, eso me evita leer la tabla de dos modos distintos
   
    const startTime = evento[PARAM.eventos.colStartTime - 1]; 
    const endTime = evento[PARAM.eventos.colEndTime - 1];
    const dias = evento[PARAM.eventos.colDias - 1].split('-');
    
    // ¿Localizamos en la tabla de instructores las iniciales del asignado a esta clase?
    const instructor = instructores.find(instructor => instructor[PARAM.instructores.colIniciales - 1] == evento[PARAM.eventos.colInstructor - 1]);
    if (!instructor) {
      saltados++;
      return [new Date(), '⭕ Instructor no existe'];
    }

    // Si es que sí, obtenemos su calendario público y su calendario privado (si hay que invitarle al evento de su clase)  
    const calendario = CalendarApp.getCalendarById(instructor[PARAM.instructores.colIdCal - 1]);
    let guests = checkInvitarInstructores ? instructor[PARAM.instructores.colEmail - 1] : '';
    
    // const guests = instructores.find(instructor => instructor[0] == evento[2])[1] + ',' + salas.find(sala => sala[0] == evento[10])[1];

    // ¿Deseamos reservar una sala?
    if (checkReservarEspacios) {
      const sala = salas.find(sala => sala[PARAM.salas.colNombre - 1] == evento[PARAM.eventos.colAula - 1]);
      if (!sala) {
      saltados++;
      return [new Date(), '⭕ Aula no existe'];
      } else guests = `${guests},${sala[PARAM.salas.colIdCal]}`;
    }

    // ¿Tenemos todos los datos necesarios para generar el evento?
    if (title && endDate && startTime && endTime && guests && dias && calendario && !idPrevio) {

      // No es necesario que el día de la semana de startTime coincida con el 1º en la serie según la recurrencia, Calendar ajusta internamente 👏
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
      // Error genérico, faltan datos necesarios en la tabla (fila del evento)
      saltados++;
      return [new Date(), '⭕ Instructor no existe'];y
    }
  });

  hojaEventos.getRange(T_RES).setValues(resultado);
  SpreadsheetApp.getUi().alert('Eventos procesados', 'Creados: ' + creados + '\nSaltados: ' + saltados, SpreadsheetApp.getUi().ButtonSet.OK);

}