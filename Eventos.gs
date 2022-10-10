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
  const tag = PARAM.eventos.tag;

  // Leer datos necesarios para la geneneración de los eventos en Calendar.
  const instructores = leerDatosHoja(hdc.getSheetByName(PARAM.instructores.hoja), PARAM.instructores.filEncabezado + 1);
  const salas = leerDatosHoja(hdc.getSheetByName(PARAM.salas.hoja), PARAM.salas.filEncabezado + 1);
  
  // Como hay casillas de verificación en la columna 1 es necesario descartar las filas en las que no hay eventos (sin Clase)
  const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
    .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
    .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

  console.info(eventosFilas);

  let creados = 0;
  let omitidos = 0;

  // Vamos a crear eventos...
  const resultados = eventosFilas.map(eventoFila => {

    const evento = eventoFila.ajustes;

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
    const descripcion = evento[PARAM.eventos.colDescripcion - 1];

    // ¿Localizamos en la tabla de instructores las iniciales del asignado a esta clase?
    const instructor = instructores.find(instructor => instructor[PARAM.instructores.colIniciales - 1] == evento[PARAM.eventos.colInstructor - 1]);
    if (!instructor) {
      omitidos++;
      return { selloTiempo: new Date(), mensaje: '⭕ Instructor no existe', fila: eventoFila.fila };
    }

    // Si es que sí, obtenemos su calendario público y su calendario privado (si hay que invitarle al evento de su clase)  
    const calendario = CalendarApp.getCalendarById(instructor[PARAM.instructores.colIdCal - 1]);

    // Comprobar: qué pasa con la coma al añadir sala si email_instructor = ''
    let guests = checkInvitarInstructores ? instructor[PARAM.instructores.colEmail - 1] : '';

    // ¿Deseamos reservar una sala?
    if (checkReservarEspacios) {
      const sala = salas.find(sala => sala[PARAM.salas.colNombre - 1] == evento[PARAM.eventos.colAula - 1]);
      if (!sala) {
        omitidos++;
        return { selloTiempo: new Date(), mensaje: '⭕ Aula no existe', fila: eventoFila.fila };
      } else guests = `${guests},${sala[PARAM.salas.colIdCal]}`;
    }

    // ¿Tenemos todos los datos necesarios para generar el evento?
    if (title && endDate && startTime && endTime && dias && calendario) {

      // Aquí toca comprobar si esa sesión (GRUPO, CLASE) ya se ha generado

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

      // Los eventos de clases se crearán en el calendario público del instructor,
      // en su caso invitando a la sala y al propio instructor (mejora: múltiples salas).
      let evento;

      return { id: 'ID_FALSO', selloTiempo: new Date(), mensaje: '🟢 Evento simulado creado', fila: eventoFila.fila };

      /*try {
        const eventoCalendar = calendario.createEventSeries(title, startTime, endTime, recurrence,
          {
            description: descripcion,
            guests: guests,
            sendInvites: checkInvitarInstructores
          });
        // Por ahora no se usa para nada
        eventoCalendar.setTag(tag, tag);
        return {
            id: eventoCalendar.getId(),
            selloTiempo: new Date(),
            mensaje: '🟢 Evento creado'
        };
      } catch (e) {
        omitidos++;
        return { selloTiempo: new Date(), mensaje: `🚨 ${e.message}`};
      }*/


    } else {
      // Error genérico, faltan datos necesarios en la tabla distintos a calendarios de sala o público del instructor
      omitidos++;
      return { selloTiempo: new Date(), mensaje: '⭕ Datos incompletos', fila: eventoFila.fila };
    }
  });

  console.info(resultados);

  /* No usado, opto por escribir resultado tras cada operación, facilita las cosas con los eventos no seleccionados
  actualizarDatosTabla(
    hojaEventos,
    resultados.map(evento => [evento.selloTiempo, evento.mensaje]),
    PARAM.eventos.filEncabezado + 1,
    PARAM.eventos.colFechaProceso);
  */

  // Copiar información de la sesión a tabla de Registro de eventos

  // Alerta informativa final
  SpreadsheetApp.getUi().alert('Eventos procesados', 'Creados: ' + creados + '\nSaltados: ' + omitidos, SpreadsheetApp.getUi().ButtonSet.OK);

}