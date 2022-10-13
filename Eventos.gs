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
 * los eventos generados y tener que hacer una búsqueda por fuerza bruta en todos los calendarios de instructores.
 */
function m_CrearEventos() {
  
  const hojaActual = SpreadsheetApp.getActiveSheet();
  hoja = SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se crearán o actualizarán eventos para las clases seleccionadas.') == SpreadsheetApp.getUi().Button.OK) {

    const hdc = SpreadsheetApp.getActive();
    const hojaEventos = hdc.getSheetByName(PARAM.eventos.hoja);
    const hojaRegistro = hdc.getSheetByName(PARAM.registro.hoja);
    let creados = 0;
    let omitidos = 0;

    // Ajustes de generación
    const checkInvitarInstructores = hojaEventos.getRange(PARAM.eventos.checkInvitarInstructores).getValue();
    const checkReservarEspacios = hojaEventos.getRange(PARAM.eventos.checkReservarEspacios).getValue();
    const checkDesmarcar = hojaEventos.getRange(PARAM.eventos.checkDesmarcarProcesados).getValue();
    const checkBorrarPrevios = hojaEventos.getRange(PARAM.eventos.checkBorrarPrevios).getValue();
    const tag = PARAM.eventos.tag;

    // Leer datos necesarios para la generación de los eventos en Calendar.
    const instructores = leerDatosHoja(hdc.getSheetByName(PARAM.instructores.hoja), PARAM.instructores.filEncabezado + 1);
    const salas = leerDatosHoja(hdc.getSheetByName(PARAM.salas.hoja), PARAM.salas.filEncabezado + 1);

    mostrarMensaje('Generado eventos...');

    // Obtener las clases seleccionados de la tabla para las que deben generarse eventos, se descartan
    // las clases en filas en las que falta grupo o clase, además se anota la fila de cada clase para
    // mostrar información del resultado de la operación en su lugar correcto más tarde.
    const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
      .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
      .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true
        && eventoFila.ajustes[PARAM.eventos.colGrupo - 1] != ''
        && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

    // console.info(eventosFilas);

    if (checkBorrarPrevios) actualizarDatosTabla(hojaEventos, null, PARAM.eventos.filEncabezado + 1, PARAM.eventos.colFechaProceso + 1);

    // Vamos a crear eventos. Como en principio no serán muchos opto
    // por actualizar la hoja de datos cada vez que se crea un evento,
    // a sabiendas de que no es lo óptimo, en general.
    eventosFilas.forEach(eventoFila => {

      // Fila con información de cada evento leída de la hoja de eventos
      const evento = eventoFila.ajustes;

      // Objeto que contiene el resultado de la operación de generación del evento en calendar
      // { idEvento, idCalendario, selloTiempo, mensaje }
      let resultado;

      // Título del evento: Grupo + Clase + (iniciales instructor)
      const title = `${evento[PARAM.eventos.colGrupo - 1]} ${evento[PARAM.eventos.colClase - 1]} (${evento[PARAM.eventos.colInstructor - 1]})`;
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
      // ⚠️ Si cadena vacía, split() devuelve un array que contiene una cadena vacía (en lugar de un array vacío)
      const dias = evento[PARAM.eventos.colDias - 1].split('-');
      const descripcion = evento[PARAM.eventos.colDescripcion - 1];

      // Se usa try para tratar situaciones que no permiten generar el evento como excepciones, evitando IFs...
      try {

        // ¿Localizamos en la tabla de instructores las iniciales del asignado a esta clase?
        const instructor = instructores.find(instructor => instructor[PARAM.instructores.colIniciales - 1] == evento[PARAM.eventos.colInstructor - 1]);
        if (!instructor) throw '⭕ Instructor no existe';

        // Si es que sí, obtenemos su calendario público y su calendario privado (si hay que invitarle al evento de su clase)
        const idCalendario = instructor[PARAM.instructores.colIdCal - 1];
        const calendario = CalendarApp.getCalendarById(idCalendario);

        // Comprobar: qué pasa con la coma al añadir sala si email_instructor = ''
        let guests = checkInvitarInstructores ? instructor[PARAM.instructores.colEmail - 1] : '';

        // ¿Deseamos reservar una sala?
        if (checkReservarEspacios) {
          const sala = salas.find(sala => sala[PARAM.salas.colNombre - 1] == evento[PARAM.eventos.colAula - 1]);
          if (!sala) throw '⭕ Aula no existe';
          else guests = `${guests},${sala[PARAM.salas.colIdCal]}`;
        }

        // ¿Tenemos todos los datos necesarios para generar el evento?
        if (!endDate) throw '⭕ Falta fecha fin';
        if (!startTime) throw '⭕ Falta hora inicio';
        if (!endTime) throw '⭕ Falta hora fin';
        if (!dias[0]) throw '⭕ Falta días';
        if (!calendario) throw '⭕ Falta calendario instructor';

        // Aquí toca comprobar si esa sesión (GRUPO, CLASE) ya se ha generado POR HACER

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

        // Los eventos de clases se crearán en el calendario público del instructor,
        // en su caso invitando a la sala y al propio instructor (mejora: múltiples salas).

        resultado = {
          idEvento: 'ID_FALSO',
          idCalendario: idCalendario,
          selloTiempo: new Date(),
          mensaje: '🟢 Evento simulado creado',
          fila: eventoFila.fila };

        // Eliminar posibles eventos ya creados previamente para este grupo y clase, no se usa el valor devuelto (nº eliminados)
        eliminarEventosPreviosRegistro(evento[PARAM.eventos.colGrupo - 1], evento[PARAM.eventos.colClase - 1]);
        
        /*
        const eventoCalendar = calendario.createEventSeries(title, startTime, endTime, recurrence,
          {
            description: descripcion,
            guests: guests,
            sendInvites: checkInvitarInstructores
          });
        eventoCalendar.setTag(tag, tag); // Por ahora no se usa para nada
        resultado = {
          id: eventoCalendar.getId(),
          idCalendario: idCalendario,
          selloTiempo: new Date(),
          mensaje: '🟢 Evento creado',
          fila: eventoFila.fila };
        }
        */
        creados++;

      } catch (e) {

        if (typeof e == 'string') resultado = { selloTiempo: new Date(), mensaje: e }
        else resultado = { selloTiempo: new Date(), mensaje: `🚨 ${e.message}` };
        omitidos++;

      } finally {

        // Actualizar tabla (columnas Fecha proceso y Resultado)
        hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colFechaProceso, 1, 2)
          .setValues([[resultado.selloTiempo, resultado.mensaje]]);

        // Si el evento se ha podido crear...
        if (resultado.idEvento) {

          // Desmarcar selección, si se ha seleccionado esa opción...
          if (checkDesmarcar) hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colCheck, 1, 1).setValue(false);

          // ...y guardar evento en tabla de registro de eventos (se preserva el orden de las columnas)
          if (hojaRegistro.getLastRow() > PARAM.registro.filEncabezado) hojaRegistro.insertRowBefore(PARAM.registro.filEncabezado + 1);
          const eventoRegistro = [
            evento[PARAM.eventos.colGrupo - 1],
            evento[PARAM.eventos.colClase - 1],
            evento[PARAM.eventos.colDias - 1],
            evento[PARAM.eventos.colHoraInicio - 1],
            evento[PARAM.eventos.colHoraFin - 1],
            evento[PARAM.eventos.colInstructor - 1],
            evento[PARAM.eventos.colAula - 1],
            evento[PARAM.eventos.colDiaInicioRep - 1],
            evento[PARAM.eventos.colDiaFinRep - 1],
            resultado.selloTiempo,
            resultado.idEvento,
            resultado.idCalendario
          ];
          hojaRegistro.getRange(PARAM.registro.filEncabezado + 1, 1, 1, eventoRegistro.length).setValues([eventoRegistro]);

        }

        SpreadsheetApp.flush();

      }

    });

    // Resumen del resultado de la operación
    mostrarMensaje('Proceso terminado.', 2);
    alerta('🟢 Creados: ' + creados + '\n🟠 Omitidos: ' + omitidos,  SpreadsheetApp.getUi().ButtonSet.OK, 'Eventos procesados');
  
  } else hojaActual.activate();

}

/**
 * Función invocada desde el menú del script.
 * Consulta la existencia de los eventos seleccionados en la tabla de registro de la hoja
 * `PARAM.registro.hoja` e intenta eliminarlos, tanto de la hoja de registro como de los
 * calendarios en los que se tiene constancia de que han sido creados previamente.
 */
function m_EliminarEventos() {

  const hojaActual = SpreadsheetApp.getActiveSheet();
  hoja = SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se eliminarán los eventos asociados a las clases seleccionadas.') == SpreadsheetApp.getUi().Button.OK) {

    mostrarMensaje('Eliminando clases...');

    const hojaEventos = SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja);
    const checkDesmarcar = hojaEventos.getRange(PARAM.eventos.checkDesmarcarProcesados).getValue();
    const checkBorrarPrevios = hojaEventos.getRange(PARAM.eventos.checkBorrarPrevios).getValue();
    let eliminados = 0;
    let noHallados = 0;

    // Obtener las clases seleccionadas de la tabla cuyos eventos deben tratarse de eliminar,
    // se descartan las clases en las filas en las que falta grupo o clase.
    const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
      .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
      .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true
        && eventoFila.ajustes[PARAM.eventos.colGrupo - 1] != ''
        && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

    if (checkBorrarPrevios) actualizarDatosTabla(hojaEventos, null, PARAM.eventos.filEncabezado + 1, PARAM.eventos.colFechaProceso + 1);

    eventosFilas.forEach(eventoFila => {

      // Fila con información de cada evento leída de la hoja de eventos
      const evento = eventoFila.ajustes;

      const instanciasEliminadas = eliminarEventosPreviosRegistro(evento[PARAM.eventos.colGrupo - 1], evento[PARAM.eventos.colClase - 1]);
            
      // Actualizar tabla (columnas Fecha proceso y Resultado)
      hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colFechaProceso, 1, 2)
        .setValues([[new Date(), instanciasEliminadas > 0 ? '✖️ Evento eliminado' : '⭕ Evento no encontrado']]);

      // Desmarcar selección, si se ha seleccionado esa opción...
      if (instanciasEliminadas > 0) {
        eliminados++;
        if (checkDesmarcar) hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colCheck, 1, 1).setValue(false);
      } else noHallados++;

      SpreadsheetApp.flush();

    });

    // Resumen del resultado de la operación
    mostrarMensaje('Proceso terminado.', 2);
    alerta('✖️ Eliminados: ' + eliminados + '\n⭕ No encontrados: ' + noHallados, SpreadsheetApp.getUi().ButtonSet.OK, 'Eventos procesados');
    
  } else hojaActual.activate();
  
}