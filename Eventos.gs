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

  // Nos movemos a la hoja de gestión de eventos antes de solicitar confirmación
  const hojaActual = SpreadsheetApp.getActiveSheet();
  const hdc = SpreadsheetApp.getActive();
  const hojaEventos = hdc.getSheetByName(PARAM.eventos.hoja).activate();
  SpreadsheetApp.flush(); // ¡Necesario para que visualmente nos movamos a la hoja activa antes de generar la alerta!

  if (alerta('Se crearán eventos nuevos para las clases seleccionadas, los ya existentes serán eliminados.') == SpreadsheetApp.getUi().Button.OK) {

    // ##################################################
    // ### Preparativos y lectura de datos necesarios ###
    // ##################################################

    const hojaRegistro = hdc.getSheetByName(PARAM.registro.hoja);
    let creados = 0;
    let omitidos = 0;

    // Ajustes de generación
    const checkInvitarInstructores = hojaEventos.getRange(PARAM.eventos.checkInvitarInstructores).getValue();
    const checkReservarEspacios = hojaEventos.getRange(PARAM.eventos.checkReservarEspacios).getValue();
    const checkDesmarcar = hojaEventos.getRange(PARAM.eventos.checkDesmarcarProcesados).getValue();
    const checkBorrarPrevios = hojaEventos.getRange(PARAM.eventos.checkBorrarPrevios).getValue();
    const tag = PARAM.eventos.tag;

    // Leer datos de instructores y salas registradas
    const instructores = leerDatosHoja(hdc.getSheetByName(PARAM.instructores.hoja), PARAM.instructores.filEncabezado + 1);
    const salas = leerDatosHoja(hdc.getSheetByName(PARAM.salas.hoja), PARAM.salas.filEncabezado + 1);

    // Obtener las clases seleccionados de la tabla para las que deben generarse eventos, SE IGNORAN
    // LAS CLASES EN FILAS EN LAS QUE FALTA GRUPO O CLASE, además se anota la fila de cada clase para
    // mostrar información del resultado de la operación en su lugar correcto más tarde.
    const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
      .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
      .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true
        && eventoFila.ajustes[PARAM.eventos.colGrupo - 1] != ''
        && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

    // console.info(eventosFilas);

    if (eventosFilas && eventosFilas.length > 0) {

      // Identificar y eliminar eventos previos en Calendar para las clases que se desean procesar, 
      // eliminar eventos previos uno a uno dentro del bucle de generación no es buena idea puesto que
      // podrían producirse conflictos a la hora de reservar espacios, por ejemplo cuando dos sesiones
      // de distintas clases del mismo grupo que se imparten en la misma aula, cuyos eventos ya han sido
      // generados previamente, intercambian horas y vuelve ha realizarse la generación con ese cambio.
      // ⚠️ La contrapartida es que los eventos ya existentes se eliminarán todos a la vez, aunque podría
      // ser que posteriomente no se generasen nuevas versiones por falta de datos o errores.
      mostrarMensaje('Eliminando eventos previos asociados a las clases ✖️...');
      const previos = eliminarEventosPreviosRegistroMultiple(eventosFilas);

      // Aquí comienza realmente la fiesta
      mostrarMensaje('Generando eventos para las clases ➕...');

      if (checkBorrarPrevios) actualizarDatosTabla(hojaEventos, null, PARAM.eventos.filEncabezado + 1, PARAM.eventos.colFechaProceso);

      // ######################################
      // ### Bucle de proceso de las clases ###
      // ######################################
      //
      // [1] Construir título de los eventos.
      // [2] Obtener valores de echas / hora de inicio y fin.
      // [3] Obtener calendario público del instructor y opcionalmente privado para enviar invitación
      // [4] OPCIONAL: Obtener sala a reservar.
      // [5] Construir y generar el evento de clase en Google Calendar.
      // [6] Registrar resulado e información de la clase y del evento generado en la tabla de registro de eventos.
      // [7] Mostrar resultado de la operación en la tabla de gestión de eventos.
      //
      // Como en principio no serán muchas opto por actualizar la hoja de datos cada vez que se procesa un evento,
      // a sabiendas de que no es óptimo y ralentizará el proceso. Usaremos un sello de tiempo común para
      // diferenciar todos los eventos generados en cada proceso de los generados en procesos anteriores, de este
      // modo se podrán eliminar de manera segura sin "pisar" otras sesiones (eventos) de la misma clase incluidos
      // en las clases incluidias en el proceso actual.

      const selloTiempoProceso = new Date();
      eventosFilas.forEach(eventoFila => {

        // Datos de cada clase leída de la hoja de generación de eventos que va a procesarse
        const evento = eventoFila.ajustes;

        // Objeto que contiene el resultado de la operación de generación del evento en calendar:
        // { idEvento, idCalendario, selloTiempo, mensaje }
        let resultado;

        // #####################################################################
        // ### [1] Título del evento: Grupo + Clase + (iniciales instructor) ###
        // #####################################################################

        const title = `${evento[PARAM.eventos.colGrupo - 1]} ${evento[PARAM.eventos.colClase - 1]} (${evento[PARAM.eventos.colInstructor - 1]})`;

        // ########################################################################################################
        // ### [2] Gestión de fecha/hora de INICIO y FIN de la primera repetición y fecha(+ ajuste hora) de FIN ###
        // ########################################################################################################

        // ⚠️ Marcianada al leer celdas con datos de hora sin fecha:
        // Ej, celda: 17:00:00 (formateada como hora) >> Objeto Date:Sat Dec 30 1899 17:24:05 GMT+0009 (Central European Standard Time) 😵‍💫
        // https://twitter.com/pfelipm/status/1572253070930878464
        // Posible solución:
        //   - Leer valor de fecha normalmente con getValue() y de hora con getDisplayValue()
        //   - Combinar haciendo eventStart = new Date(fecha.setHours(fechaComoTexto.split(':')[0], fechaComoTexto.split(':')[1]))
        // Paso de movidas y "monto" fecha + hora mediante fórmulas en la hoja de cálculo, eso me evita leer la tabla de dos modos distintos.
        const startTime = evento[PARAM.eventos.colStartTime - 1];
        const endTime = evento[PARAM.eventos.colEndTime - 1];

        // Además, tenemos esta otra marcianada https://issuetracker.google.com/issues/236615807
        // ...siguiendo el mismo criterio que con startTime y endTime, compongo mediante fórmulas en la tabla
        // de eventos y leo aquí una fecha de fin de repetición como: dia_fin_repetición + hora_inicio para
        // utilizar a la hora de definir la recurrencia de la sesión (evento).
        const endDateTime = evento[PARAM.eventos.colEndDateTime - 1];

        // ⚠️ Si cadena vacía, split() devuelve un array que contiene una cadena vacía (en lugar de un array vacío)
        const dias = evento[PARAM.eventos.colDias - 1].split(PARAM.eventos.separadorDias);
        const descripcion = evento[PARAM.eventos.colDescripcion - 1];

        // ### AQUÍ LA GENERACIÓN DE EVENTOS ######

        // Se usa try para tratar situaciones que no permiten generar el evento como excepciones, evitando IFs...

        try {

          // #####################################################################
          // ### [3] Gestión de calendarios de INSTRUCTORES (público, privado) ###
          // #####################################################################

          let guests = '';

          // a) ¿La clase tiene instructor?
          const instructorClase = evento[PARAM.eventos.colInstructor - 1];
          if (!instructorClase) throw '⭕ Falta instructor';

          // b) ¿Localizamos en la tabla de instructores las iniciales del asignado a esta clase?
          const infoInstructor = instructores.find(instructor => instructor[PARAM.instructores.colIniciales - 1] == instructorClase);
          if (!infoInstructor) throw '⭕ Instructor no registrado';

          // c) ¿Existe en la tabla el ID de su calendario público donde se generará el evento?
          const idCalendario = infoInstructor[PARAM.instructores.colIdCal - 1];
          if (!idCalendario) throw '⭕ Falta ID calendario instructor';

          // d) ¿El calendario realmente existe?
          const calendario = CalendarApp.getCalendarById(idCalendario);
          if (!calendario) throw '⭕ Calendario instructor innacesible';

          // e) ¿Se debe enviar una invitación al instructor? (opcional)

          if (checkInvitarInstructores) {
            guests = infoInstructor[PARAM.instructores.colEmail - 1];
            // Por defecto no se lanza excepción cuando falta email instructor para enviar invitación,
            // Parametrizable mediante constante PARAM.permitirOmitirEmailInstructor.
            if (!guests && !PARAM.permitirOmitirEmailInstructor) throw '⭕ Falta email instructor';
          }


          // ################################################
          // ### [4] Gestión de la reserva de AULA (sala) ###
          // ################################################  
          // Para acomodar ciertos casos límite en mi centro, *por defecto* solo lanzaremos una excepción si se cumplen todas:
          //   a) Se ha indicado que se deben reservar espacios.
          //   b) El aula indicada para la clase no es una cadena vacía.
          //   c) El aula indicada para la clase no se encuentra en la tabla de salas o falta ID o calendario innacesible
          // Parametrizable mediante constante PARAM.permitirOmitirSala.
          if (checkReservarEspacios) {

            const aulaClase = evento[PARAM.eventos.colAula - 1];
            if (!aulaClase && !PARAM.permitirOmitirSala) throw '⭕ Falta aula';

            if (aulaClase) {

              const infoSala = salas.find(sala => sala[PARAM.salas.colNombre - 1] == aulaClase);
              if (!infoSala) throw '⭕ Aula no registrada';

              const idSala = infoSala[PARAM.salas.colIdCal - 1];
              if (!idSala) throw '⭕ Falta ID calendario aula';

              const testSala = CalendarApp.getCalendarById(idSala);
              if (!testSala) throw '⭕ Calendario aula inaccesible';

              guests = guests ? `${guests},${idSala}` : `${infoSala[PARAM.salas.colIdCal - 1]}`;

            }

          }

          // ##############################################################
          // ### [5] Generación de evento recurrente en Google Calendar ###
          // ##############################################################

          // Comprobaciones previas a la generación del evento (algunas ya se fuerzan en la hoja de gestión de eventos)
          if (!evento[PARAM.eventos.colHoraInicio - 1] || !startTime) throw '⭕ Falta hora inicio';
          if (!evento[PARAM.eventos.colHoraFin - 1] || !endTime) throw '⭕ Falta hora fin';
          if (!evento[PARAM.eventos.colDiaFinRep - 1] || !endDateTime) throw '⭕ Falta fecha fin';
          if (endTime <= startTime) throw '⭕ Hora fin < hora inicio';
          if (evento[PARAM.eventos.colDiaFinRep - 1] < evento[PARAM.eventos.colDiaInicioRep - 1]) throw '⭕ Día fin ≤ Día inicio';
          if (!dias[0]) throw '⭕ Falta días semana repetición';

          // ⚠️ Es necesario que el día de la semana de startTime coincida con uno de los indicados en la regla de recurrencia,
          // de lo contrario se genera una repetición fantasma en el día indicado, aunque no forme parte de los usados en dicha
          // regla. ¡Esto no ocurre cuando se crean eventos periódicos manualmente desde Calendar!
          // 👍 Con la fecha de finalización de la recurrencia no hay problema, las repeticiones finalizan cuando corresponde.

          const recurrence = CalendarApp.newRecurrence()
            //.setTimeZone(Session.getTimeZone())
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
            })).until(endDateTime);

          // Los eventos de clases se crearán en el calendario público del instructor,
          // en su caso invitando a la sala y al propio instructor.
          // ➕ Mejora: permitir reserva de múltiples salas por evento.

          // Eliminar posibles eventos ya creados previamente para este grupo y clase,
          // esto está descartado y procede de una versión preliminar del script, te lo dejo
          // por si te apetece jugar con esta alternativa funcional.
          // const previos = eliminarEventosPreviosRegistro(evento[PARAM.eventos.colGrupo - 1], evento[PARAM.eventos.colClase - 1], selloTiempoProceso);

          // ...y ahora generamos los nuevos
          const eventoCalendar = calendario.createEventSeries(title, startTime, endTime, recurrence,
            {
              description: descripcion,
              guests: guests,
              // ➕ Mejora: ajuste para invitar a instructores pero no enviarles invitaciones
              sendInvites: checkInvitarInstructores
            });
          eventoCalendar.setTag(tag, tag); // Por ahora no se usa para nada
          resultado = {
            idEvento: eventoCalendar.getId(),
            idCalendario: idCalendario,
            selloTiempo: new Date(),
            // Esto tampoco se utiliza ya, se usaba cuando se borraban los eventos previosgenerado uno a uno
            // mensaje: previos > 0 ? `🟣 Evento actualizado [${previos}]` : '🟢 Evento creado',
            mensaje: '🟢 Evento generado',
            fila: eventoFila.fila
          };
          creados++;

        } catch (e) {

          // Capturar mensaje de excepción, controlada o no
          if (typeof e == 'string') resultado = { selloTiempo: new Date(), mensaje: e }
          else resultado = { selloTiempo: new Date(), mensaje: `🚨 ${e.message}` };
          omitidos++;

        } finally {

          // ####################################################################################################
          // ### [6] Registrar resultado de la operación y archivar evento en la tabla de registro de eventos ###
          // ####################################################################################################

          // a) Resultado en columnas Fecha proceso y Resultado de la tabla de gestión de eventos
          hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colFechaProceso, 1, 2)
            .setValues([[resultado.selloTiempo, resultado.mensaje]]);

          // b) Si el evento se ha podido crear, archivar en tabla de registro de eventos
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

        }

      });

      SpreadsheetApp.flush();

      // Preparar botón de conmutación general de selección para que ACTIVE todas las casillas de verificación si se han desmarcado
      if (checkDesmarcar) PropertiesService.getScriptProperties().setProperty(PARAM.propiedadEstadoCheck, false);

      // Resumen del resultado de la operación
      mostrarMensaje('Proceso terminado.', 5);
      alerta(
        '🟢 Generados: ' + creados + '\n⭕ Omitidos: ' + omitidos + '\n✖️ Eliminados: ' + previos,
        SpreadsheetApp.getUi().ButtonSet.OK,
        'Eventos procesados');

    } else mostrarMensaje('No se han seleccionado clases.');

  } else hojaActual.activate();

}

/**
 * Función invocada desde el menú del script.
 * Consulta la existencia de los eventos seleccionados en la tabla de registro de la hoja
 * `PARAM.registro.hoja` e intenta eliminarlos, tanto de la hoja de registro como de los
 * calendarios en los que se tiene constancia de que han sido creados previamente.
 */
function m_EliminarEventos() {

  // Nos movemos a la hoja de gestión de eventos antes de solicitar confirmación
  const hojaActual = SpreadsheetApp.getActiveSheet();
  const hojaEventos = SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se eliminarán los eventos asociados a las clases seleccionadas.') == SpreadsheetApp.getUi().Button.OK) {

    const checkDesmarcar = hojaEventos.getRange(PARAM.eventos.checkDesmarcarProcesados).getValue();
    const checkBorrarPrevios = hojaEventos.getRange(PARAM.eventos.checkBorrarPrevios).getValue();
    let eliminados = 0;
    let noHallados = 0;

    // Obtener las clases seleccionadas de la tabla cuyos eventos deben tratarse de eliminar,
    // SE IGNORAN LAS CLASES EN FILAS EN LAS QUE FALTA GRUPO O CLASE,
    const eventosFilas = leerDatosHoja(hojaEventos, PARAM.eventos.filEncabezado + 1)
      .map((evento, indice) => { return { ajustes: evento, fila: indice + 1 } })
      .filter(eventoFila => eventoFila.ajustes[PARAM.eventos.colCheck - 1] == true
        && eventoFila.ajustes[PARAM.eventos.colGrupo - 1] != ''
        && eventoFila.ajustes[PARAM.eventos.colClase - 1] != '');

    if (eventosFilas && eventosFilas.length > 0) {

      mostrarMensaje('Eliminando eventos previos asociados a las clases ✖️....');

      if (checkBorrarPrevios) actualizarDatosTabla(hojaEventos, null, PARAM.eventos.filEncabezado + 1, PARAM.eventos.colFechaProceso);

      const selloTiempoProceso = new Date();
      eventosFilas.forEach(eventoFila => {

        // Fila con información de cada evento leída de la hoja de eventos
        const evento = eventoFila.ajustes;

        const instanciasEliminadas = eliminarEventosPreviosRegistro(
          evento[PARAM.eventos.colGrupo - 1],
          evento[PARAM.eventos.colClase - 1],
          selloTiempoProceso
        );

        // Actualizar tabla (columnas Fecha proceso y Resultado)
        hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colFechaProceso, 1, 2)
          .setValues([[new Date(), instanciasEliminadas > 0 ? `✖️ Eliminado [${instanciasEliminadas}]` : '❔ No existe o ya eliminado']]);

        // Desmarcar selección, si se ha seleccionado esa opción...
        if (instanciasEliminadas > 0) {
          eliminados++;
          if (checkDesmarcar) hojaEventos.getRange(PARAM.eventos.filEncabezado + eventoFila.fila, PARAM.eventos.colCheck, 1, 1).setValue(false);
        } else noHallados++;

      });

      SpreadsheetApp.flush();

      // Preparar botón de conmutación general de selección para que ACTIVE todas las casillas de verificación si se han desmarcado
      if (checkDesmarcar) PropertiesService.getScriptProperties().setProperty(PARAM.propiedadEstadoCheck, false);

      // Resumen del resultado de la operación
      mostrarMensaje('Proceso terminado.', 5);
      alerta('✖️ Eliminados: ' + eliminados + '\n❔ No existen / ya eliminados: ' + noHallados, SpreadsheetApp.getUi().ButtonSet.OK, 'Eventos procesados');

    } else mostrarMensaje('No se han seleccionado clases.');

  } else hojaActual.activate();

}

/**
 * Elimina la información de marca de tiempo y resultado del proceso en las columnas de
 * la tabla de gestión de eventos.
 */
function m_EliminarResultados() {

  // Nos movemos a la hoja de gestión de eventos antes de solicitar confirmación
  const hojaActual = SpreadsheetApp.getActiveSheet();
  hojaEventos = SpreadsheetApp.getActive().getSheetByName(PARAM.eventos.hoja).activate();
  SpreadsheetApp.flush();

  if (alerta('Se borrarán todas las marcas de tiempo y resultados de la tabla.') == SpreadsheetApp.getUi().Button.OK) {
    actualizarDatosTabla(hojaEventos, null, PARAM.eventos.filEncabezado + 1, PARAM.eventos.colFechaProceso);
  } else hojaActual.activate();

}