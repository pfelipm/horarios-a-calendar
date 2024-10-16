/***************************************************************************
 * Horarios-a-Calendar (HaC)                                               *
 *                                                                         *
 * Una herramienta que te ayuda a transformar horarios de clase en eventos *
 * recurrentes en Google Calendar.                                         *
 *                                                                         *
 * Copyright (C) Pablo Felip (@pfelipm) v1.11 OCT 2024                      *
 * Se distribuye bajo licencia GNU GPL v3.                                 *
 *                                                                         *
 ***************************************************************************
 *
 * 16/10/24 v1.11 | Pequeños ajustes cosméticos:
 *   - Comandos de menú para descargar calendarios de docentes y aulas renombrados ("Obtener...").
 *   - Texto "Última ejecución" renombrado a "Última obtención" en hojas 👥 Instructores y 🏫 Salas.
 *   - Nota sobre hoja 👥 Instructores → E5 para indicar que solo se descargan calendarios visibles en la lista.
 * 
 * @OnlyCurrentDoc
 */


/**
 * Constantes generales de parametrización del script
 */
const PARAM = {

  nombre: 'HaC',
  version: 'Versión 1.11 (octubre 2024)',
  icono: '🗓️',
  urlRepoGitHub: 'https://github.com/pfelipm/horarios-a-calendar',
  propiedadEstadoCheck: 'estadoCheck01',

  // Constantes funcionales
  permitirOmitirEmailInstructor: true, // FALSE si se desea lanzar excepción cuando se invita a instructores pero falta puntualmente un email
  permitirOmitirSala: true, // FALSE si se desea lanzar excepción cuando se reservan espacios pero puntualmente no se ha asignado aula

  // Hoja (oculta) que contiene la plantilla del horario
  plantillaHorario: {
    hoja: '⏰ Plantilla horario',
    codigoGrupo: 'D5',
    longMaxCodigo: 8
  },

  // Tabla de eventos
  eventos: {
    hoja: '➕ Gestión eventos',
    separadorDias: ',',
    filEncabezado: 6,
    colCheck: 1,
    colGrupo: 2,
    colClase: 3,
    colDias: 4,
    colHoraInicio: 5,
    colHoraFin: 6,
    colInstructor: 7,
    colAula: 8,
    colDiaInicioRep: 9,
    colDiaFinRep: 10,
    colDescripcion: 11,
    colStartTime: 12,
    colEndTime: 13,
    colEndDateTime: 14,
    colFechaProceso: 15,
    colResultado: 16,
    checkInvitarInstructores: 'G2',
    checkReservarEspacios: 'G3',
    checkDesmarcarProcesados: 'J2',
    checkBorrarPrevios: 'J3',
    tag: '#HaC' // Utilizado para marcar los eventos creados, sin utilidad efectiva
  },

  // Tabla de instructores
  instructores: {
    hoja: '👥 Instructores',
    filEncabezado: 5,
    colIniciales: 2,
    colEmail: 3,
    colNombreCal: 4,
    colIdCal: 5,
    colNombreCalObtenido: 6,
    prefijo: 'B2',
    ultEjecucion: 'E2'
  },

  // Tabla de salas (aulas)
  salas: {
    hoja: '🏫 Salas',
    filEncabezado: 4,
    colNombre: 1,
    colIdCal: 5,
    ultEjecucion: 'E2'
  },

  // Tabla de registro de eventos
  registro: {
    hoja: '📦 Registro eventos',
    filEncabezado: 2,
    colGrupo: 1,
    colClase: 2,
    colDias: 3,
    colHoraInicio: 4,
    colHoraFin: 5,
    colInstructor: 6,
    colAula: 7,
    colDiaInicioRep: 8,
    colDiaFinRep: 9, 
    colFechaProceso: 10,
    colIdEv: 11,
    colIdCal: 12
  }

};