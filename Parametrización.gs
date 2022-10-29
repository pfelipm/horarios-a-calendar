/**************************************************************************
 * Horarios-a-Calendar (HaC)                                              *
 * Una herramienta que te ayuda a transformar horarios de clase a eventos *
 * recurrentes en Google Calendar                                         *
 * Pablo Felip (@pfelipm)                                                 *
 **************************************************************************
 *
 * @OnlyCurrentDoc
 */


/**
 * Constantes generales de parametrización del script
 */
const PARAM = {
  nombre: 'HaC',
  version: 'Versión 1.0 (octubre 2022)',
  icono: '🗓️',
  urlRepoGitHub: 'https://github.com/pfelipm/horarios-a-calendar',
  propiedadEstadoCheck: 'estadoCheck01',

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
    tag: '#HaC' // Utilizado para marcar los eventos creados, sin uso efectivo
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
    colFechaProceso: 10,
    colIdEv: 11,
    colIdCal: 12
  }
};