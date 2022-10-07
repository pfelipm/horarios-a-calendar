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

const T_EVENTOS = 'A3:M11';
const T_RES = 'L3:M11';
const T_CAL_PROF = 'O3:Q11';
const T_CAL_AULA = 'S3:T11';

const PARAM = {
  nombre: 'HaC',
  version: 'Versión 1.0 (octubre 2022)',
  icono: '🗓️',
  urlRepoGitHub: 'https://github.com/pfelipm/horarios-a-calendar',

  // Tabla de eventos
  eventos: {
    hoja: 'Gestión eventos',
    filEncabezado: 6,
    colCheck: 1,
    colGrupo: 2,
    colClase: 3,
    colDias: 4,
    colHoraInico: 5,
    colHoraFin: 6,
    colStartTime: 7,
    colEndTime: 8,
    colDiaInicioRep: 9,
    colDiaFinRep: 10,
    colDocente: 11,
    colIdAula: 12,
    colFechaProces: 13,
    colIdSerieEvento: 14
  },

  // Tabla de instructores
  instructores: {
    hoja: 'Instructores',
    filEncabezado: 5,
    colIniciales: 2,
    colEmail: 3,
    colNombreCal: 4,
    colIdCal: 5,
    prefijo: 'B2',
    ultEjecucion: 'E2'
  },

  // Tabla de salas (aulas)
  salas: {
    hoja: 'Salas',
    filEncabezado: 4,
    colNombre: 1,
    colIdCal: 5,
    colDatos: 1,
    ultEjecucion: 'E2'
  }
};