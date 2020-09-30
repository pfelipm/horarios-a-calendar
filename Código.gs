/**
 * Extrae los eventos (clases) a partir de un horario representado en un rango especificado
 * Fila 1: Días de la semana.
 * Columna 1: Hora de inicio de la clase.
 * Columna N: Hora de fin de la clase.
 * @param {A1:G13} horario Intervalo de datos que contiene el horario.
 * @param {FALSO} agrupar Indica si se deben tratar de agrupar las clases que se repiten
 *                a lo largo de la semana en el mismo horario.
 * @return Tabla de clases.
 *
 * @CustomFunction
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer(@pfelipm)
 */

function EXTRAERCLASES(horario, agrupar = false) {
  
  // Comprobar parámetros
  
  // Inicializaciones varias
  
  const numDias = horario[0].length - 2;
  const numFranjas = horario.length - 1;
  let eventos = [];
  let tablaEventos = [];
  
  // Etapa 1: identificar eventos
 
  for (let dia = 1; dia <= numDias; dia++){
    
    // Primera clase del día
    
    let eventoActual = {desc: horario[1][dia], dia: horario[0][dia], hInicio: horario[1][0], hFin: horario[1][horario[0].length - 1] };
    
    for (let franjaHoraria = 2; franjaHoraria <= numFranjas ; franjaHoraria++){
      
      if (eventoActual.desc == horario[franjaHoraria][dia]) {
        eventoActual.hFin = horario[franjaHoraria][horario[0].length - 1];
      }
      else {
 
        // Añadir evento
          
         eventos.push(eventoActual);
         eventoActual = {desc: horario[franjaHoraria][dia], dia: horario[0][dia],
                         hInicio: horario[franjaHoraria][0],
                         hFin: horario[franjaHoraria][horario[0].length - 1] };
      }
    }
    
    // Registrar el último evento del día
    
    eventos.push(eventoActual);
  }
  
  // Etapa 2: Agrupar eventos que se repiten en varios días a la semana
  
  // Generar matriz resultado
    
  return eventos.map(evento => [evento.desc, evento.dia, evento.hInicio, evento.hFin]);
    
}