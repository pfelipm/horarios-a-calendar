/**
 * Extrae los eventos (clases) del horario representado en un intervalo de datos especificado en formato tabla.
 * @param {A1:G13} horario Intervalo de datos que contiene el horario:
 *                         [Fila 1]: Etiqueta de cada sesión, típicamente los días de la semana.
 *                         [Columna inicial]: Hora de inicio del evento.
 *                         [Columna final]: Hora de fin del evento.
 *                         [Celdas interiores]: Descripción del evento.
 * @param {FALSO} agrupar Indica si se deben tratar de agrupar los eventos que se repiten
 *                        a lo largo de la semana en el mismo horario (VERDADERO | FALSO).
                          Si se omite se asume FALSO.
 * @param {'-'} separador Secuencia de caracteres utilizada para separar la etiqueta de cada sesión,
 *                        en el caso de que se haya solicitado su agrupación. Si se omite se
 *                        concatenan las etiquetas sin más.
 * @return Tabla de clases.
 *
 * @CustomFunction
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer(@pfelipm)
 */

function EXTRAEREVENTOS(horario, agrupar = false, separador = '') {
  
  // Comprobar parámetros (implementar)
  
  if (!Array.isArray(horario)) throw 'No se ha indicado un intervalo de datos o no es del tipo correcto.';
  if (typeof agrupar != 'boolean') throw('El parámetro "agrupar" debe ser VERDADERO o FALSO');
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto.';
  
  // Inicializaciones varias
  
  const numDias = horario[0].length - 2;
  const numFranjas = horario.length - 1;
  let eventos = [];
  let tablaEventos = [];
  
  // Etapa 1: identificar eventos
 
  for (let dia = 1; dia <= numDias; dia++){
    
    // Primera clase del día
    
    let eventoActual = {desc: horario[1][dia],
                        dia: horario[0][dia],
                        hInicio: horario[1][0],
                        hFin: horario[1][horario[0].length - 1],
                       };
    
    for (let franjaHoraria = 2; franjaHoraria <= numFranjas ; franjaHoraria++){
      
      if (eventoActual.desc == horario[franjaHoraria][dia]) eventoActual.hFin = horario[franjaHoraria][horario[0].length - 1];
      else {
 
        // Añadir evento
          
         eventos.push(eventoActual);
         eventoActual = {desc: horario[franjaHoraria][dia],
                         dia: horario[0][dia],
                         hInicio: horario[franjaHoraria][0],
                         hFin: horario[franjaHoraria][horario[0].length - 1],
                        };
      }
    }
    
    // Registrar el último evento del día
    
    eventos.push(eventoActual);
  }
  
  // Convertir a matriz
  
  eventos = eventos.map(evento => [evento.desc, evento.dia, evento.hInicio, evento.hFin]);
  
  // Etapa 2: Agrupar eventos que se repiten en varios días a la semana (utiliza fx ACOPLAR)
  
  if (agrupar) {
    
    eventos = ACOPLAR(eventos, false, separador, 1, 3, 4);
  
  }
  
  // Generar matriz resultado
    
  return eventos;
    
}