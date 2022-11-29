/**
 * Esta función acopla (combina) las filas de un intervalo de datos que corresponden a una misma entidad. Para ello, 
 * se debe indicar la columna (o columnas) *clave* que identifican los datos de cada entidad única. Los valores registrados
 * en el resto de columnas se agruparán, para cada una de ellas, utilizando como delimitador la secuencia de caracteres 
 * indicada. Se trata de una función que realiza una operación complementaria a DESACOPLAR(), aunque no perfectamente simétrica.
 * @param {A1:D10} intervalo Intervalo de datos.
 * @param {VERDADERO} encabezado Indica si el rango tiene una fila de encabezado con etiquetas para cada columna ([VERDADERO] | FALSO).
 * @param {", "} separador Secuencia de caracteres a emplear como separador de los valores múltiples. Opcional, si se omite se utiliza ", " (coma espacio).
 * @param {1} columna Número de orden, desde la izquierda, de la columna clave que identifica los datos de la fila como únicos.
 * @param {2} [más_columnas] Columnas clave adicionales, opcionales, que actúan como identificadores únicos, separadas por ";".
 *
 * @return Intervalo de datos desacoplados
 *
 * @customfunction
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer (@pfelipm)
 */ 
function ACOPLAR(intervalo, encabezado, separador, columna, ...masColumnas) {

  // Control de parámetros inicial
  
  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo.';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas.';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto.';
  let columnas = typeof columna != 'undefined' ? [columna, ...masColumnas].sort() : [...masColumnas].sort();
  if (columnas.length == 0) throw 'No se han indicado columnas clave.';
  if (columnas.some(col => typeof col != 'number' || col < 1)) throw 'Las columnas clave deben indicarse mediante números enteros';
  if (Math.max(...columnas) > intervalo[0].length) throw 'Al menos una columna clave está fuera del intervalo.';

  // Se construye un conjunto (set) para evitar automáticamente duplicados en columnas CLAVE
    
  const colSet = new Set();
  columnas.forEach(col => colSet.add(col - 1));
  
  // ...y en este conjunto se identifican las columnas susceptibles de contener valores que deben concatenarse
  
  const colNoClaveSet = new Set();
  for (let col = 0; col < intervalo[0].length; col++) {
  
    if (!colSet.has(col)) colNoClaveSet.add(col);
  
  }
  
  // Listos para comenzar
  
  if (encabezado) encabezado = intervalo.shift();
  
  const intervaloAcoplado = [];

  // 1ª pasada: recorremos el intervalo fila a fila para identificar entidades (concatenación de columnas clave) únicas
  
  const entidadesClave = new Set();
  intervalo.forEach(fila => {
    
    const clave = [];
    // ⚠️ A la hora de diferenciar dos entidades únicas (filas) usando una serie de columnas clave:
    //    a) No basta con concatenar los valores de las columnas clave como cadenas y simplemente compararlas. Ejemplo:
    //       clave fila 1 → col1 = 'pablo' col2 = 'felip'     >> Clave compuesta: 'pablofelip'
    //       clave fila 2 → col1 = 'pa'    col2 = 'blofelip'  >> Clave compuesta: 'pablofelip'
    //       ✖️ Misma clave compuesta, pero entidades diferentes
    //    b) No basta con con unir los valores de las columnas clave como cadenas utilizando un carácter delimitador. Ejemplo ('/'):
    //       clave fila 1 → col1 = 'pablo/' col2 = 'felip'    >> Clave compuesta: 'pablo//felip' 
    //       clave fila 2 → col1 = 'pablo'  col2 = '/felip'   >> Clave compuesta: 'pablo//felip'
    //       ✖️ Misma clave compuesta, pero entidades diferentes
    //    c) No es totalmente apropiado eliminar espacios antes y después de valores clave y unirlos usando un espacio delimitador (' '):
    //       clave fila 1 → col1 = ' pablo' col2 = 'felip'    >> Clave compuesta: 'pablo felip'
    //       clave fila 2 → col1 = 'pablo'  col2 = 'felip'    >> Clave compuesta: 'pablo felip'
    //       ✖️ Misma clave compuesta, pero entidades estrictamente diferentes (a menos que espacios anteriores y posteriores no importen)
    // 💡 En su lugar, se generan vectores con valores de columnas clave y se comparan sus versiones transformadas en cadenas JSON.
    for (const col of colSet) clave.push(String(fila[col])) 
    entidadesClave.add(JSON.stringify(clave));
                    
  });

  // 2ª pasada: obtener filas para cada clave única, combinar columnas no-clave y generar filas resultado

  for (const clave of entidadesClave) {

    const filasEntidad = intervalo.filter(fila => {
    
      const claveActual = [];
      for (const col of colSet) claveActual.push(String(fila[col]));
      return clave == JSON.stringify(claveActual);
     
    });

    // Acoplar todas las filas de cada entidad, concatenando valores en columnas no-clave con separador indicado

    const filaAcoplada = filasEntidad[0];  // Se toma la 1ª fila del grupo como base
    const noClaveSets = [];
    for (let col = 0; col < colNoClaveSet.size; col++) {noClaveSets.push(new Set())}; // Vector de sets para recoger valores múltiples   
    filasEntidad.forEach(fila => {
      
      let conjunto = 0;   
      for (const col of colNoClaveSet) {noClaveSets[conjunto++].add(String(fila[col]));}
                            
    });

    // Set >> Vector >> Cadena única con separador

    let conjunto = 0;
    for (const col of colNoClaveSet) {filaAcoplada[col] = [...noClaveSets[conjunto++]].join(separador);}

    intervaloAcoplado.push(filaAcoplada);
  
  }

  return encabezado.map ? [encabezado, ...intervaloAcoplado] : intervaloAcoplado;
  
}