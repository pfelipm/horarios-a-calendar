![Banner horarios-a-calendar](https://user-images.githubusercontent.com/12829262/95454414-b43dd080-096c-11eb-99d1-854f66187e81.png)

![Creado con - Google Apps Script](https://img.shields.io/static/v1?label=Creado+con&message=Google+Apps+Script&color=blue&logo=GAS)

# Horarios-a-Calendar

**Horarios a Calendar** (**HaC** en adelante) es un tinglado montado sobre **hojas de cálculo de Google** y **Apps Script** para facilitar la creación y gestión de eventos recurrentes en Google Calendar para representar un conjunto de sesiones de clase definidas mediante un tabla horaria semanal. Se ha diseñado con el objetivo de facilitar la generación y mantenimiento de calendarios docentes y de ocupación de aulas en el contexto de las actividades formativas de un centro de formación, aunque podría resultar también de utilidad en otros entornos.

Es posible que hayas llegado aquí desde este artículo que explica los cómos y porqués de HaC, pero si no es el caso probablemente sea una buena idea que le eches un vistazo para entender qué problema pretende resolver HaC antes de seguir:

:point\_right: [Eventos recurrentes en Google Calendar para tus horarios de clase con Apps Script y HaC](https://pablofelip.online/horarios-a-calendar/)

El proceso a desarrollar se puede descomponer en varias partes para:

1.  La persona responsable de la generación de los horarios genera una copia de la plantilla de horario semanal facilitada e introduce en ella la información básica del horario: nombre del grupo, fecha de inicio y fin de las clases.
2.  A continuación, introduce el nombre de cada clase (asignatura, materia o módulo profesional) en las celdas de la tabla horaria de la plantilla, tabla que dispone las sesiones de clase en días de la semana (columnas) y franjas horarias diarias (filas).
3.  Se genera automáticamente una lista de sesiones de clase que las enumera y recoge su información característica (día de la semana, hora de inicio y fin) extraída en tiempo real a partir de la tabla horaria semanal, es decir, al mismo tiempo que se introducen en ella las clases. La automatización permite agrupar las sesiones de clase que se repiten en el mismo horario a lo largo de la semana, de manera opcional.
4.  La persona responsable de la generación de los horarios establece los instructores y espacios (aulas) utilizados en cada sesión.
5.  Desde una hoja de control se selecciona la hoja horaria cuyos clases desean generarse como eventos recurrentes en Google Calendar. También es posible tanto añadir sesiones de clase no presentes en la hoja del horario seleccionado, como prescindir totalmente de una horario ya existente e introducir manualmente la información de todas las sesiones de clase en esta hoja de control (aunque este no es el modo recomendado de funcionamiento).
6.  Esta hoja de control permite seleccionar las clases para las que se desea generar eventos en Calendar. También es posible actualizarlos, tras realizar modificaciones en su definición, o eliminarlos, si es que ya han sido creados con anterioridad.

# Fx personalizada EXTRAEREVENTOS()

```
=EXTRAEREVENTOS( intervalo_horario ; [agrupar]; [separador] ) 
```

*   `intervalo_horario`: Intervalo de datos que contiene el horario en el formato requerido: fila 1 (etiqueta de cada sesión, típicamente los **días de la semana**), columna inicial (**hora de inicio del evento**), columna final (**hora de fin del evento**), celdas (**descripción del evento**). Se requiere que cada franja horaria (fila) disponga de indicación de la hora de inicio y fin de cada sesión y no se combinen las celdas en el interior de la tabla.

<table><tbody><tr><td>&nbsp;</td><td><strong>L</strong></td><td><strong>M</strong></td><td><strong>X</strong></td><td><strong>J</strong></td><td><strong>V</strong></td><td>&nbsp;</td></tr><tr><td><strong>H inicio</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>H Fin</strong></td></tr><tr><td><strong>...</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>...</strong></td></tr><tr><td><strong>H inicio</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>H Fin</strong></td></tr></tbody></table>

*   `agrupar`: Indica si se deben tratar de agrupar los eventos que se repiten a lo largo de la semana en el mismo horario (`VERDADERO` o `FALSO`). Si se omite se asume `FALSO`.
*   `separador`: Secuencia de caracteres utilizada para separar la etiqueta de cada sesión en el caso de que se haya solicitado la agrupación de eventos. Si se omite se concatenan las etiquetas sin más.

Ejemplo:

```
=EXTRAEREVENTOS( A1:G13 ; VERDADERO ; "-" ) 
```

![extraereventodemo](https://user-images.githubusercontent.com/12829262/95462129-64183b80-0977-11eb-8a67-1eb50234893a.png)

# **Modo de uso**

Para usar `EXTRAEREVENTOS()` en tus proyectos, abre el editor GAS de tu hoja de cálculo (`Herramientas` **⇒** `Editor de secuencias de comandos`), pega el código que encontrarás dentro de los archivos `fx ACOPLAR.gs` y `fx EXTRAEREVENTOS.gs` de este repositorio y guarda los cambios. Debes asegurarte de que se esté utilizando el nuevo motor GAS JavaScript V8 (`Ejecutar` **⇒** `Habilitar ... V8`).
