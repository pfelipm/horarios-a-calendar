![Banner horarios-a-calendar](https://user-images.githubusercontent.com/12829262/95454414-b43dd080-096c-11eb-99d1-854f66187e81.png)

# Horarios-a-Calendar

Un tinglado Google Apps Script que automatiza la creación y gestión de eventos en Google Calendar para representar los eventos (clases en un contexto de enseñanza) a partir de una definición horaria en formato tabla en una hoja de cálculo de Google.

**⚒️ En construcción** ⚒️

# Fases

1.  ✅ Definir horario semanal en filas y columnas de una tabla.
2.  ✅ Extraer los eventos utilizando la función de hojas de cálculo personalizada `EXTRAEREVENTOS()`, permitiendo la agrupación de instancias de la misma actividad que se repiten en el mismo horario a lo largo de la semana. Utiliza la función `ACOPLAR()` para agrupar eventos en horarios semanales coincidentes, tal y como se describe en el [repositorio desacoplar-acoplar](https://github.com/pfelipm/desacoplar-acoplar).
3.  Integrar eventos en una tabla de gestión que posibilite su administración y gestión automatizada (no generar, eliminar, modificar, etc.).
4.  Generar eventos en calendar a partir de la tabla anterior mediante script activado desde menú de la hoja de cálculo.

# Fx personalizada EXTRAEREVENTOS()

```
=EXTRAEREVENTOS( intervalo_horario ; [agrupar]; [separador] )
```

*   `intervalo_horario`: Intervalo de datos que contiene el horario en el formato requerido: fila 1 (etiqueta de cada sesión, típicamente los **días de la semana**), columna inicial (**hora de inicio del evento**), columna final (**hora de fin del evento**), celdas (**descripción del evento**):

<table><tbody><tr><td>&nbsp;</td><td><strong>L</strong></td><td><strong>M</strong></td><td><strong>X</strong></td><td><strong>J</strong></td><td><strong>V</strong></td><td>&nbsp;</td></tr><tr><td><strong>H inicio</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>H Fin</strong></td></tr><tr><td><strong>...</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>...</strong></td></tr><tr><td><strong>H inicio</strong></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><i>Descripción</i></td><td><strong>H Fin</strong></td></tr></tbody></table>

*   `agrupar`: Indica si se deben tratar de agrupar los eventos que se repitan a lo largo de la semana en el mismo horario (`VERDADERO` o `FALSO`). Si se omite se asume `FALSO`.
*   `separador`: Secuencia de caracteres utilizada para separar la etiqueta de cada sesión en el caso de que se haya solicitado la agrupación de eventos. Si se omite se concatenan las etiquetas sin más.

Ejemplo:

```
=EXTRAEREVENTOS( A1:G13 ; VERDADERO ; "-" )
```

![extraereventodemo](https://user-images.githubusercontent.com/12829262/95462129-64183b80-0977-11eb-8a67-1eb50234893a.png)
