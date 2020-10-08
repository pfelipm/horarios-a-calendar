# Horarios-a-Calendar

Un conjunto de scripts GAS para automatizar la creación  y gestión de eventos en Google Calendar para representar las clases a partir de una definición de horaria en formato tabla en una hoja de cálculo de Google.

**⚒️ En construcción** ⚒️

# Fases

1.  ✔️ Definir horario semanal en filas y columnas de una tabla.
2.  ✔️ Extraer los eventos utilizando una función de hojas de cálculo personalizada (`EXTRAEREVENTOS`), permitiendo la agrupación de instancias de la misma actividad que se repiten en el mismo horario a lo largo de la semana.
3.  Integrar eventos en tabla de gestión que posibilite su administración y gestión automatizada (no generar, eliminar, modificar, etc.).
4.  Generar eventos en calendar a partir de la tabla anterior mediante script activado desde menú de la hoja de cálculo.

# Fx personalizada EXTRAEREVENTOS()

```
=EXTRAEREVENTOS( intervalo_horario ; [agrupar] )
```

*   `intervalo_horario`: Intervalo de datos que contiene el horario en el formato requerido: Fila 1 (días de la semana), columna inicial (hora de inicio del evento), columna final (hora de fin del evento), celdas (descripción del evento):

<table><tbody><tr><td>&nbsp;</td><td>L</td><td>M</td><td>X</td><td>J</td><td>V</td><td>&nbsp;</td></tr><tr><td>H inicio</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>H Fin</td></tr><tr><td>...</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>...</td></tr><tr><td>H inicio</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>Descripción</td><td>H Fin</td></tr></tbody></table>

*   `agrupar`: Indica si se deben tratar de agrupar los eventos que se repiten a lo largo de la semana en el mismo horario (`VERDADERO` o `FALSO`). Si se omite se asume `FALSO`.
