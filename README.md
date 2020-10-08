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

*   `intervalo_horario`: Intervalo de datos que contiene el horario en el formato requerido:
    *   Fila 1: Días de la semana.
    *   Columna 1: Hora de inicio de la clase.
    *   Columna N: Hora de fin de la clase.
*   `agrupar`: Indica si se deben tratar de agrupar los eventos que se repiten a lo largo de la semana en el mismo horario (`VERDADERO` o `FALSO`). Si se omite se asume `FALSO`.
