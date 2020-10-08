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

<table><tbody><tr><td>&nbsp;</td><td>L</td><td>M</td><td>X</td><td>J</td><td>V</td><td>&nbsp;</td></tr><tr><td>8:00</td><td>IN/AA</td><td>IN/AA</td><td>PM/FZ</td><td>PM/FZ</td><td>IN/AA</td><td>8:30</td></tr><tr><td>8:30</td><td>IN/AA</td><td>IN/AA</td><td>PM/FZ</td><td>PM/FZ</td><td>IN/AA</td><td>9:00</td></tr><tr><td>9:00</td><td>IN/AA</td><td>IN/AA</td><td>PM/FZ</td><td>PM/FZ</td><td>PM/FZ</td><td>9:30</td></tr><tr><td>9:30</td><td>IN/AA</td><td>IN/AA</td><td>PM/FZ</td><td>PM/FZ</td><td>PM/FZ</td><td>10:00</td></tr><tr><td>10:00</td><td>DESCANSO</td><td>DESCANSO</td><td>DESCANSO</td><td>DESCANSO</td><td>IC/FZ</td><td>10:30</td></tr><tr><td>10:30</td><td>GEE/MM</td><td>GEE/MM</td><td>IC/FZ</td><td>IC/FZ</td><td>IC/FZ</td><td>11:00</td></tr><tr><td>11:00</td><td>GEE/MM</td><td>GEE/MM</td><td>IC/FZ</td><td>IC/FZ</td><td>MD/PF</td><td>11:30</td></tr><tr><td>11:30</td><td>GEE/MM</td><td>GEE/MM</td><td>IC/FZ</td><td>IC/FZ</td><td>MD/PF</td><td>12:00</td></tr><tr><td>12:00</td><td>GEE/MM</td><td>GEE/MM</td><td>MD/PF</td><td>MD/PF</td><td>GEE/MM</td><td>12:30</td></tr><tr><td>12:30</td><td>GEE/MM</td><td>GEE/MM</td><td>MD/PF</td><td>MD/PF</td><td>GEE/MM</td><td>13:00</td></tr><tr><td>13:00</td><td>GEE/MM</td><td>FOL/MM</td><td>MD/PF</td><td>MD/PF</td><td>FOL/MM</td><td>13:30</td></tr><tr><td>13:30</td><td>GEE/MM</td><td>FOL/MM</td><td>MD/PF</td><td>MD/PF</td><td>FOL/MM</td><td>14:00</td></tr></tbody></table>

*   `agrupar`: Indica si se deben tratar de agrupar los eventos que se repiten a lo largo de la semana en el mismo horario (`VERDADERO` o `FALSO`). Si se omite se asume `FALSO`.
