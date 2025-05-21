![Banner horarios-a-calendar](https://user-images.githubusercontent.com/12829262/95454414-b43dd080-096c-11eb-99d1-854f66187e81.png)

![Creado con - Google Apps Script](https://img.shields.io/static/v1?label=Creado+con&message=Google+Apps+Script&color=blue&logo=GAS)

# Horarios-a-Calendar

**Horarios a Calendar** (**HaC** en adelante) es un tinglado montado sobre **hojas de cálculo de Google** y **Apps Script** para facilitar la creación y gestión de eventos recurrentes en Google Calendar para representar un conjunto de sesiones de clase definidas mediante una tabla horaria semanal. Se ha diseñado con el objetivo de facilitar la generación y mantenimiento de calendarios docentes y de ocupación de aulas en el contexto de las actividades formativas de un centro de formación, aunque podría resultar también de utilidad en otros entornos.

Es posible que hayas llegado aquí desde este artículo que explica los cómos y porqués de HaC, pero si no es el caso probablemente sea una buena idea que le eches un vistazo para entender qué problema pretende resolver HaC antes de seguir:

:point\_right: [Eventos recurrentes en Google Calendar para tus horarios de clase con Apps Script y HaC](https://pablofelip.online/horarios-a-calendar/)

El proceso a desarrollar se puede descomponer en varias partes para:

1.  La persona responsable de la generación de los horarios genera una copia de la plantilla de horario semanal facilitada e introduce en ella la información básica del horario: nombre del grupo, fecha de inicio y fin de las clases.
2.  A continuación, introduce el nombre de cada clase (asignatura, materia o módulo profesional) en las celdas de la tabla horaria de la plantilla, tabla que dispone las sesiones de clase en días de la semana (columnas) y franjas horarias diarias (filas).
3.  Se genera automáticamente una lista de sesiones de clase que las enumera y recoge su información característica (día de la semana, hora de inicio y fin) extraída en tiempo real a partir de la tabla horaria semanal, es decir, al mismo tiempo que se introducen en ella las clases. La automatización permite agrupar las sesiones de clase que se repiten en el mismo horario a lo largo de la semana, de manera opcional.
4.  La persona responsable de la generación de los horarios establece los instructores y espacios (aulas) utilizados en cada sesión.
5.  Desde una hoja de control se selecciona la hoja horaria cuyos clases desean generarse como eventos recurrentes en Google Calendar. También es posible tanto añadir sesiones de clase no presentes en la hoja del horario seleccionado, como prescindir totalmente de una horario ya existente e introducir manualmente la información de todas las sesiones de clase en esta hoja de control (aunque este no es el modo recomendado de funcionamiento).
6.  Esta hoja de control permite seleccionar las clases para las que se desea generar eventos en Calendar. También es posible actualizarlos, tras realizar modificaciones en su definición, o eliminarlos, si es que ya han sido creados con anterioridad.

![diagrama funcional HaC](https://docs.google.com/drawings/d/1UsJVHvxpDVvWrwY0m1xZ6re7UAttSVxxSoWjrnKuGZY/export/png)

## Funcionamiento

**Horarios a Calendar** opera mediante una interacción coordinada entre las Hojas de Cálculo de Google, el código Apps Script y Google Calendar. A continuación, se detalla esta interacción y el papel de los principales archivos `.gs` involucrados:

**Interacción entre Componentes:**

1.  **Hojas de Cálculo de Google:** Sirven como la interfaz principal para el usuario. Aquí se introducen los datos del horario (nombre del grupo, fechas, asignaturas, horas, etc.) en una plantilla estructurada. También contienen hojas de control para gestionar la creación y actualización de eventos.
2.  **Apps Script:** Es el motor que automatiza el proceso. Los scripts leen los datos de las hojas de cálculo, los procesan y luego utilizan la API de Google Calendar para crear, actualizar o eliminar eventos.
3.  **Google Calendar:** Es el destino final donde se visualizan los horarios como eventos. Los eventos creados por HaC incluyen toda la información relevante, como el nombre de la asignatura, el profesor, el aula y las horas.

**Roles de los Archivos `.gs`:**

El código Apps Script se organiza en varios archivos `.gs`, cada uno con responsabilidades específicas:

*   **`Eventos.gs`**: Contiene la lógica principal para interactuar con Google Calendar. Se encarga de crear, buscar, actualizar y eliminar eventos basándose en la información procesada de las hojas de cálculo. Gestiona la recurrencia de los eventos y la asignación de información detallada a cada uno.
*   **`Instructores.gs`**: Gestiona la información relativa a los instructores o profesores. Permite obtener y asignar instructores a las diferentes sesiones de clase.
*   **`Salas.gs`**: Similar a `Instructores.gs`, pero enfocado en las aulas o espacios físicos donde se imparten las clases. Permite obtener y asignar salas a las sesiones.
*   **`Parametrización.gs`**: Almacena configuraciones y parámetros utilizados por los scripts, como pueden ser los nombres de las hojas de cálculo específicas, rangos de celdas clave, o valores por defecto. Facilita la adaptación de HaC a diferentes hojas de cálculo sin necesidad de modificar el código principal.
*   **`Auxiliar.gs`**: Reúne funciones de utilidad general que son utilizadas por otros módulos. Estas funciones pueden incluir tareas como el formateo de fechas, la manipulación de cadenas de texto, o la gestión de errores comunes.
*   **`fx Acoplar.gs`**: Define la función personalizada `=ACOPLAR()`. Esta función se utiliza dentro de las hojas de cálculo para combinar o concatenar cadenas de texto de manera más flexible que las funciones nativas, útil para generar descripciones de eventos o nombres compuestos.
*   **`fx Extraer eventos.gs`**: Define la función personalizada `=EXTRAEREVENTOS()`, descrita en detalle más adelante en la sección "Fx personalizada EXTRAEREVENTOS()". Su propósito es analizar la tabla horaria introducida por el usuario y transformarla en una lista estructurada de sesiones de clase, identificando el día, hora de inicio, hora de fin y descripción de cada una.

**Flujo de Datos:**

1.  **Entrada de Datos:** El usuario introduce la información del horario en la hoja de cálculo designada (la plantilla de horario). Esto incluye los nombres de las asignaturas en las celdas correspondientes a los días y franjas horarias.
2.  **Procesamiento con `EXTRAEREVENTOS()`:** La función `=EXTRAEREVENTOS()` (definida en `fx Extraer eventos.gs`) lee la tabla horaria y genera una lista detallada de todas las sesiones de clase. Esta lista incluye el día de la semana, la hora de inicio y fin, y el nombre de la asignatura. Opcionalmente, puede agrupar sesiones que se repiten.
3.  **Enriquecimiento de Datos:** En la hoja de control, el usuario puede asignar instructores (gestionados por `Instructores.gs`) y salas (gestionados por `Salas.gs`) a estas sesiones. También se pueden añadir sesiones manualmente o modificar las existentes.
4.  **Generación de Eventos:** Desde la hoja de control, el usuario activa la creación o actualización de eventos.
5.  **Interacción con Calendar:** El script `Eventos.gs` toma la lista de sesiones procesada y enriquecida. Para cada sesión, crea un evento recurrente en Google Calendar con toda la información asociada (asignatura, instructor, sala, horario, fechas de inicio y fin del periodo de clases). Si los eventos ya existen, los actualiza o los elimina según la acción solicitada por el usuario.
6.  **Retroalimentación:** Se puede proporcionar información al usuario sobre el estado de la creación de eventos (por ejemplo, IDs de los eventos creados) directamente en la hoja de cálculo.

Este flujo permite una gestión centralizada y eficiente de los horarios, transformando una simple tabla en una serie de eventos de calendario listos para ser consultados por profesores y alumnos.

## Arquitectura Interna

La arquitectura de **Horarios a Calendar** se fundamenta en la sinergia de varios componentes y servicios de Google Workspace, configurados y orquestados para lograr la automatización deseada.

**Componentes Principales:**

1.  **Hojas de Cálculo de Google (Google Sheets):**
    *   **Fuente de Datos:** Almacenan toda la información de entrada, como los detalles del horario, las asignaturas, los instructores, las aulas, y las fechas de inicio y fin del ciclo.
    *   **Interfaz de Usuario (UI):** Proporcionan la interfaz principal para que los usuarios introduzcan y gestionen los datos del horario, así como para disparar las acciones de creación, actualización o eliminación de eventos a través de menús personalizados o botones que ejecutan funciones de Apps Script.

2.  **Google Apps Script:**
    *   **Lógica de Backend:** Actúa como el motor de procesamiento en segundo plano. Contiene todo el código JavaScript (en archivos `.gs`) que implementa la lógica de negocio: leer datos de las Hojas de Cálculo, procesarlos, validar la información, interactuar con Google Calendar y gestionar los recursos del dominio.
    *   **Orquestación:** Coordina las interacciones entre las Hojas de Cálculo y Google Calendar. Define funciones personalizadas para las hojas (`=EXTRAEREVENTOS()`, `=ACOPLAR()`) y funciones que se ejecutan desde la UI de la hoja o de forma programada.

3.  **Google Calendar:**
    *   **Plataforma de Salida:** Es el destino final donde se crean, actualizan o eliminan los eventos del horario. Los usuarios (profesores, alumnos) consultan sus horarios a través de esta plataforma.
    *   **Gestión de Eventos:** Los eventos generados por HaC son eventos de Google Calendar estándar, pudiendo incluir detalles como título, descripción, horas de inicio y fin, recurrencia, invitados (instructores) y ubicaciones (aulas/recursos de Calendar).

**Configuración del Proyecto (`appsscript.json`):**

El archivo de manifiesto `appsscript.json` juega un papel crucial en la configuración del proyecto de Apps Script:

*   **`timeZone`**: Especifica la zona horaria en la que se ejecutarán los scripts y se interpretarán las fechas y horas. Es fundamental para la correcta programación de los eventos.
*   **`dependencies`**: Declara las dependencias de servicios avanzados de Google. En HaC, es especialmente importante para:
    *   **`AdminDirectory` (Directory API Service):** Permite a los scripts acceder a recursos del dominio de Google Workspace, como las salas de Calendar (recursos de calendario). Esto es esencial para buscar y asignar aulas a los eventos. Se requiere que el usuario que ejecuta el script tenga los permisos adecuados a nivel de dominio para usar esta API.
*   **`runtimeVersion`**: Define la versión del motor de ejecución de Apps Script (por ejemplo, `V8`). Utilizar el motor V8 es recomendable por sus mejoras en rendimiento y compatibilidad con sintaxis moderna de JavaScript.
*   **`oauthScopes`**: Define los permisos de autorización que el script necesita para operar. Estos scopes se presentan al usuario la primera vez que ejecuta una función que los requiere. Incluiría scopes como:
    *   `https://www.googleapis.com/auth/script.container.ui` (para crear menús o diálogos en Sheets)
    *   `https://www.googleapis.com/auth/spreadsheets` (para leer y escribir en las Hojas de Cálculo)
    *   `https://www.googleapis.com/auth/calendar` (para crear y gestionar eventos en Calendar)
    *   `https://www.googleapis.com/auth/admin.directory.resource.calendar.readonly` (para leer los recursos de calendario del dominio, si se usa `AdminDirectory` para las salas).

**Servicios de Google Apps Script Utilizados:**

El código de HaC hace uso extensivo de los servicios integrados de Apps Script para interactuar con los diferentes componentes de Google Workspace:

*   **`SpreadsheetApp`**: Este servicio proporciona los métodos para acceder y modificar los datos en las Hojas de Cálculo de Google. Se utiliza para:
    *   Obtener la hoja de cálculo activa o una específica por su ID o nombre.
    *   Leer datos de rangos de celdas (e.g., la tabla del horario, la lista de instructores).
    *   Escribir datos en celdas (e.g., IDs de eventos creados, mensajes de estado).
    *   Acceder a la interfaz de usuario de la hoja para crear menús personalizados.

*   **`CalendarApp`**: Este servicio permite la interacción con Google Calendar. Es fundamental para:
    *   Acceder a un calendario específico (por su ID o el calendario por defecto del usuario).
    *   Crear nuevos eventos, especificando todos sus atributos (título, descripción, horas, recurrencia, invitados, ubicación).
    *   Buscar eventos existentes (para actualizarlos o eliminarlos).
    *   Modificar eventos existentes.
    *   Eliminar eventos.

*   **`AdminDirectory` (a través del servicio avanzado `AdminDirectory`):** Cuando se habilitan los servicios avanzados, este objeto permite interactuar con la API de Directory de Google Workspace. En HaC, su uso principal es:
    *   Listar y obtener detalles de los recursos de calendario del dominio (salas de reuniones, aulas) para poder asignarlos como ubicación de los eventos. Esto asegura que se utilizan los identificadores correctos de los recursos y permite verificar su disponibilidad si fuera necesario (aunque HaC se centra en la creación, no en la gestión compleja de conflictos de reserva).

Esta arquitectura modular y basada en los servicios de Google permite que **Horarios a Calendar** sea una herramienta robusta y extensible para la gestión de horarios académicos.

# Fx personalizada EXTRAEREVENTOS()

```
=EXTRAEREVENTOS( intervalo_horario ; [agrupar] ; [separador] ) 
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
