# Form_FormCalidadRiesgoAceptadoRetiradoVisado

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadRiesgoAceptadoRetiradoVisado.cls` (349 líneas) y `src/forms/FormCalidadRiesgoAceptadoRetiradoVisado.form.txt` · **PRD:** [README](../README.md)

Ventana de visado de calidad. Resuelve una sola operación sobre un riesgo: **aceptarlo o retirarlo**, y esa resolución puede ser «poner», «quitar lo puesto», «rechazar» o «quitar el rechazo». No mantiene los datos: los dibuja y devuelve el resultado mediante ocho eventos públicos. Las citas son `Form_FormCalidadRiesgoAceptadoRetiradoVisado.cls:línea` salvo indicación.

## Entrada y salida

El único punto de entrada en `src` es el formulario de tareas de riesgos aceptados y retirados, que lo declara con `WithEvents` y lo abre como modal de datos (`Form_FormCalidadTareaRiesgosAceptadosRetirados.cls:17, 27-28, 49-51`).

| Parámetro | Cómo | Qué significa | Evidencia |
|---|---|---|---|
| `OpenArgs` | `m_EsMitigacion = Nz(Me.openArgs(), "")` | `«Sí»` = el riesgo está en proceso de aceptación; cualquier otro valor, `«No»`, significa retirada | `:177` |

El riesgo sobre el que se opera **no viaja como parámetro**: se toma de la global `m_ObjRiesgoActivo`, y solo se comprueba que no esté vacía (`:182-185`). Es el mismo objeto que leen los formularios hermanos del formulario de tareas.

**Salidas:** ocho eventos públicos, cuatro de aceptación y cuatro de retirada (`:16-19`, `:22-25`). De los ocho, el formulario contenedor solo atiende cuatro: `AceptacionAprobada`, `AceptacionAprobadaQuitado`, `AceptacionRechazada` y `AceptacionRechazadaQuitada` (`Form_FormCalidadTareaRiesgosAceptadosRetirados.cls:160, 315, 341, 367`). Tras resolver, el formulario se cierra solo (`:81`, `:149`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Valida el parámetro, comprueba el riesgo activo, calcula el estado y pinta el formulario | `:168-210` |
| `ComandoAceptar` | botón | `ComandoAceptar_Click` | Resuelve la aceptación o la retirada, según dos cosas: el parámetro y **su propio rótulo** | `:30-95`, `.form.txt:1247-1248` |
| `ComandoRechazar` | botón | `ComandoRechazar_Click` | Resuelve el rechazo y recalcula las tareas de calidad | `:97-161`, `.form.txt:1283-1284` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra sin guardar, con `On Error Resume Next` | `:163-166`, `.form.txt:1329-1330` |
| `Justificacion` | cuadro de texto | — | Muestra la justificación escrita por el técnico | `:244`, `:292`, `.form.txt:1003-1007` |
| `FechaJustificacion` | cuadro de fecha | — | Muestra la fecha de justificación | `:247`, `:295`, `.form.txt:1035-1039` |
| `FechaAprobacion` | cuadro de fecha | — | Muestra la fecha en que calidad aprobó | `:250`, `:298`, `.form.txt:1067-1071` |
| `FechaRechazo` | cuadro de fecha | — | Muestra la fecha en que calidad rechazó | `:253`, `:301`, `.form.txt:1099-1103` |
| `lblTitulo1` | etiqueta | — | Rótulo fijo «VISADO POR CALIDAD» | `.form.txt:902-903` |
| `Comando39` | botón | — | Tiene `OnClick` declarado pero **ningún manejador** en el código | `.form.txt:923-926` |

## Flujos

### Apertura (`Form_Load`, `:168-210`)

1. Reloj de arena y `m_Error` a vacío (`:173-175`).
2. `AjustarTamaño Me` (`:176`).
3. Lee el parámetro y exige que sea `«Sí»` o `«No»`; si no, avisa «El formulario se ha abierto con datos insuficientes» y **cierra** (`:177-181`).
4. Exige que `m_ObjRiesgoActivo` exista (`:182-185`) y calcula si el riesgo está activo (`:186-190`).
5. `EstablecerDatos` (`:191`) y cierre si devolvió error (`:192-194`).
6. En error, el manejador avisa y cierra (`:201-209`).

### Pintado del estado (`EstablecerDatos`, `:215-347`)

1. Toma el estado del riesgo (`:225`).
2. Pregunta si está en retirada (`:227`, `Funciones Generales.bas:6620`); si no, pregunta si está en aceptación (`:232`, `Funciones Generales.bas:6595`); si tampoco, avisa «El riesgo no está ni en estado de aceptación ni de retirada» (`:235-239`).
3. **Rama de aceptación** (`:240-289`): vuelca justificación y fechas (`:243-254`) y, según el estado, fija rótulos y habilitación de los dos botones.
4. **Rama de retirada** (`:290-340`): lo mismo con las propiedades de retirada.

La habilitación responde a la misma condición en los ocho casos: `EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True` (`:264, 275, 284, 312, 323, 332`). Es decir: **solo quien no es técnico y solo sobre un riesgo activo**.

### Normal: aceptar o retirar (`:30-95`)

1. Reloj de arena (`:36-38`).
2. Elegir la operación con la doble condición de `:37-59`: mitigación o no, y rótulo del botón igual a `«Aceptar»` o distinto.
3. Si el objeto devuelve error, `Err.Raise 1000` (`:48-50`, `:57-59`).
4. Emitir el evento correspondiente (`:60-77`).
5. Quitar el reloj y cerrarse (`:78-81`).

### Alternativo: rechazar (`:97-161`)

Repite la misma doble condición con `«Rechazar»` (`:101-125`) y, además, **recalcula las tareas de calidad** con `EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error` (`:127-130`; `Funciones Generales.bas:2054`). Aprobar no lo hace.

### De error

Los cuatro manejadores repiten el mismo patrón: si el número de error **no** es 1000, compone el mensaje con número y detalle, **envía correo al administrador** y muestra un diálogo crítico; si es 1000 exacto, muestra solo una advertencia. En ninguno se cancela la operación ya escrita ni se avisa de un estado intermedio (`:85-94`, `:152-161`, `:200-209`).

## Reglas, validaciones y cálculos

- **El rótulo del botón decide la operación.** Cuatro métodos distintos —`.AceptacionAprobar`, `.AceptacionAprobarQuitar`, `.RetiroAprobar` y `.RetiroAprobacionQuitar` (`:41, 44, 52, 55`) — se eligen comparando `Me.ComandoAceptar.Caption = "Aceptar"` (`:37, 42, 60, 66`). Lo mismo en el rechazo (`:101, 107, 128, 134`). El rótulo no es decoración: es parte de la decisión de negocio.
- **El estado del riesgo fija los rótulos** (`:255-289`, `:303-336`): los estados «sin justificar» dejan ambos botones deshabilitados; «sin visar» los habilita; los estados decididos convierten el botón en «Quitar Aceptar» o «Quitar Rechazar».
- **La fecha de resolución es la del sistema:** los cuatro métodos de decisión reciben `Date` (`:41, 52, 105, 116`). El usuario no puede fecharla.
- **`EsTecnico` es una global pública de tipo `EnumSiNo`** (`Variables Globales.bas:109`) que se inicializa a `Empty` (`:189`), y se compara con `<>` contra `EnumSiNo.Sí`.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo que se visa | Global `m_ObjRiesgoActivo` | `:182-185`, `:240-291` |
| Estado y clasificación | `EstadoEnum`, `RiesgoEnRetirada`, `RiesgoEnAceptacion` | `:225-239` |
| Justificación y fechas | Propiedades de aceptación o de retirada del riesgo | `:243-254`, `:290-301` |
| Decisión escrita | `AceptacionAprobar`, `AceptacionRechazo`, `RetiroAprobar`, `RetiroRechazo` y sus variantes «quitar» | `:41-56`, `:105-120` |
| Tareas de calidad | `EstablecerTareasCalidad`, solo en el rechazo | `:127` |

## Coexistencia

Pendiente de `../../forms.md`. Lo que ata esta ventana al Access es el contrato de eventos: el resultado **no** se devuelve como dato ni por código de salida, sino por ocho eventos que el contenedor implementa uno a uno. Y el dato de entrada no viaja con la petición, sino en una global compartida por la sesión.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-82.** La lógica de negocio se despacha por el **rótulo del botón** (`Me.ComandoAceptar.Caption = "Aceptar"`, `:37, 42`; `Me.ComandoRechazar.Caption = "Rechazar"`, `:101, 107`). Cualquier traducción, cambio de texto o reinicio del rótulo por diseño ejecuta la operación contraria sin avisar. Conviene comprobar si algún otro formulario toca esos rótulos.
- **H-GR-83.** `EstablecerDatos` fija los rótulos «Aceptar» y «Rechazar» solo para ocho estados concretos (`:255-336`) y la cadena `If/ElseIf` **no tiene rama final**. Si el riesgo llegara en un estado no cubierto y las dos clasificaciones lo hubieran dejado pasar, los botones quedarían con los rótulos del diseño y sin la lógica que los acompaña.
- **H-GR-84.** Los cuatro campos de solo lectura (`Justificacion`, `FechaJustificacion`, `FechaAprobacion`, `FechaRechazo`) están ordenados para recibir el foco (`TabIndex =1, 2, 3`, `.form.txt:1032, 1064, 1096`) y el diseño **no declara `Locked` en ninguno** (`.form.txt:1003-1103`), mientras que el código solo los lee y nunca los escribe (`:243-254`, `:290-301`). Si en ejecución son editables, lo que el usuario escriba se pierde al pulsar Aceptar. Hay que verificar en el binario.
- **H-GR-85.** Asimetría entre las dos decisiones: el rechazo recalcula las tareas de calidad (`:127`) y la aceptación no. Tras aceptar, la lista de tareas del contenedor puede quedar desfasada salvo que el evento emitido lo compense; conviene comprobarlo en `Form_FormCalidadTareaRiesgosAceptadosRetirados.cls:160`.
- **H-GR-86.** Cuatro de los ocho eventos emitidos no tienen destinatario: `RetiroAprobado`, `RetiroRechazado`, `RetiroAprobadoQuitado` y `RetiroRechazadoQuitada` (`:22-25`, emitidos en `:72, 74, 140, 142`) no tienen manejador en el contenedor (`Form_FormCalidadTareaRiesgosAceptadosRetirados.cls` solo declara los cuatro de aceptación). La decisión se escribe en el modelo, pero **nadie se entera**.
- **H-GR-87.** Dos de los eventos llevan sufijo femenino y los otros dos no: `AceptacionRechazadaQuitada` (`:19`) frente a `AceptacionRechazada` (`:17`), `RetiroRechazadoQuitada` (`:25`) frente a `RetiroRechazado` (`:23`). Es la misma convención en dos convenciones; el contenedor reproduce el desorden.
- **H-GR-88.** `EsTecnico` se inicializa a `Empty` (`Variables Globales.bas:189`) y la condición de habilitación es `EsTecnico <> EnumSiNo.Sí` (`:264` y cinco veces más). Antes de que esa global se resuelva, `Empty <> 1` es cierto y los botones se muestran deshabilitados; hay que verificar el orden real de inicialización en el arranque.
- **H-GR-89.** Todos los `Err.Raise 1000` del fichero (`:49, 58, 114, 123, 129, 180, 184, 193, 229, 234, 238`) van sin `vbObjectError` ni `Source`. Al no ser errores internos de Access, conviene comprobar si el manejador propio se ejecuta o si el error sube al formulario contenedor.
- **H-GR-90.** Los seis diálogos usan `pregunta` (`Variables Globales.bas:69`), una **global pública compartida por toda la aplicación**, como simple contenedor del valor devuelto por `MsgBox` (`:89, 91, 157, 159, 205, 207`). Dos ventanas abiertas a la vez se pisan el valor.
- **H-GR-91.** `m_RiesgoEnAceptacion` solo se calcula dentro de la rama «no está en retirada» (`:227-236`). Cuando el riesgo sí está en retirada, la variable conserva el valor por defecto de Access y la decisión entre ramas se resuelve por comparación contra una enumeración que no tiene ese valor (`Constantes.bas:114-117`, `Sí = 1`, `No = 2`). Funciona, pero por una razón que el código no dice.
- **H-GR-92.** `Comando39` declara `OnClick ="[Event Procedure]"` (`.form.txt:925`) y **no existe `Comando39_Click` en el código**: el manejador fue borrado y el botón se quedó enlazado. Su rótulo es «Salir» (`.form.txt:924`) y su texto emergente es «Ayuda» (`.form.txt:926`), de modo que hay además dos botones de salida en el formulario (`.form.txt:924` y `:1329`).
- **H-GR-93.** El rótulo del botón de rechazo es «Eliminar» (`.form.txt:1284`) mientras que su texto emergente dice «Botón que Rechaza por Calidad» (`.form.txt:1287`) y el código ejecuta un rechazo (`:105-120`). El rótulo describe una operación que no existe en este formulario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir la operación con un parámetro explícito en la ruta en lugar de leerla del rótulo; devolver el resultado de la decisión como estado de la propia operación en lugar de por eventos; mover la fecha de resolución del reloj del sistema a un dato que el servidor asigne de forma trazable; recalcular tareas en un único sitio después de cualquier decisión.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un riesgo en estado de aceptación o de retirada, la ventana muestra justificación y fechas, y activa solo las acciones que correspondan a ese estado.
2. Un técnico, o un riesgo no activo, ve la ventana sin acciones ejecutables.
3. Aprobar, rechazar y «quitar lo decidido» producen estados distintos y comunicados a la lista que abrió la ventana.
4. Abrir la ventana sin datos suficientes avisa y no deja una ventana a medias.

## Pendientes

- Si los cuatro campos de la ventana admiten edición en el binario y, por tanto, si lo escrito llega a algum sitio (H-GR-84).
- Qué hacen los cuatro eventos de retirada en su ausencia de manejador (H-GR-86).
- Cuándo se inicializa `EsTecnico` respecto a la primera apertura de esta ventana (H-GR-88).
- Qué hace exactamente `AjustarTamaño` con esta ventana (`Funciones Generales.bas:1043`, citado en `Form_frmSplash.md`).