# Form_FormCalidadTareaRiesgosAceptadosRetirados

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareaRiesgosAceptadosRetirados.cls` (391 líneas) y `src/forms/FormCalidadTareaRiesgosAceptadosRetirados.form.txt` · **PRD:** [README](../README.md)

Hoja de detalle de la worklist de calidad que reúne los **riesgos aceptados y los retirados pendientes de visar**. Se muestra como subformulario embebido, no se abre de forma independiente, y no escribe nada: solo proyecta datos del objeto `riesgo` ya cargado y ofrece dos botones de navegación. Su trabajo real está en los eventos que recibe del formulario de visado. Las citas son `Form_FormCalidadTareaRiesgosAceptadosRetirados.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`. Toma el riesgo de la **variable pública de instancia** del formulario padre.

| Quién lo carga | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalidadTareas` | `Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosAceptadosRetirados"` | `Form_FormCalidadTareas.cls:516` |
| `Form_FormCalidadTareas` (nodos) | `SeleccionarRiesgoAceptadoRetirado`, invocado desde `Arbol_NodeClick` | `Form_FormCalidadTareas.cls:136`, `:506-522` |

`SeleccionarRiesgoAceptadoRetirado` vacía antes el subformulario (`SourceObject = ""`) y copia `m_ObjSeleccionado` a `m_RiesgoSeleccionado` (`Form_FormCalidadTareas.cls:512-514`). Los nodos que disparan la carga son `RIESGOACEPTADOPORVISAR` y `RIESGORETIRADOPORVISAR` (`Form_FormCalidadTareas.cls:133-134`; `ArbolTareasCalidad.cls:455-486`).

Salidas: abre `FormCalidadRiesgoAceptadoRetiradoVisado` pasando `m_EsAceptacion` como séptimo argumento de `OpenForm` (`:49`) y `FormRiesgo` (`:99`). En ambos casos cierra antes la instancia previa con `DoCmd.Close … acSaveNo` (`:27-29`, `:77-79`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Coge el riesgo seleccionado y llama a `EstablecerDatos` | `:118-157`, `.form.txt:1` |
| (formulario) | — | `m_FormVisado_AceptacionAprobada` | Recalcula la worklist de calidad | `:160-186` |
| (formulario) | — | `m_FormVisado_AceptacionAprobadaQuitado` | Recalcula la worklist de calidad | `:315-339` |
| (formulario) | — | `m_FormVisado_AceptacionRechazada` | Recalcula la worklist de calidad | `:341-365` |
| (formulario) | — | `m_FormVisado_AceptacionRechazadaQuitada` | Recalcula la worklist de calidad | `:367-391` |
| `ComandoIrACalidad` | botón | `ComandoIrACalidad_Click` | Abre el formulario de visado | `:20-67`, `.form.txt:1548` (rótulo «Visar/Rechazar») |
| `ComandoVerRiesgo` | botón | `ComandoVerRiesgo_Click` | Abre `FormRiesgo` con el riesgo activo | `:69-116`, `.form.txt:1584` |
| `Autorizados` | cuadro de texto | — | Lista de usuarios autorizados, con saltos de línea | `:241`, `.form.txt:1521` |
| `lblAceptacionRetiro` | etiqueta | — | Rótulo que cambia entre «Mitigación» y «F. Retirado» | `:262`, `:281`, `.form.txt:1095` |
| `DiasSinRespuesta`, `FechaAprobacion`, `Fecha`, `MitigacionRetiro`, `Justificacion`, `Descripcion`, `CausaRaiz`, `Proyecto`, `NombreProyecto`, `Supervisor` | varios | — | Campos proyectados | `:234-298`, `.form.txt:560-900` |
| `IDExpediente`, `IDProyecto`, `IDEdicion`, `IDRiesgo` y sus etiquetas | varios | — | Identificadores, visibles solo para administrador | `:209-231`, `.form.txt:652-880` |

## Flujos

### Apertura (`Form_Open`, `:118-157`)

1. Reloj de arena y `DoEvents` (`:122-124`).
2. Si el padre no tiene riesgo seleccionado, **sale sin más** (`:127-129`).
3. Copia el riesgo, pone `AllowEdits = False` (`:130-136`) y llama a `EstablecerDatos` (`:139-140`).
4. En error, avisa y **se cierra a sí mismo** con `acSaveNo` (`:157`).

### Carga de datos (`EstablecerDatos`, `:188-313`)

1. Deshabilita los dos botones (`:201-202`).
2. Muestra u oculta los cuatro identificadores según `EsAdministrador` o `EsAdministradorConectadoIntegramente` (`:209-231`).
3. Proyecta proyecto, supervisor, descripción con el código de riesgo entre paréntesis y autorizados (`:234-241`).
4. **Deriva el modo de la lista del estado del riesgo**: `RiesgoEnAceptacion` → `m_EsAceptacion = "Sí"`; si no, `RiesgoEnRetirada` → `"No"`; si tampoco, error «El Riesgo no está camino de aceptarse ni de retirarse» (`:244-253`). Los predicados están en `Funciones Generales.bas:6595` y `:6620`.
5. Ramifica por `m_EsAceptacion` (`:261-299`): en el caso de aceptación rotula «Mitigación», fecha de justificación de aceptación, `MitigacionRetiro = "Aceptar"` y días sin respuesta de aceptación; en el de retirada rotula «F. Retirado», fecha de retirada y días sin respuesta de retirada.
6. El botón de visar/rechazar solo se habilita si **no** es técnico (`:256-259`); el de ver riesgo, con el mismo criterio (`:300-303`).

### Normal — visar o rechazar (`ComandoIrACalidad_Click`, `:20-67`)

Cierra la instancia previa del formulario de visado (`:27-29`), resuelve el riesgo siStill no lo tiene —primero desde el padre, y si no por identificador— (`:30-46`), publica el riesgo activo en la global `m_ObjRiesgoActivo` (`:48`), abre el visado con `m_EsAceptacion` como argumento (`:49`) y se queda escuchando los cuatro eventos de resultado (`:50-52`).

### Alternativo — ver la ficha del riesgo (`ComandoVerRiesgo_Click`, `:69-116`)

Mismo patrón de resolución del riesgo, pero marca la global `m_EsAlta = EnumSiNo.No` (`:81`) para que la ficha se abra en modo consulta y abre `FormRiesgo` (`:99`).

### De error

Los seis manejadores repiten el mismo bloque: quitar el reloj, y **según `Err.Number`**, o bien componer el mensaje, avisar por correo al administrador y mostrar un `MsgBox` crítico, o bien mostrar solo un `MsgBox` de advertencia (`:58-66`, `:105-115`, `:148-156`, `:176-185`, `:330-338`, `:352-360`, `:382-390`). Los cuatro últimos no llegan a cerrar el formulario.

## Reglas, validaciones y cálculos

- **El formulario es de solo lectura.** `AllowEdits = False` se fija en `Form_Open` (`:136`) y de nuevo al final de `EstablecerDatos` (`:306`). No hay ningún control de entrada: la decisión se toma en el formulario de visado.
- **El modo depende del estado, no de la lista.** Aceptado y retirado comparten hoja, y la etiqueta «Mitigación» o «F. Retirado» es lo que los distingue (`:262`, `:281`).
- **La habilitación del botón depende del rol**, y se comprueba dos veces (`:256`, `:300`); la primera comprobación queda anulada por la segunda si el rol cambia entremedias.
- **`EstablecerTareasCalidad` se invoca siempre igual**: `EnumSiNo.Sí, EnumSiNo.No` (`:167`, `:321`, `:347`, `:373`), es decir, reconstruye la worklist entera. No distingue el motivo del evento, solo el hecho de que algo cambió.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo seleccionado | `Form_FormCalidadTareas.m_RiesgoSeleccionado` | `:130` |
| Riesgo por identificador | `Constructor.getRiesgo` | `:37`, `:88` |
| Datos del proyecto y autorización | propiedades del objeto `riesgo` | `:234-241` |
| Recálculo de contadores | `EstablecerTareasCalidad` → `TareasCalidad` | `:167`; `Funciones Generales.bas:2054`, `:2113` |
| Riesgo activo publicado | global `m_ObjRiesgoActivo` | `:48`, `:98`; `Variables Globales.bas:80` |

Sin escrituras en base de datos: la única escritura es en memoria de sesión.

## Coexistencia

El formulario no tiene coexistencia propia; hereda la de su padre, que es un subformulario conmutado por `SourceObject`. La dependencia que sí trae es la escucha de eventos de otro formulario mediante `WithEvents` (`:16-17`), un patrón que en web se traduce a un retorno de la acción al servidor, no a un evento de ventana.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-112.** `Form_Open` pone el reloj de arena en `:123` y sale en `:127-129` **sin apagarlo**. Si el padre no tiene riesgo seleccionado, el puntero de espera queda activo en el formulario padre, que es quien lo hereda.
- **H-GR-113.** La condición de salida de `Form_Open` comprueba la variable del padre y, en las líneas siguientes, vuelve a comprobar `m_Riesgo` (`:127-132`) que se acaba de asignar: la segunda comprobación solo puede ser falsa si la asignación falló, y en ese caso la línea siguiente lo trata como error válido. Es una guarda redundante, no un fallo.
- **H-GR-114.** Los cuatro manejadores de evento del formulario de visado hacen exactamente lo mismo (`:160-186`, `:315-339`, `:341-365`, `:367-391`). Diferencian el motivo en el texto del mensaje de error y nada más: el efecto es idéntico para aprobación y rechazo, y para quitar y para poner. Conviene confirmar si algún caso necesitaba un comportamiento distinto.
- **H-GR-115.** `EstablecerDatos` fija `Me.ComandoIrACalidad.Enabled` en `:256-259` y vuelve a fijar `Me.ComandoVerRiesgo.Enabled` en `:300-303`; el segundo bloque se ejecuta **después** del `If m_EsAceptacion` y no depende de él, de modo que la comprobación de rol se hace dos veces sobre objetos distintos sin motivo aparente.
- **H-GR-116.** El rótulo del formulario de visado se decide con una cadena libre, `m_EsAceptacion` (`"Sí"` / `"No"`, `:246`, `:250`), que se pasa como argumento a `OpenForm` (`:49`). Es un contrato entre dos formularios con dos valores literales en castellano, sujeto a la convención de idioma del lector y no al tipo `EnumSiNo` que usa el resto del formulario (`:256`).
- **H-GR-117.** `Justificacion` solo se rellena si el valor del riesgo no viene vacío (`:265-267`, `:285-287`); en caso contrario el control conserva lo que tuviera, y como el subformulario se reutiliza entre nodos del árbol (`:512`), conviene verificar que no puede quedar texto de un riesgo anterior.
- **H-GR-118.** La fecha de aprobación se muestra en el mismo control que la de rechazo (`:271-274` y `:277-280`, y `:291-294` y `:297-300`): el segundo `If` sobrescribe el primero cuando ambas fechas existen. La última que se escribe gana, y el usuario no puede saber cuál se está viendo.
- **H-GR-119.** `EstablecerDatos` declara que devuelve `String` (`:188`) y devuelve siempre `Empty`, porque el resultado viaja en el argumento `ByRef p_Error`. La firma sugiere un contrato de valor de retorno que no existe; es la convención del resto de la aplicación, pero conviene tenerla presente al portar.
- **H-GR-120.** La reconstrucción completa de la worklist (`EstablecerTareasCalidad` con reseteo) se dispara en cada uno de los cuatro eventos (`:167`, `:321`, `:347`, `:373`). Con trabajo de calidad acumulado, esto recalcula contadores y árbol por cada aprobación o rechazo, también por los «quitado».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: servir la hoja como vista de solo lectura encajada en la propia ruta del árbol, sin formulario hijo; sustituir la escucha de eventos por un retorno de la acción a la lista que refresca solo la fila afectada; mostrar fecha de aprobación y de rechazo en campos separados en lugar de uno solapado; expresar el modo con un tipo enumerado en lugar de un literal en castellano.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al elegir un riesgo aceptado o retirado pendiente de visar, la hoja muestra sus datos y el rótulo del modo correcto.
2. Un usuario técnico ve la hoja sin el botón de visar/rechazar.
3. Un administrador ve además los cuatro identificadores.
4. Tras aprobar o rechazar, la lista de tareas de calidad refleja el cambio sin recargar a mano.
5. Si el riesgo no está en ninguno de los dos estados, se avisa y no se muestran datos que no correspondan.

## Pendientes

- Qué estados exactos de `EnumRiesgoEstado` llegan en producción a esta hoja y si alguno de los dos predicados puede dar «Sí» a la vez.
- Si el campo de justificación conserva valor al reutilizarse el subformulario entre nodos.
- Qué hace exactamente `TareasCalidad` al reconstruirse: consultas, caché en disco o solo contadores en memoria.
- Si `FormularioAbierto` (`Funciones Generales.bas:832`) cubre también subformularios embebidos, porque el cierre previo de `:27-29` se hace por nombre.