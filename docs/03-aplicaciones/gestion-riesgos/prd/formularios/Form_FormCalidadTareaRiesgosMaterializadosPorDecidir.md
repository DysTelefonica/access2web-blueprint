# Form_FormCalidadTareaRiesgosMaterializadosPorDecidir

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareaRiesgosMaterializadosPorDecidir.cls` (203 líneas) y `src/forms/FormCalidadTareaRiesgosMaterializadosPorDecidir.form.txt` · **PRD:** [README](../README.md)

Hoja de detalle de la worklist de calidad que reúne los **riesgos materializados que todavía no se han decidido**. Se muestra como subformulario embebido, no se abre de forma independiente, y no escribe nada: solo proyecta datos del par materialización/riesgo ya cargado y ofrece dos botones de navegación. Es la más pobre de las tres hojas de riesgo: apenas muestra información y delega todo el trabajo en otro formulario. Las citas son `Form_FormCalidadTareaRiesgosMaterializadosPorDecidir.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`. Toma el riesgo materializado de la **variable pública de instancia** del formulario padre.

| Quién lo carga | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalidadTareas` | `Me.FormDetalle.SourceObject = "FormCalidadTareaRiesgosMaterializadosPorDecidir"` | `Form_FormCalidadTareas.cls:534` |
| `Form_FormCalidadTareas` (nodos) | `SeleccionarRiesgoMaterializado`, invocado desde `Arbol_NodeClick` | `Form_FormCalidadTareas.cls:161`, `:524-540` |

A diferencia de las otras dos hojas, la selección aquí es de un `RiesgoMaterializacion`, no de un `riesgo`: `Arbol_NodeClick` trocea la clave del nodo, extrae el identificador, llama a `Constructor.getRiesgoMaterializado` y lo publica en `m_RiesgoMaterializadoSeleccionado` (`Form_FormCalidadTareas.cls:141-166`). La clave del nodo se construye en `ArbolTareasCalidad.cls:328`.

Salidas: abre `FormCalidadRiesgoMaterializaciones` (`:32`) y `FormRiesgo` (`:77`), cerrando antes la instancia previa de cada una (`:27-29`, `:58-60`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Resuelve el par materialización/riesgo y llama a `EstablecerDatos` | `:95-140`, `.form.txt:1` |
| `ComandoIrFormularioCalidad` | botón | `ComandoIrFormularioCalidad_Click` | Abre el formulario de materializaciones | `:21-48`, `.form.txt:903` (rótulo «Ir formulario calidad») |
| `ComandoVerRiesgo` | botón | `ComandoVerRiesgo_Click` | Abre `FormRiesgo` con el riesgo activo | `:50-93`, `.form.txt:938` (rótulo «Ver riesgo») |
| `DescripcionRiesgo` | cuadro de texto | — | Descripción con el código de riesgo entre paréntesis | `:178`, `.form.txt:308` |
| `Proyecto`, `Supervisor` | cuadros de texto | — | Nombre del proyecto y usuario de calidad | `:179-180`, `.form.txt:341`, `:374` |
| `DetalleEdicion` | cuadro de texto | — | Número de edición y si está activa | `:187`, `.form.txt:470` |
| `FechaMaterializado` | cuadro de texto | — | Fecha de materialización, si es fecha válida | `:189-191`, `.form.txt:407` |
| `lblTitulo` | etiqueta | — | Título fijo «Riesgo materializado NC por decidir» | `:195-196`, `.form.txt:439` |
| `IDExpediente`, `IDProyecto`, `IDEdicion`, `IDRiesgo` y sus etiquetas | varios | — | Identificadores, visibles solo para administrador | `:150-167`, `.form.txt:652-880` |

## Flujos

### Apertura (`Form_Open`, `:95-140`)

1. Reloj de arena y `DoEvents` (`:99-101`).
2. Si el padre no tiene materialización seleccionada, avisa «No se ha seleccionado ningún riesgo materializado» y **se cierra** (`:104-107`, `:139`).
3. Copia la materialización y resuelve **qué riesgo se muestra** (`:109-113`): prefiere `RiesgoEnEdicionActiva` y, si no existe, cae a la materialización. Esta preferencia es la clave del comportamiento de la hoja.
4. Si tampoco hay riesgo, repite el mismo aviso (`:115-118`).
5. Guarda la edición de la materialización (`:119`) y llama a `EstablecerDatos` (`:121-124`).

### Carga de datos (`EstablecerDatos`, `:142-203`)

1. Vacía explícitamente el error de salida (`:149`).
2. Muestra u oculta los cuatro identificadores según `EsAdministrador` o `EsAdministradorConectadoInicialmente` (`:150-167`).
3. Habilita el botón de ver riesgo (`:169`).
4. Proyecta identificadores y descripción del riesgo (`:174-178`), proyecto y supervisor (`:179-180`), el estado de la edición (`:182-187`) y la fecha de materialización (`:189-191`).
5. Fija un título constante (`:195-196`).

### Normal — ir al formulario de calidad (`ComandoIrFormularioCalidad_Click`, `:21-48`)

Cierra la instancia previa, publica el riesgo en la global `m_ObjRiesgoActivo` (`:31`) y abre `FormCalidadRiesgoMaterializaciones` (`:32`).

### Alternativo — ver la ficha del riesgo (`ComandoVerRiesgo_Click`, `:50-93`)

Marca `m_EsAlta = EnumSiNo.No` (`:61`), intenta deducir el riesgo (`:62-75`) y abre `FormRiesgo` (`:77`).

### De error

Los tres manejadores repiten el bloque habitual: quitar el reloj y, **según `Err.Number`**, componer el mensaje, avisar por correo al administrador y mostrar un `MsgBox` crítico, o bien mostrar solo un `MsgBox` de advertencia (`:40-47`, `:84-92`, `:130-138`). En `Form_Open` el cierre del formulario es adicional (`:139`).

## Reglas, validaciones y cálculos

- **La hoja no decide nada.** Ni una línea escribe en base de datos: el trabajo está en `FormCalidadRiesgoMaterializaciones`, y esta hoja solo navega hasta él.
- **El riesgo mostrado puede no ser el materializado.** Si la materialización tiene riesgo en la edición activa, se muestra ese (`:109-110`); el título, en cambio, siempre habla de «Riesgo materializado NC por decidir» (`:195`), con lo que el título y el contenido pueden describir objetos distintos.
- **La edición mostrada y la edición del supervisor vienen de fuentes distintas.** `DetalleEdicion` y `Supervisor` se leen de la materialización (`:180`, `:187`), mientras que los cuatro identificadores se leen del riesgo resuelto (`:174-177`).
- **No hay filtro por rol.** `EstablecerDatos` habilita `ComandoVerRiesgo` siempre (`:169`, `:188`), a diferencia de las otras dos hojas de riesgo, donde el botón depende de que el usuario no sea técnico.
- **La fecha solo se muestra si es una fecha válida** (`:189-190`): en caso contrario el control conserva su valor anterior.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Materialización seleccionada | `Form_FormCalidadTareas.m_RiesgoMaterializadoSeleccionado` | `:108` |
| Riesgo en edición activa / riesgo de la materialización | propiedades de `RiesgoMaterializacion` | `:109-113` |
| Edición de la materialización | `.Edicion` | `:119`, `:182-187` |
| Identificadores, descripción, proyecto, supervisor | propiedades del objeto `riesgo` y de su edición | `:174-180` |
| Fecha de materializado | `m_Riesgo.FechaMaterializado` | `:189-190` |
| Riesgo activo publicado | global `m_ObjRiesgoActivo` | `:31`, `:76`; `Variables Globales.bas:80` |

Sin escrituras en base de datos.

## Coexistencia

No aporta coexistencia propia. Lo relevante es que el resultado de la decisión se produce en otro formulario de la misma aplicación y que esta hoja no tiene retorno: al cerrar `FormCalidadRiesgoMaterializaciones`, no hay evento que refresh la lista.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-121.** `ComandoVerRiesgo_Click` **nunca asigna `m_Riesgo`**. El bloque `:62-75` solo comprueba si hay selección materializada y si el control identificador está vacío, pero en ningún momento hace `Set m_Riesgo = …`. Tras el paso, `:76` publica en la global un objeto `Nothing`, y `FormRiesgo` se abre sin riesgo activo. La función que debería resolverlo se sustituyó por dos guardas que solo detectan el problema sin resolverlo.
- **H-GR-122.** Consecuencia de lo anterior: el segundo `If m_Riesgo Is Nothing` de `:68` es redundante con el primero, porque en `:62` no se ha modificado nada. Es la firma de una reescritura incompleta: la asignación que falta se sustituyó por una segunda comprobación.
- **H-GR-123.** `EstablecerDatos` es `Public` (`:142`) mientras que en las otras dos hojas de riesgo es `Private`. Con `Option Explicit` declarado y sin efecto en el resto del módulo, conviene confirmar si el modificador responde a una necesidad real o es un descuido.
- **H-GR-124.** El título se fija con una constante en línea (`:195`), sin la variable local que lo acompaña. Ningún otro rótulo del formulario depende de la materialización, de modo que el usuario no distingue en la hoja un riesgo materializado de otro.
- **H-GR-125.** `Me.ComandoVerRiesgo.Enabled = True` se ejecuta dos veces (`:169` y `:188`), la segunda dentro del `With` anidado. La primera basta y la segunda no depende de ningún dato.
- **H-GR-126.** `Proyecto` (`:179`), `Supervisor` (`:180`) y `DetalleEdicion` (`:187`) se leen de la materialización, pero los cuatro identificadores vienen del riesgo resuelto (`:174-177`), que puede ser el de otra edición. Conviene verificar que la combinación que se muestra es consistente cuando existe riesgo en edición activa.
- **H-GR-127.** `EstablecerDatos` no fija `AllowEdits = False`. Los controles son todos de lectura por omisión en el diseño, pero el formulario no expresa esa condición: cualquier campo desbloqueado en el `.form.txt` sería editable sin que el código lo impida. Las otras dos hojas sí lo fijan.
- **H-GR-128.** La apertura devuelve el error «No se ha seleccionado ningún riesgo materializado» en dos casos distintos (`:105` y `:116`) con el mismo texto: el usuario no puede distinguir que falta la selección de que la materialización no tiene riesgo en ninguna edición.
- **H-GR-129.** No hay ningún evento que actualice la worklist al cerrar o guardar `FormCalidadRiesgoMaterializaciones`, a diferencia de la hoja de aceptados y retirados, que sí escucha cuatro eventos y reconstruye la lista. Conviene comprobar si la actualización la hace el formulario padre por otro camino.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mostrar el riesgo materializado y el de la edición activa como dos elementos distinguibles en lugar de elegir uno en silencio; devolver a la lista el resultado de la decisión como acción explícita; sustituir el salto de dos formularios por un paso único de decisión; derivar el título del propio dato en lugar de una constante.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al elegir un riesgo materializado pendiente de decidir, la hoja muestra su proyecto, edición, supervisor y fecha de materialización.
2. Si el riesgo está en edición activa, se muestra ese y queda identificado como tal.
3. El botón de ver riesgo abre la ficha del riesgo que la hoja está mostrando, no otro.
4. Un usuario no administrador no ve los identificadores.
5. Sin selección válida, se avisa con un mensaje que distinga los dos motivos y no se abre nada a medias.

## Pendientes

- Si `RiesgoEnEdicionActiva` y `riesgo` pueden ser el mismo objeto o si la materialización puede apuntar a ediciones distintas simultáneamente.
- Qué estados de materialización llegan a esta lista y qué se considera «materializado no decidido».
- Qué persiste `FormCalidadRiesgoMaterializaciones` al guardar y si el padre actualiza la lista por su cuenta.
- Si el `.form.txt` deja algún control desbloqueado, dado que el código no impone `AllowEdits = False`.