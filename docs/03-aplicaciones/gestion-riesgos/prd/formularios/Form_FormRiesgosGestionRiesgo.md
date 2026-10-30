# Form_FormRiesgosGestionRiesgo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgosGestionRiesgo.cls` (374 líneas) y `src/forms/FormRiesgosGestionRiesgo.form.txt` (2078 líneas; el código empieza en `.form.txt:1713`) · **PRD:** [README](../README.md)

Ficha resumida de un riesgo, tal y como se ve cuando el usuario selecciona su nodo en el árbol de la edición. **No se abre**: se carga como subformulario del contenedor `FormRiesgosGestion` y se recarga cada vez que cambia el nodo. No escribe nada. Las citas son `Form_FormRiesgosGestionRiesgo.cls:línea` salvo indicación.

Es la vista de lectura del riesgo y el punto desde el que se pasa a la edición (`FormRiesgo`), al plan de mitigación o al de contingencia (`FormPlanPrincipal`) y al informe HTML. Muestra también el **histórico de estados** del riesgo hasta la edición publicada.

## Entrada y salida

No acepta `OpenArgs`. Todo el contexto le llega por estado compartido:

| Contexto | Evidencia |
|---|---|
| Se selecciona un nodo de riesgo en el árbol | `Form_FormRiesgosGestion.cls:296` (`SourceObject = "FormRiesgosGestionRiesgo"`) |
| Riesgo activo | global `m_ObjRiesgoActivo`, fijado en `Form_FormRiesgosGestion.cls:293` |
| Edición y proyecto | globales `m_ObjEdicionActiva` y `m_ObjProyectoActivo` (`Variables Globales.bas:75-80`) |

Su `Form_Load` (`:149-165`) es, de hecho, la routine de cambio de contexto: si no hay riesgo activo, sale sin avisar y deja el formulario vacío (`:155-158`).

Salidas:

| Botón | Destino | Evidencia |
|---|---|---|
| `ComandoDetalle` («Ver detalles») | `FormRiesgo`, cerrando antes el que hubiera abierto | `:83-88` |
| `ComandoAltaPM` («Alta P. mitigación») | `FormPlanPrincipal` con `«Sí»` en `OpenArgs` | `:70-72` |
| `ComandoAltaPC` («Alta P. contingencia») | `FormPlanPrincipal` con `«No»` en `OpenArgs` | `:47-49` |
| `ComandoVerInformeRiesgo` («Informe») | genera el informe HTML del riesgo y lo abre | `:136-138` |
| `ComandoEliminar` («Eliminar riesgo») | borra el riesgo y redibuja el árbol del contenedor | `:112-125` |

El argumento que recibe `FormPlanPrincipal` es el mismo flag `m_EsMitigacion` que el contenedor usa para el subformulario del plan (`Form_FormPlanPrincipal.cls:134`). Cuando se da de alta un plan, el evento `m_FormPlan_PlanNuevo` (`:351-371`) inserta el nodo nuevo en el árbol del contenedor y deja seleccionado ese nodo: el árbol se entera del alta desde aquí, no al revés.

## Controles y eventos

| Control | Tipo | Qué muestra | Evidencia |
|---|---|---|---|
| `Detalle` | subformulario | El resto de campos del riesgo | `.form.txt:566`, `:226-229` |
| `lblTitulo` | etiqueta | Código del riesgo; en rojo y con asteriscos si es alto o muy alto | `:196-202`, `.form.txt:1424` |
| `Estado`, `FechaEstado` | cuadros de texto | Estado calculado y fecha de ese estado | `:265-266`, `.form.txt:714`, `:1334` |
| `FechaMaterializado`, `FechaRetirado` | cuadros de texto | Se colorean en rojo si tienen fecha, en verde si no | `:250-262`, `Constantes.bas:99-100` |
| `CausaRaiz` | cuadro de texto | Solo si el proyecto exige riesgo de biblioteca calculado | `:231-243`, `.form.txt:620` |
| `lblHistorico` + `lstEstadosHistoricos` | etiqueta y lista | Histórico de estados con su fecha | `:297-328`, `.form.txt:1448`, `:1479` |
| `IDRiesgo` | cuadro de texto oculto | Identificador del riesgo | `.form.txt:1396` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:149-338`)

1. Deja el formulario en solo lectura (`:153`) y, si no hay riesgo activo, sale en silencio (`:155-158`).
2. `EstablecerDatos` desactiva primero todos los botones (`:186-191`) y luego los activa uno a uno, de modo que el estado por defecto es «nada habilitado» (`:215-224`).
3. Compone el título: rojo y con cuatro asteriscos cuando el riesgo es alto o muy alto (`:196-198`).
4. Pinta los campos, **omitiendo los que están vacíos** (`:225-262`).
5. Colorea las fechas de materialización y retirada según tengan o no valor (`:250-262`).
6. Calcula la **publicabilidad** del riesgo y colorea el botón de informe en consecuencia (`:277-292`).
7. Carga el **histórico de estados** (`:297-328`).

### Alternativo: dar de alta un plan (`:33-77`)

Ambos botones comprueban primero `FichaRiesgoCompleta`. Si la ficha está incompleta, avisan, abren la edición del riesgo y **no salen** (`:37-46`, `:60-68`); si el botón de detalle estuviera deshabilitado, continuarían y abrirían el plan de todos modos.

### Alternativo: informe (`:131-143`)

Delega en `GenerarInformeRiesgoHTML` (`InformeRiesgoHTML.bas:14`) pasando el identificador del riesgo y el handle de la ventana de Access, y cierra el indicador de progreso (`:138`). la publicación se decide antes, en la carga (`:277-292`), y solo se refleja en el color del botón.

### De error

Un único manejador, `GestionarError` (`:340-349`): si el número no es 1000, compone el mensaje, **envía correo al administrador** y avisa con un `MsgBox` crítico; si es 1000, muestra una advertencia con `m_Error`, que a menudo está vacío porque los `Err.Raise 1000` del código no llevan descripción.

## Reglas, validaciones y cálculos

- **La máquina de estados no está aquí.** El formulario no decide el estado: lee `ESTADOCalculadoTexto` (`Riesgo.cls:2513`), que a su vez lo obtiene de la enumeración `EnumRiesgoEstado` (`Constantes.bas:153-167`, catorce estados). Aquí solo se pinta.
- **El permiso de edición tiene una rama muerta**: se concede igual se esté o no autorizado (`:208-213`).
- **La lista de campos depende del proyecto, no del riesgo**: `RequiereRiesgoDeBibliotecaCalculado` decide si aparece la causa raíz y cambia el alto de la descripción (`:231-243`), con dos constantes de diseño en twips (`:24-25`).
- **La fecha se usa como color**: tener fecha de materialización se pinta en rojo y no tenerla en verde (`:250-262`), con dos constantes globales compartidas (`Constantes.bas:99-100`).
- **El histórico se corta por la publicación de la edición**: `getEstadosDiferentesHastaEdicion` recibe la edición, el código de riesgo y su fecha de publicación (`Funciones Generales.bas:7516`, llamada en `:302-307`). El resultado es un diccionario de cadenas con estado y fecha separados por una barra vertical (`:314`), y el formulario los reformatea para la lista (`:317-326`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Datos del riesgo | global `m_ObjRiesgoActivo` | `:193-271` |
| Evaluación de publicabilidad | `PublicabilidadRiesgo.bas:40` y `:263` | `:279-280` |
| Histórico de estados | `Funciones Generales.bas:7516` | `:302-307` |
| Errores de la ficha | `Riesgo.ErroresRiesgoTexto` (`Riesgo.cls:1310`) | `:270` |
| Informe | `InformeRiesgoHTML.bas:14` | `:136` |

Solo escribe cuando el usuario pulsa el botón de borrar: entonces llama a `Borrar` sobre el riesgo (`:112`) y redibuja el árbol del contenedor (`:123-124`).

## Coexistencia

Es la única vista del riesgo que **no depende de un fichero generado**: lee de la base de datos y devuelve objetos. Por eso es la candidata más directa a reutilizar en la web. Su dependencia fuerte es la contraria: necesita que el contenedor haya fijado las globales y que el árbol siga vivo, porque el borrado manipula el árbol del contenedor desde aquí (`:117-124`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-172.** Si no hay riesgo activo, `Form_Load` sale sin mostrar nada ni avisar (`:155-158`): el subformulario queda en blanco con todos sus botones desactivados y ninguna explicación.
- **H-GR-173.** El permiso de edición concede acceso con o sin autorización del usuario: las dos ramas asignan `True` (`:210-213`). Como el resto de comandos sí comprueban `UsuarioConectadoAutorizado` (`:105`), el criterio es incoherente dentro del mismo formulario.
- **H-GR-174.** El aviso de ficha incompleta dice «No se puede editar **nin** dar de alta…» (`:39-40`, `:62-63`): errata en el texto que ve el usuario, en los dos botones de alta.
- **H-GR-175.** Los `Err.Raise 1000` casi nunca llevan descripción (`:113`, `:160`, `:272`, `:281`, `:309`, `:364`, `:366`), de modo que el `MsgBox` de advertencia sale con `m_Error` vacío (`:348`): el usuario ve «Advertencia» y nada más.
- **H-GR-176.** El histórico se carga dentro de un `On Error Resume Next` deliberado para tolerar que el control no exista (`:298`, restaurado en `:329`). Si la carga falla, el `Err.Raise 1000` de `:309` **no se propaga** y el formulario se presenta sin histórico, sin indicarlo.
- **H-GR-177.** La lista del histórico trae **datos de ejemplo en el diseño** (`.form.txt:1481`: dos estados con fecha). El código la vacía al cargar (`:299`), pero si la carga se aborta antes, esos datos ficticios llegan a verse.
- **H-GR-178.** El texto de la etiqueta del histórico está sin cerrar: «Histórico de Estados (Hasta Edición Actual» (`.form.txt:1449`).
- **H-GR-179.** `m_Error` es una cadena compartida por todo el formulario (`:27`) que se reinicia al inicio de cada manejador pero no al de `EstablecerDatos`, y que además se sobrescribe con la propiedad `Error` del riesgo (`:271`): si esa propiedad arrastra un fallo anterior, el mensaje mostrado ya no corresponde a lo que ha fallado.
- **H-GR-180.** La publicabilidad se calcula entera y se **descarta**: `m_Publicable` (`:280`) y `m_ChecksPublicabilidad` (`:280`) no se usan después, y solo queda el color del botón (`:283-289`). El usuario ve un botón-rojo sin saber qué le falta.
- **H-GR-181.** Los campos se pintan solo si no están vacíos (`:226-262`): un valor vacío y un valor ausente son indistinguibles, y lo que había en el control de una carga anterior puede quedar si el objeto se reutiliza.
- **H-GR-182.** El borrado pide confirmación (`:108`), borra (`:112`) y luego confirma el éxito con un `MsgBox` (`:125`) sin comprobar nada más. Además manipula las estructuras internas del contenedor —la colección de riesgos mostrados y el árbol— desde aquí (`:117-124`), acoplamiento directo entre dos formularios.
- **H-GR-183.** Se declaran tres eventos públicos —`RiesgoAceptado`, `RiesgoRetirado`, `RiesgoMaterializado` (`:13-15`)— y **no se lanzan nunca en el código fuente**. OConsumers esperaría un evento que no existe, o la funcionalidad se movió a otro sitio.
- **H-GR-184.** La distinción entre mitigación y contingencia viaja como texto en `OpenArgs` («Sí» / «No», `:48`, `:71`) y se reinterpreta en `FormPlanPrincipal`; el mismo flag se llama `m_EsMitigacion` en dos formularios con valores que no se corresponden con su nombre.
- **H-GR-185.** La evaluación de publicabilidad se ejecuta **solo si la edición está activa** (`:277`), y con ella la visibilidad del botón de informe. En una edición cerrada no hay forma de saber si un riesgo era publicable.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: servir esta ficha como recurso de solo lectura del riesgo, con el estado como dato calculado y no como texto; incluir el histórico de estados en la misma respuesta en lugar de una lista aparte; mostrar las carencias de publicabilidad en lugar de codificarlas en el color del botón; decidir quién es el dueño del borrado (esta ficha o la vista de listado), para que el árbol no tenga que ser manipulado desde el detalle.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un riesgo seleccionado, la ficha muestra sus datos, su estado y su fecha de estado.
2. Un riesgo alto o muy alto se distingue de inmediato, y un riesgo sin materializar o sin retirada también.
3. El histórico de estados aparece con el estado y su fecha, hasta la edición publicada.
4. Un riesgo con la ficha incompleta no permite dar de alta un plan, y el aviso explica por qué.
5. Solo el usuario autorizado y solo en una edición activa ve las acciones que modifican el riesgo.
6. Un fallo al cargar el riesgo se comunica con su causa, y nunca se muestran datos de ejemplo.

## Pendientes

- Si existe algún código que lance los eventos `RiesgoAceptado`, `RiesgoRetirado` o `RiesgoMaterializado` (`:13-15`), y qué consumidor los esperaría.
- Qué distingue «fecha de materialización ausente» de «fecha de materialización correcta»: el código usa el rojo para la que existe (`Constantes.bas:99-100`) y conviene saber si es intencionado.
- Qué reglas produce `EvaluarPublicabilidadRiesgo` (`PublicabilidadRiesgo.bas:40`) y cuál de ellas consulta el usuario cuando el botón aparece apagado.
- Cómo se recalcula el histórico cuando se corrigen datos de ediciones anteriores al corte de publicación.