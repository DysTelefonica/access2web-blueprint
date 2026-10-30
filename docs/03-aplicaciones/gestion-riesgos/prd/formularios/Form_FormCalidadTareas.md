# Form_FormCalidadTareas

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareas.cls` (542 líneas) y `src/forms/FormCalidadTareas.form.txt` · **PRD:** [README](../README.md)

Árbol de tareas pendientes de calidad. Es el punto de trabajo del perfil: a la izquierda el árbol de tareas, a la derecha un subformulario que va cambiando según el tipo de nodo pulsado (explicación, retipificación, Risks aceptado o retirado, materialización o detalle de edición). Todo el reparto se decide por el **tipo de nodo** que devuelve `ArbolTareasCalidad`. No escribe directamente en base de datos: recalcula el contador de tareas y recarga el árbol. Las citas son `Form_FormCalidadTareas.cls:línea` salvo indicación.

## Entrada y salida

No acepta `OpenArgs`. Dos entradas, ambas indirectas:

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_frmSplash` | `DoCmd.OpenForm "FormCalidadTareas"` si hay tareas pendientes | `Form_frmSplash.cls:85` |
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormCalidadTareas"` | `Form_Form0BDOpciones.cls:207` |
| `Funciones Generales` | Refresco externo: si el formulario está abierto, llama a `CargarArbol` | `Funciones Generales.bas:2139-2140` |

La selección se publica en el miembro público `m_EdicionSeleccionada` (`:37`), que es lo que lee `Form_FormCalidadTareasDetalleEdicion.cls:252`. **Es el contrato entre los dos formularios de este lote**: el árbol no abre el detalle, lo deja incrustado dentro de sí.

Salidas del propio árbol: incrusta cinco subformularios distintos en `FormDetalle` (`FormCalidadTareaExplicacion` `:479`, `FormCalidadTareasDetalleEdicion` `:50`, `FormCalidadTareaRiesgosRetipificacion` `:498`, `FormCalidadTareaRiesgosAceptadosRetirados` `:516`, `FormCalidadTareaRiesgosMaterializadosPorDecidir` `:534`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Comprueba el rol, monta el combo de responsables, fija el título y carga el árbol | `:296-342` |
| `Arbol` | ActiveX TreeView | `Arbol_NodeClick` | Despacha por tipo de nodo y recarga el subformulario | `:59-185`, `.form.txt:544` |
| `FormDetalle` | subformulario | — | Contenedor del subformulario activo | `.form.txt:1030` |
| `ListaImagenes` | ActiveX ImageList | — | Icons del árbol, la pasa a `ArbolTareasCalidad` | `.form.txt:1052`; `:349-350` |
| `Sep` | separador | `MouseDown`/`MouseMove`/`MouseUp` | Arrastre que reparte el ancho entre árbol y detalle | `:367-416`, `.form.txt:1012` |
| `ComboResponsableCalidad` | combo | `AfterUpdate` | Refiltra las tareas del responsable elegido | `:272-294`, `.form.txt:1560` |
| `ComandoActualizarContador` | botón | `Click` | Resuelve el responsable y recalcula tareas + árbol | `:187-237`, `.form.txt:1477` |
| `ComandoAyuda` | botón | `Click` | Abre la ayuda | `:239-270`, `.form.txt:1596` |

## Flujos

### Apertura (`Form_Load`, `:296-342`)

1. Reloj de arena y `m_Error = ""`.
2. **Si el usuario es técnico**, se aborta con el mensaje de que el formulario es exclusivo de Calidad (`:305-308`).
3. Habilita el combo y lo puebla con `EstablecerComboUsuariosCalidad` (`:309-313`).
4. Si es de Calidad, preselecciona en el combo al usuario conectado (`:314-316`).
5. Compone el título con el rótulo del usuario conectado (`:318`).
6. `CargarArbol` (`:319`) y coloca el separador junto al subformulario (`:324`), vaciando el detalle si ya había algo (`:325-327`).

El manejador de error, común a los cinco procedimientos, **notifica por correo y avisa con `MsgBox`**, y en este caso concreto **cierra el formulario** (`:337-342`), que es el único que lo hace.

### Normal (`Arbol_NodeClick`, `:59-185`)

1. Si el árbol aún no está construido, lo construye y conecta `Arbol` y `ListaImagenes` (`:73-77`).
2. Pregunta el **tipo de nodo** (`:81`); si el tipo es `0`, lo considera error (`:85-88`).
3. Recupera el **objeto** seleccionado (`:90-93`).
4. Ramifica en cascada de `If` sobre el tipo, y en cada rama deja el subformulario correspondiente:

| Tipo de nodo | Subformulario | Evidencia |
|---|---|---|
| Nodo contenedor o explicable (7 tipos) | Explicación | `:94-108`, `:479` |
| Edición a punto de caducar, preparada o caducada | Detalle de edición | `:111-121`, `:50` |
| Riesgo pendiente de retipificar | Retipificación | `:124-132`, `:498` |
| Riesgo aceptado o retirado por visar | Aceptados y retirados | `:134-143`, `:516` |
| Riesgo materializado por decidir | Materializados por decidir | `:145-171`, `:534` |

5. El último caso es el único que **descompone la clave del nodo**: la separa por `|`, toma el identificador del materializado, lo carga por el constructor (`:146-152`), lo deja en una **variable global compartida** (`:153-156`) y solo entonces cambia el subformulario (`:157-161`).
6. Cada rama termina con `DoEvents` y `Hourglass False`, y las siete ramas repiten ese bloque.

### Alternativo (`ComandoActualizarContador_Click`, `:187-237`)

Recoge el responsable del combo y **reutiliza la variable de sesión** `m_ObjUsuarioParaTareas`:

- Si ya existe y el combo está vacío, la anula (`:195-197`).
- Si ya existe y el nombre ha cambiado, recarga el usuario (`:198-204`).
- Si no existe y hay nombre, lo crea (`:206-212`).

Después llama a `EstablecerTareasCalidad` con el primer parámetro forzado a `Sí` y el segundo a `No` (`:214`), recalcula el contador y recarga el árbol (`:218-220`). El combo llama a este procedimiento directamente desde su `AfterUpdate` (`:276`), así que el filtro por responsable y la recarga del árbol son la misma operación.

### Error (`:176-185`)

Los cinco manejadores repiten el mismo bloque: `Err.Number <> 1000` compone un mensaje con el número y el detalle, **envía correo al administrador** y muestra un `MsgBox` crítico; si es exactamente `1000`, muestra un `MsgBox` de advertencia con `m_Error` y **no notifica**.

## Reglas, validaciones y cálculos

- **La pertenencia al perfil se comprueba a medias.** Solo se bloquea al técnico (`:305`). El resto de roles, incluidos los que no son de Calidad, entran sin comprobación; `EsCalidad` (`:314`) únicamente decide si se preselecciona el combo.
- **El tipo de nodo es la única clave de reparto.** No hay `Select Case`: son siete `If` encadenados con `Or`, evaluados en orden.
- **`Err.Raise 1000` es el canal de error interno.** Se usa en dieciséis puntos del archivo; el manejador discrimina por el número (`176`, y los equivalentes en los otros cuatro).
- **El reparto del ancho se calcula al iniciar el arrastre**, no en cada movimiento (`:372-373`), y el botón izquierdo no se distingue de otros por `Shift` ni `Ctrl`.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Árbol de tareas | `ArbolTareasCalidad.CargarArbolTareasCalidad` | `:352-355` |
| Tipo y objeto del nodo | `TipoNodoTareaSeleccionado` y `ObjetoSeleccionado` | `:81`, `:90` |
| Responsable | `Constructor.getUsuario` | `:200`, `:208` |
| Recuento de tareas | `EstablecerTareasCalidad` | `:214` |
| Riesgo materializado | `Constructor.getRiesgoMaterializado` | `:150` |
| Estado global del materializado | `m_ObjRiesgoMaterializadoActivo` | `:155`, `:158`; `Variables Globales.bas:93` |

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí se ve el patrón que más complica una traducción web: **el árbol es un control ActiveX de Windows** (`.form.txt:544`), cargado con una `ImageList` propia (`:349-350`), y su contenido se construye en código a partir de un objeto de dominio (`ArbolTareasCalidad`). Además, el estado de la selección viaja por **miembros públicos de la instancia del formulario** y por una **global compartida**, no por un identificador: dos pestañas o dos sesiones de navegador no podrían reproducir esta semántica. La paleta de iconos tampoco tiene equivalente directo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-139.** El control de rol solo expulsa al técnico (`:305-308`). Un usuario que no sea de Calidad abre el formulario sin mensaje y ve el árbol de tareas de calidad; `EsCalidad` (`:314`) no se usa como permiso. Conviene confirmar si el árbol llega vacío para esos usuarios o si hay que tratarlo como un fallo de autorización real.
- **H-GR-140.** El combo de responsables trae una **lista literal de nombres de personas** como `RowSource` de tipo lista de valores (`.form.txt:1562`), con `AllowValueListEdits =0` (`.form.txt:1568`). No se reproduce aquí. Contradice `EstablecerComboUsuariosCalidad` (`:310`), que debería poblarlo en ejecución: si esa llamada falla, la lista de diseño queda como último valor y expone personal en el diseño del formulario. Es además un dato que hay que sanear antes de cualquier entrega.
- **H-GR-141.** `pregunta = MsgBox(...)` aparece en los cinco manejadores de error y `pregunta` es una **variable pública compartida** de tipo `Long` (`Variables Globales.bas:69`), no una local. El valor devuelto por los `MsgBox` se deposita en un global que cualquier otro formulario puede leer o pisar entre el aviso y su uso.
- **H-GR-142.** `Err.Raise 1000` se lanza sin código fuente ni descripción (`:83`, `:87`, `:92`, `:103`, `:116`, `:127`, `:138`, `:152`, `:163`) y el manejador discrimina por `Err.Number <> 1000` (`:177`). Un error real cuyo número coincida con `1000` se trataría como advertencia interna y **se mostraría sin notificar por correo**. Es el mismo patrón que ya se ha marcado en otros formularios del lote, pero aquí se repite dieciséis veces.
- **H-GR-143.** `Arbol_NodeClick` declara una variable local `m_TipoNodoSeleccionado` (`:61`) que **oculta** el miembro público declarado en `:29`. El miembro público nunca se actualiza: el estado real de la selección solo existe dentro del procedimiento. Cualquier lector externo ve siempre el valor inicial.
- **H-GR-144.** `SeleccionarNodoRiesgo` (`:418-469`) reconstruye la clave del nodo a partir del estado del riesgo (`428-447`) y, si no la encuentra, devuelve dos mensajes genéricos, «Nodo no encontrado» y «No se encuentra el riesgo» (`:453`, `:459`), sin identificar el riesgo. El acceso al nodo va envuelto en `On Error Resume Next` (`:454`) con `On Error GoTo errores` inmediatamente después (`:461`).
- **H-GR-145.** `SeleccionarNodoRiesgo` es `Public` y **no tiene ningún llamante en `src/`**: es superficie de API pública sin uso en el código exportado. Conviene confirmar si lo consume alguna macro o si es residuo.
- **H-GR-146.** El riesgo materializado seleccionado se guarda en la **global `m_ObjRiesgoMaterializadoActivo`** (`:155`, `:158`), no en el formulario, y no se limpia al recargar el árbol (`:344-355`) ni al seleccionar otro nodo. Si el `getRiesgoMaterializado` devuelve `Nothing` (`:152-157`), el subformulario se abre igualmente con el objeto de la selección anterior.
- **H-GR-147.** Hay dos mecanismos de construcción del árbol: uno perezoso en `Arbol_NodeClick` (`:73-77`) y otro explícito en `CargarArbol` (`:349`), que crea un `ArbolTareasCalidad` nuevo cada recarga. El objeto anterior queda sin liberar y el árbol se reconstruye entero en cada cambio de responsable.
- **H-GR-148.** El arrastre del separador se apoya en tres manejadores con `On Error Resume Next` (`:368`, `:381`, `:405`), y el segundo lo repite anidado (`:384`). El ancho total se fija al pulsar (`:372-373`) y no se recalcula, así que un redimensionado de la ventana durante el arrastre deja el árbol y el detalle descuadrados. Además `Sepleft` se declara (`:22`) y no se usa en ningún punto.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el árbol ActiveX por una vista de árbol renderizada en servidor, conservando la misma clave de tipo de nodo como contrato de reparto; pasar la selección por un identificador en la ruta en lugar de por miembros públicos y una global; un desplegable de responsable poblado desde el directorio de usuarios en vez de una lista literal; un filtro de rol evaluado en servidor que devuelva el árbol vacío en vez de avisar en cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario de Calidad ve el árbol de sus tareas y el detalle vacío hasta que elige un nodo.
2. Cada tipo de nodo abre su subformulario correspondiente.
3. Elegir otro responsable recalcula el recuento y recarga el árbol sin dejar el detalle de la selección anterior.
4. Un usuario técnico recibe el aviso y no accede al árbol.

## Pendientes

- Qué calcula exactamente `EstablecerTareasCalidad` (`Funciones Generales.bas:2054`) y qué papel tienen sus dos parámetros, forizados a `Sí` y `No`.
- Qué nodos construye `ArbolTareasCalidad` y si la clave `tipo identificador` separada por `|` es un contrato estable o interno.
- Si `FormularioAbierto("FormCalidadTareas")` (`Funciones Generales.bas:2139`) es el único mecanismo de refresco externo, y desde dónde se llama.
- Dónde se implementa `EstablecerComboUsuariosCalidad`, que no aparece en `src/` con esa firma, y por qué la lista literal de diseño la contradice.