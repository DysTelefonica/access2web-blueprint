# Form_FormCalidadTareaExplicacion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormCalidadTareaExplicacion.cls` (70 líneas) y `src/forms/FormCalidadTareaExplicacion.form.txt` · **PRD:** [README](../README.md)

Panel de explicación de un nodo del árbol de tareas. No recibe nada, no escribe nada y no tiene botones: solo dos etiquetas que muestran el título y el texto explicativo guardados para el nodo seleccionado. Todo su comportamiento está en un único evento. Las citas son `Form_FormCalidadTareaExplicacion.cls:línea` salvo indicación.

## Entrada y salida

No se abre con `DoCmd.OpenForm`: es un **subformulario** que el formulario de tareas monta en su control `FormDetalle` (`Form_FormCalidadTareas.cls:471-479`, dentro del método `AbrirExplicacion`).

| Dato | Cómo llega | Evidencia |
|---|---|---|
| Nodo seleccionado | Variable pública del formulario contenedor: `Form_FormCalidadTareas.m_NodoSeleccionado` | `:30`; declarada en `Form_FormCalidadTareas.cls:32`, asignada en `:78` |

El identificador del nodo es lo único que se necesita: con él se busca el texto en `m_ObjEntorno.ColExplicacionesArbolTareas` (`:35-36`). No hay `OpenArgs`, ni `OpenForm`, ni lectura de base de datos.

**Salidas:** ninguna. El formulario no cierra nada, no notifica y no guarda. Su única forma de terminar es el cierre forzado del manejador de error (`:68`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Lee el nodo del contenedor, busca su explicación, ajusta el tamaño del título y fija el rótulo de la ventana | `:16-69` |
| `lblTitulo` | etiqueta | — | Título del apartado; el cuerpo adjusts su tamaño de letra | `:27`, `:40-51`, `.form.txt:891-892` |
| `lblExplicacion` | etiqueta | — | Texto explicativo | `:28`, `:51`, `.form.txt:918-919` |

El diseño del formulario tiene exactamente esos dos controles y ninguna otra superficie interactiva.

## Flujos

### Apertura (`Form_Open`, `:16-69`)

1. Reloj de arena y `m_Error` a vacío (`:22-26`).
2. Deja ambos rótulos con el **texto de relleno** del diseño: «TÍTULO DEL APARTADO» y «EXPLICACIÓN DEL APARTADO» (`:27-28`).
3. Toma el nodo del contenedor; si no hay ninguno, avisa «No se sabe el nodo seleccionado» y sale por error (`:30-34`).
4. Si la colección de explicaciones tiene ese nodo (`:35`), parte el valor guardado por `|` en dos trozos: el título y la explicación (`:36-39`).
5. Ajusta el tamaño de letra del título según su longitud (`:40-48`) y pinta ambos rótulos (`:50-51`).
6. Fija `Me.Caption` con el usuario conectado (`:54`) y quita el reloj de arena (`:55-57`).

### Alternativo: nodo sin explicación

Si la colección no contiene la clave del nodo, **no ocurre nada más**: el formulario queda mostrando los dos textos de relleno del punto 2, sin aviso, sin distinguir «no hay explicación» de «todavía no se ha escrito». Es un camino de uso real, no un error.

### De error (`:59-68`)

Quita el reloj de arena y, según el número de error: si **no** es 1000, compone el mensaje con número y detalle, **envía correo al administrador** y muestra un diálogo crítico (`:62-64`); si es 1000 exacto, muestra solo una advertencia con `m_Error` (`:65-66`). En ambos casos intenta cerrar el formulario (`:68`).

## Reglas, validaciones y cálculos

- **El texto se transporta en una única cadena con `|` como separador** (`:37`), y el código asume exactamente dos trozos. Título y explicación no son campos independientes.
- **El tamaño de la letra es función de la longitud del título**, con tres tramos y umbrales en 50 y 60 caracteres (`:40-48`), sin ninguna constante con nombre ni comentario que explique el criterio.
- **No hay control de permisos**: cualquiera que llegue al panel ve la explicación del nodo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Nodo seleccionado | Variable pública del contenedor | `:30` |
| Explicación del nodo | `m_ObjEntorno.ColExplicacionesArbolTareas`, colección en memoria | `:35-36` |
| Usuario conectado | `m_ObjEntorno.TituloUsuarioConectado` | `:54` |

Sin escrituras: ni en base de datos ni en memoria. El formulario es de lectura pura.

## Coexistencia

Pendiente de `../../forms.md`. Es el caso más atado al modelo de Access de todo este lote: lee una variable pública de la clase del formulario contenedor, que es estado de sesión compartido. En una web el equivalente es un parámetro en la ruta, y el hecho de que hoy no exista obliga a decidir **quién compone esa ruta**.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-105.** El formulario depende del estado interno de su contenedor: lee `Form_FormCalidadTareas.m_NodoSeleccionado` (`:30`), una variable pública de módulo (`Form_FormCalidadTareas.cls:32`). El subformulario se monta vaciando y rellenando el `SourceObject` del control (`Form_FormCalidadTareas.cls:475-479`), pero en ningún momento se le pasa qué nodo es: se supone que el contenedor ya tiene el correcto. Dos montajes seguidos sobre nodos distintos dependen de que la asignación (`Form_FormCalidadTareas.cls:78`) se haya hecho antes.
- **H-GR-106.** Si el nodo no tiene explicación, el panel **muestra los textos de relleno del diseño** («TÍTULO DEL APARTADO», «EXPLICACIÓN DEL APARTADO», `:27-28`) con el mismo aspecto que el contenido real. Quien lo lea no puede distinguir que falta el texto. Es el camino de `If … Exists` que no entra (`:35`).
- **H-GR-107.** El valor guardado se parte por `|` y se leen dos elementos sin comprobar nada (`:37-39`): si el separador falta, `dato(1)` falla; si aparece más de una vez, el resto se pierde. El fallo acaba en correo al administrador y cierre (`:62-68`), es decir, un texto mal formado notifica al soporte técnico como si fuera una incidencia.
- **H-GR-108.** Los umbrales de tamaño de letra (50 y 60 caracteres, `:40-48`) son números sueltos sin constante ni comentario, y el `FontSize` de la etiqueta queda **alterado para siempre** en esa instancia del subformulario: como el subformulario se reutiliza al cambiar de nodo (el `SourceObject` se vacía y se vuelve a rellenar, `Form_FormCalidadTareas.cls:475-479`), el tamaño del título anterior puede heredarse si el nodo nuevo no llega a la rama que lo fija.
- **H-GR-109.** `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:54`) sobrescribe el rótulo del subformulario con el nombre del usuario conectado, que el diseño ya traía escrito (`.form.txt:40`). En un panel embebido, ese rótulo no debería ser el título de nada.
- **H-GR-110.** El manejador de error cierra el formulario con `DoCmd.Close acForm, Me.Name` (`:68`) siendo un **subformulario**, y no captura el error de ese cierre. Hay que verificar en ejecución si el cierre se produce, porque un subformulario no siempre puede cerrarse desde su propio evento por nombre.
- **H-GR-111.** Las dos etiquetas tienen altura fija (`.form.txt:885` y `:914`) y no hay ninguna superficie de desplazamiento: una explicación larga se sale del área visible sin que nada lo indique. El diseño declara además una altura de diseño mayor que la guardada (`.form.txt:927`), señal de que el control se redimensionó al maquetar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar el identificador del nodo por parámetro en lugar de leerlo del contenedor; devolver título y explicación como campos separados en la respuesta en lugar de una cadena con separador; mostrar un estado vacío explícito («este apartado aún no tiene explicación») en lugar del texto de relleno; decidir si el panel se sustituye por una sección desplegable de la propia página de tareas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un nodo seleccionado que tiene explicación, el panel muestra su título y su texto.
2. Con un nodo seleccionado sin explicación, el panel dice que no la tiene, sinparecerse a un contenido real.
3. Con un nodo cuyo título es muy corto, largo o intermedio, el texto se muestra íntegro y legible.
4. Cambiar de nodo actualiza el panel sin arrastre del contenido ni del formato del anterior.

## Pendientes

- Cómo y desde dónde se rellena `m_ObjEntorno.ColExplicacionesArbolTareas`, y con qué formato exacto se serializa título y explicación (H-GR-107).
- Si el texto de un nodo puede llegar a contener el separador `|` (H-GR-107).
- Qué altura tiene el `FormDetalle` del contenedor y si el panel debería ajustarse a ella o al revés (`Form_FormCalidadTareas.cls:479`).
- Si el cierre del subformulario desde su propio manejador de error llega a ejecutarse (H-GR-110).