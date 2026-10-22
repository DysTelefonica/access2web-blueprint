# Form_FormPublicacionCalidadNotas

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionCalidadNotas.cls` (117 líneas) y `src/forms/FormPublicacionCalidadNotas.form.txt` (1134 líneas) · **PRD:** [README](../README.md)

Página de notas para publicación de calidad. Una sola caja de texto y un botón: escribe el texto en el objeto de edición en memoria y pide a la clase `Edicion` que lo registre. Las citas son `Form_FormPublicacionCalidadNotas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormPublicacionCalidad` | Botón de navegación del panel lateral que apunta a `FormPublicacionCalidadPublicar`-style subformulario | `FormPublicacionCalidad.form.txt:624-641` |

Se abre como subformulario dentro del contenedor principal. La edición sobre la que escribe no llega como parámetro: se toma de `m_ObjEdicionActiva` (`Variables Globales.bas:79`), igual que el resto de la cadena.

No abre ningún formulario ni llama a ninguno. Sale solo por su propio cierre en el manejador de error de `Form_Open` (`:88`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Carga el texto existente y habilita o deshabilita el botón de grabar | `:62-90`, `.form.txt:41` |
| `lbl1` | etiqueta | — | Rótulo del bloque: «NOTAS PARA PUBLICACIÓN DE CALIDAD» | `.form.txt:934-935` |
| `NotasCalidadParaPublicar` | cuadro de texto | — | Zona de escritura; el diseño la marca como obligatoria y explica que solo la lee calidad | `:33, 44, 103`, `.form.txt:965-968` |
| `ComandoRegistrar` | botón | `ComandoRegistrar_Click` | Vuelca el texto en la edición y lo registra | `:18-60`, `.form.txt:993-997` |

El diseño conserva controles heredados del formulario original (etiqueta, rectángulo, imagen, grupo de opciones, lista, desplegable y marco de objeto sin enlazar) que no tienen código asociado en esta clase; el botón real es `ComandoRegistrar` (`.form.txt:770-996`).

## Flujos

### Apertura (`Form_Open`, `:62-90`)

1. Poner el reloj de arena y limpiar `m_Error` (`:67-71`).
2. `EstablecerDatos` y, si devuelve error, `Err.Raise 1000` (`:70-73`).
3. El manejador de error **cierra el formulario** (`:88`) sin asignar `Cancel`.

### Normal (`ComandoRegistrar_Click`, `:18-60`)

1. Poner el reloj (`:23-26`).
2. Copia el control al objeto de edición con `Nz(..., "")` (`:33`).
3. Llama a `RegistrarNotasParaCalidad` (`Edicion.cls:3370`) pasando el error por referencia (`:34`).
4. Si la clase devuelve error, `Err.Raise 1000` (`:35-37`).
5. Mira el **control**, no la edición, para elegir el mensaje final: si la caja quedó vacía, «Notas eliminadas correctamente»; si no, «Notas registradas correctamente» (`:44-48`).

El texto no se vuelve a leer de la edición después de registrar: el mensaje depende del estado de la pantalla, no del estado guardado.

### Alternativo: borrar las notas

El mismo botón hace las dos cosas. Con la caja vacía, `RegistrarNotasParaCalidad` recibe una cadena vacía y el mensaje informa de eliminación (`:33, 44-45`). No hay confirmación previa ni mensaje que distinga una operación destructiva.

### De error

`Form_Open` avisa con el patrón común (error inesperado: notificar al administrador y mostrar crítico; error 1000: advertencia) y cierra (`:80-89`). `ComandoRegistrar_Click` avisa con el mismo patrón pero **no cierra** (`:53-59`): el usuario permanece en la pantalla con el texto escrito, lo que deja el formulario en un estado que difiere del guardado.

## Reglas, validaciones y cálculos

- **Solo se puede escribir con la edición activa.** `EstablecerDatos` habilita `ComandoRegistrar` únicamente si `EsActivo = EnumSiNo.Sí` (`:105-109`).
- **No hay validación de longitud ni de obligatoriedad.** La etiqueta `Tag = "OBLIGATORIO"` del diseño (`.form.txt:967`) y el texto de ayuda «Nota que solo lee calidad para ayuda a la publicación» (`.form.txt:968`) no se corresponden con ninguna comprobación: **una caja vacía se acepta y se registra** como eliminación (`:33`).
- **La carga inicial solo copia si hay algo.** `EstablecerDatos` solo escribe en el control cuando la edición ya tiene notas (`:102-104`); si están vacías, el control conserva lo que hubiera, normalmente vacío.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Texto de notas | `m_ObjEdicionActiva.NotasCalidadParaPublicar` (`Edicion.cls:38`) | `:33, 102-103` |
| Registro en base de datos | `Edicion.RegistrarNotasParaCalidad` (`Edicion.cls:3370`) | `:34` |
| Estado de edición | `m_ObjEdicionActiva.EsActivo` | `:105` |
| Estado de edición en memoria | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:33, 99-109` |

Escritura efectiva: una llamada a la clase de edición, que es quien decide qué tabla toca. La vista no escribe por sí misma.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí ya se ve: la pantalla depende de una sesión global y de un subformulario anidado; en web el texto sería un recurso propio de la edición, con guardado explícito y sin simultaneidad accidental entre pestañas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-25.** El diseño marca la caja como obligatoria (`.form.txt:967`) y el código **no comprueba nada**: guardar vacío borra las notas sin confirmación ni aviso previo (`:33, 44-45`). En una pantalla etiquetada como obligatoria, la ausencia de validación es una contradicción entre diseño y comportamiento.
- **H-GR-26.** El mensaje final se decide mirando el control (`:44`) y no la edición tras el registro. Si la clase `Edicion` normaliza o recorta el texto, el mensaje puede no corresponder con lo guardado.
- **H-GR-27.** `EstablecerDatos` no distingue «sin edición activa» de «edición no activa»: si `m_ObjEdicionActiva` está vacía, sale sin error (`:99-101`) y `Form_Open` abre igualmente la pantalla, con el botón deshabilitado por defecto y el rótulo del diseño («Usuario conectado: …»).
- **H-GR-28.** Tras un error de registro, el formulario **no se cierra ni se restaura** (`:53-59`): el texto sigue en pantalla y el usuario puede reintentar. Es un comportamiento distinto del de `Form_Open`, que sí cierra (`:88`); conviene decidir cuál es el correcto.
- **H-GR-29.** `EstablecerDatos` es una `Private Function` que devuelve `As String` y no devuelve nada (`:92-113`), como en el resto de la cadena; el error viaja solo por el segundo argumento.
- **H-GR-30.** El formulario conserva en el diseño controles sin código asociado (grupo de opciones, lista desplegable, marco de objeto sin enlazar: `.form.txt:806-910`). Están inertes en esta pantalla; conviene confirmar que no responden a eventos heredados.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: campo de notas con guardado explícito y contador de caracteres; confirmación al vaciar; validación de obligatoriedad coherente con la señalización del diseño; permisos de lectura y escritura diferenciados en el recurso de notas, ya que hoy el código solo distingue edición activa o no.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla muestra las notas que ya tiene la edición y deja el botón habilitado solo si la edición está activa.
2. Grabar con texto registra las notas y lo confirma.
3. Grabar sin texto elimina las notas y lo confirma, con confirmación previa si se decide.
4. Con la edición no activa, el botón de grabar no está disponible.
5. Un fallo de registro avisa y permite reintentar sin perder lo escrito.

## Pendientes

- Qué tabla o procedimiento usa `RegistrarNotasParaCalidad` dentro de `Edicion.cls:3370` y si admite historial de cambios.
- Si la obligatoriedad anunciada en el diseño se cumple en algún otro punto de la cadena de publicación (**H-GR-25**).
- Para qué se usan los controles inertes del diseño (**H-GR-30**).