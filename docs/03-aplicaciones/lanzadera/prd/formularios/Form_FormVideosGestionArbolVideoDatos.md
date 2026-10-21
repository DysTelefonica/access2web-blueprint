# Form_FormVideosGestionArbolVideoDatos

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosGestionArbolVideoDatos.cls` (408 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Datos» de un vídeo: edita título, descripción, observaciones, uso y destinatarios, y permite eliminar el vídeo. Comparte diseño y casi todo el código con [Form_FormVideosGestionArbolAplicacion](Form_FormVideosGestionArbolAplicacion.md). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); escribe a través de `Video.Registrar` y `Video.Eliminar`. Las citas son `Form_FormVideosGestionArbolVideoDatos.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormVideosGestionArbolVideo](Form_FormVideosGestionArbolVideo.md) | `MENU1_Click` fija `FormDetalle.SourceObject = "FormVideosGestionArbolVideoDatos"` y llama a `CargarDatos` | `Form_FormVideosGestionArbolVideo.cls:18-21` |

Tras eliminar, llama a `Form_FormVideosGestion.ComandoActualizarContador_Click` para recargar el árbol (`:159`). Lee las globales `m_ObjAplicacionActiva` y `m_ObjVideoActivo` (`Variables Globales.bas:101-102`), que fija el árbol (`Form_FormVideosGestion.cls:32-37`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblNombreAplicacion`, `ImagenAplicacion` | etiqueta, imagen | — | Nombre en mayúsculas e icono de la aplicación | `:26-29` |
| `lblOperacion` | etiqueta | — | «EDICIÓN DE DATOS DE VÍDEO» (o «ALTA DE VÍDEO» sin vídeo activo) | `:32, 50` |
| `Titulo`, `Descripcion`, `Observaciones` | textos | — | Cargados desde el vídeo activo | `:57, 69-78` |
| `ComboUso` | combo `Vídeo-Guía;Encuesta;Ambos`, `LimitToList` | — | Uso del vídeo | `:58-68, 269-301`, `.form.txt:1268-1270` |
| `ComboDirigidoA` | combo `Para Todos;No Para Técnicos;No Para Calidad`, `LimitToList` | — | Destinatarios | `:79-91, 302-337`, `.form.txt:1202-1204` |
| `ComandoAnadirVideo` (rótulo de diseño «Comando68») | botón | `OnClick` | Elige fichero; deshabilitado en edición | `:51, 108-131`, `.form.txt:917-919` |
| `URLLocalVideo` | texto bloqueado | — | Ruta local; solo se usa en alta | `:195`, `.form.txt:317, 329` |
| `ComandoRegistrar` | botón | `OnClick` | Guarda los cambios | `:177-262`, `.form.txt:357-359` |
| `ComandoEliminar` | botón | `OnClick` | Confirma y elimina el vídeo | `:52-53, 133-175`, `.form.txt:989-991` |

## Flujos

### Apertura (`CargarDatos`, `:17-103`; `Form_Load`, `:264-267`)

1. `Form_Load` toma el árbol del formulario abuelo (`Me.Parent.Parent.Controls("Arbol")`, `:265`).
2. Sin aplicación activa, error «No se ha seleccionado ningún nodo» (`:22-25`).
3. Con vídeo activo, guarda una copia releída en `m_ObjVideoAlInicio` (`:46-49`), habilita `ComandoEliminar` y `ComandoRegistrar` y deshabilita `ComandoAnadirVideo` (`:50-54`).
4. Traduce las categorías a `ComboUso` y los tres destinatarios a `ComboDirigidoA` (`:58-91`).

### Edición (`ComandoRegistrar_Click`, `:177-262`)

1. `HaHabidoCambios` compara título, descripción, observaciones, uso y destinatarios con `m_ObjVideoAlInicio` (`:338-402`); sin cambios, «Ningún cambio que grabar» (`:185-191`).
2. Construye un `Video` nuevo y llama a `Registrar` con el vídeo inicial (`:192-215`): `Edit` de `TbVideos` (`Video.cls:541-553`) y, si cambió el uso, borrado y alta de categorías (`Video.cls:598-631`).
3. Relee vídeo y aplicación (`:225-232`) y, si cambió el título, actualiza el texto del nodo seleccionado (`:242-245`).

### Eliminación (`ComandoEliminar_Click`, `:133-175`)

1. Pide confirmación «Desea eliminar el vídeo seleccionado» con «No» por defecto (`:145-148`).
2. `Video.Eliminar`: si el fichero está abierto, error; borra categorías, la fila y el fichero (`Video.cls:658-676`).
3. Recarga el árbol con `ComandoActualizarContador_Click` (`:159`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`:122-130, 166-174, 253-261`). `CargarDatos` solo devuelve el texto en `p_Error` (`:99-102`).

## Reglas, validaciones y cálculos

- **Uso y destino:** misma traducción que el alta (`:273-293, 306-330`); administrador siempre «Sí».
- **Carga de destino:** los tres «Sí» ⇒ `Para Todos`; técnicos «No» y Calidad «Sí» ⇒ `No Para Técnicos`; técnicos y Calidad «No» ⇒ `No Para Calidad`; cualquier otra combinación ⇒ `Para Todos` (`:79-90`).
- **Validaciones de `Video.Registrar` en edición:** título obligatorio y no repetido en la aplicación salvo el propio vídeo; ayuda o encuesta; descripción obligatoria (`Video.cls:464-491, 390-393`). El fichero y la aplicación no cambian (`Video.cls:435, 541`).
- **Observaciones vacías** se guardan como `Null` (`Video.cls:545-549`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Vídeo inicial | `TbVideos` | SQL dinámico concatenado | no | `:46`; `Constructor.bas:1259-1261` |
| Edición | `TbVideos` | DAO `Edit` | no | `Video.cls:511-553` |
| Categorías | `TbVideosCategorias` | SQL dinámico + DAO `AddNew` | no | `Video.cls:598-631` |
| Eliminación | `TbVideosCategorias`, `TbVideos`, fichero en servidor | SQL dinámico + `fso.DeleteFile` | no | `Video.cls:666-676` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE tabla.* FROM` de Jet/ACE (`Video.cls:666-672`), sin transacción entre las dos tablas y el fichero; la detección de fichero abierto depende de bloqueos del sistema de ficheros.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `FicheroAbierto` salta al manejador si el fichero está bloqueado y devuelve `False` (`Funciones Genreales.bas:124-137`): `Eliminar` borra las filas y falla después al borrar el fichero (`Video.cls:660, 666-675`).
- Tras eliminar no se anula `m_ObjVideoActivo` ni se vacía el panel (`:155-159`), y el árbol se recarga desde cachés no invalidadas (`Aplicacion.cls:1238-1240`).
- Un vídeo con técnicos «Sí» y Calidad «No» se carga como `Para Todos` (`:88-89`); cualquier guardado lo convierte en «para todos».
- Si el vídeo no tiene categoría, `ComboUso` conserva el valor anterior (`:58-68`).
- `ParaTodosCalculado` se llama a sí mismo en cada guardado (`Video.cls:352, 492`).
- `m_ObjVideoAlInicio` no se refresca tras guardar (`:225`): un segundo guardado compara con el estado original.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: borrado transaccional con borrado diferido del fichero; destinatarios como casillas independientes.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar sin cambios devuelve «Ningún cambio que grabar» y no escribe.
2. Cambiar el título actualiza el nodo del árbol.
3. Cambiar el uso de «Vídeo-Guía» a «Ambos» deja las dos categorías.
4. Eliminar pide confirmación; al aceptar, el vídeo desaparece de la base, del servidor y del árbol.
5. Un vídeo con fichero en uso no se elimina y no pierde sus datos.

## Pendientes

- Comportamiento real de `FicheroAbierto` con `On Error GoTo` (`Funciones Genreales.bas:124-128`).
- Si `Form.Parent.Parent` (`:234`) resuelve el árbol cuando `m_Arbol` está vacío.
- Uso de `m_DatosCargados` y `m_IDVideo` (`:4, 154`): sin efecto visible.
