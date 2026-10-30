# Form_FormVideosGestionArbolAplicacion

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosGestionArbolAplicacion.cls` (417 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Panel de alta de un vídeo de ayuda para la aplicación seleccionada en el árbol: elige el fichero local, recoge título, descripción, observaciones, uso y destinatarios, y registra el vídeo copiando el fichero al servidor. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); escribe a través de `Video.Registrar`. Las citas son `Form_FormVideosGestionArbolAplicacion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormVideosGestion](Form_FormVideosGestion.md) | Al pulsar un nodo de aplicación: `FormDetalle.SourceObject = "FormVideosGestionArbolAplicacion"` y `CargarDatos` | `Form_FormVideosGestion.cls:55-57` |

Tras registrar, `CrearNodo` añade el nodo del vídeo y llama a `Form_FormVideosGestion.Arbol_NodeClick`, que sustituye este panel por la edición del vídeo nuevo (`:386`; `Funciones Genreales.bas:2728-2732`). Lee las globales `m_ObjAplicacionActiva` y `m_ObjVideoActivo` (`Variables Globales.bas:101-102`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblNombreAplicacion`, `ImagenAplicacion` | etiqueta, imagen | — | Nombre en mayúsculas e icono de la aplicación activa | `:25-28` |
| `lblOperacion` | etiqueta | — | «ALTA DE VÍDEO» o «EDICIÓN DE DATOS DE VÍDEO» | `:31, 49` |
| `ComandoAnadirVideo` (rótulo de diseño «Comando68») | botón | `OnClick` | Diálogo de selección de fichero; escribe la ruta en `URLLocalVideo` | `:250-273`, `.form.txt:918-920` |
| `URLLocalVideo` | texto bloqueado | — | Ruta local del vídeo elegido | `:259, 262`, `.form.txt:317, 329` |
| `Titulo`, `Descripcion`, `Observaciones` | textos | — | Datos del vídeo | `:36-40, 346-348` |
| `ComboUso` | combo `Vídeo-Guía;Encuesta;Ambos`, `LimitToList` | — | Se traduce a `ParaAyuda` y `ParaEncuesta` | `:181-213`, `.form.txt:1270-1272` |
| `ComboDirigidoA` | combo `Para Todos;No Para Técnicos;No Para Calidad`, `LimitToList` | — | Se traduce a `ParaTecnicos`, `ParaCalidad`, `ParaAdministrador` | `:214-249`, `.form.txt:1204-1206` |
| `ComandoRegistrar` | botón | `OnClick` | Valida cambios y registra el vídeo | `:325-410`, `.form.txt:357-359` |
| `ComandoEliminar` | botón, oculto en diseño | `OnClick` | Borra el vídeo activo; oculto en el alta | `:275-323, 32`, `.form.txt:978, 990-992` |

## Flujos

### Apertura (`CargarDatos`, `:16-107`; `Form_Load`, `:414-417`)

1. `Form_Load` toma el árbol del contenedor (`Me.Parent.Controls("Arbol")`, `:415`).
2. Sin aplicación activa, error «No se ha seleccionado ningún nodo» (`:21-24`).
3. Sin vídeo activo (caso normal desde el árbol): rótulo de alta, oculta `ComandoEliminar`, vacía los campos y propone `Para Todos` y `Vídeo-Guía` (`:30-43`).
4. Con vídeo activo: relee el vídeo inicial con `Constructor.getVideo` y carga sus datos en modo edición (`:45-98`).

### Alta (`ComandoRegistrar_Click`, `:325-410`)

1. `HaHabidoCambios` traduce los combos; sin vídeo inicial siempre hay cambios (`:113-124`). Si no los hay, «Ningún cambio que grabar» (`:333-339`).
2. Crea un `Video` nuevo con la ruta local (solo en alta), los textos, las marcas de uso y destino y el `IDAplicacion` activo, y llama a `Registrar` (`:340-363`).
3. Relee el vídeo y la aplicación (`:373-380`) y crea el nodo en el árbol (`:384-389`); en edición solo cambia el texto del nodo si cambió el título (`:391-393`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`:264-272, 314-322, 401-409`). Los mensajes de validación de `Registrar` llegan por la misma vía.

## Reglas, validaciones y cálculos

- **Uso:** `Vídeo-Guía` ⇒ ayuda; `Encuesta` ⇒ encuesta; `Ambos` ⇒ ambas; vacío ⇒ ninguna (`:185-205`).
- **Destino:** `Para Todos` o vacío ⇒ los tres «Sí»; `No Para Técnicos` ⇒ técnicos «No»; `No Para Calidad` ⇒ técnicos y Calidad «No». Administrador siempre «Sí» (`:218-242`).
- **Validaciones de `Video.Registrar` en el alta** (`Video.cls`): el fichero local existe (`:436-439`); nombre de fichero no repetido en la aplicación (`:440-448`); no existe ya en el servidor (`:449-457`); `IDAplicacion` y título obligatorios (`:460-467`); título no repetido en la aplicación (`:468-475`); al menos ayuda o encuesta (`:484-487`); descripción obligatoria (`:488-491`).
- **Identificadores:** `IDVideo` e `IDCategoriaVideo` = máximo + 1 con `DameID` (`Video.cls:519, 573`; `Funciones Genreales.bas:473-486`).
- **Autoría:** `SubidoPor` = usuario de red de la sesión y `FechaSubido = Now()` (`Video.cls:534-535`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Duplicados de nombre y título | `TbVideos` | SQL dinámico concatenado | no | `Video.cls:391-402` |
| Alta del vídeo | `TbVideos` | DAO `AddNew` | no | `Video.cls:509, 515-539` |
| Copia del fichero | carpeta de vídeos de la aplicación en el servidor | `fso.CopyFile` | no | `Video.cls:561-564`; `Aplicacion.cls:1097` |
| Categorías | `TbVideosCategorias` (borrado y alta) | SQL dinámico + DAO `AddNew` | no | `Video.cls:565-596` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: IDs por `MAX + 1` en VBA, sin transacción entre `TbVideos`, la copia del fichero y `TbVideosCategorias`; literales «Sí»/«No» como texto; el fichero vive en una carpeta de red, fuera de la base.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Registrar` llama a `ParaTodosCalculado`, que se invoca a sí mismo en su primera condición (`Video.cls:352, 492`): recursión hasta agotar la pila. La línea siguiente compara `ParaCalidadCalculado` dos veces y omite `ParaAdministradorCalculado` (`Video.cls:357`).
- Alta sin transacción: si falla la copia o las categorías, queda la fila en `TbVideos` sin categoría, invisible en el árbol (`Video.cls:539, 564-596`; filtro en `Constructor.bas:1396`).
- `DameID` (máximo + 1) no protege de altas concurrentes (`Funciones Genreales.bas:473-486`).
- Un título con apóstrofo rompe la consulta de duplicados (`Video.cls:391, 396`).
- `Seleccionar` muestra el diálogo antes de fijar título, carpeta inicial y filtros (`Funciones Genreales.bas:2644-2651`); al cancelar, guarda una ruta vacía como último archivo (`Funciones Genreales.bas:2656-2657`).
- El alta continúa tras `CrearNodo`, que ya ha sustituido este subformulario (`:386`; `Funciones Genreales.bas:2732`).
- Código de edición y borrado (`:45-98, 275-323`) inalcanzable desde el árbol, que solo abre este panel sin vídeo activo (`Form_FormVideosGestion.cls:55-56`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta transaccional con almacenamiento de objetos para el fichero; identificadores generados por la base; destinatarios como casillas independientes en lugar de tres opciones fijas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Registrar sin título, sin descripción o sin fichero muestra su mensaje y no crea nada.
2. Un nombre de fichero o un título ya usado en la misma aplicación se rechaza.
3. Un alta correcta crea el vídeo con sus categorías, copia el fichero y lo muestra seleccionado en el árbol.
4. «No Para Calidad» deja el vídeo visible solo para administradores.

## Pendientes

- Si la recursión de `ParaTodosCalculado` llega a ejecutarse o el binario difiere del fuente (`Video.cls:352`).
- Contenido de `URLDirVideos` y de `getCategorias` (`Aplicacion.cls:1090`; `Video.cls:82`).
- Uso de `m_DatosCargados` (`:4`): declarada y sin efecto visible.
