# Form_FormAuditoriaDocumentos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormAuditoriaDocumentos.cls` (580 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar para gestionar los documentos de una **auditoría** (nivel auditoría, no de NC ni de tarea). Tiene dos modos según el contexto: **modo memoria** cuando la auditoría aún no existe (los documentos se acumulan en la colección global `m_ObjColDocumentosAuditoriaParaAlta` hasta que [Form_FormAuditoria](Form_FormAuditoria.md) graba) y **modo directo** cuando ya existe (cada operación escribe en base de datos y disco). Es **no enlazado** (lista «Value List», `.form.txt:672-673`). Las citas son `Form_FormAuditoriaDocumentos.cls:línea` salvo indicación; `Ops` abrevia `classes/DocumentoAuditoriaOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormAuditoria | Botón «Adjuntar documentación»: cierra la instancia previa y abre el formulario (con auditoría vacía si es alta) | `Form_FormAuditoria.cls:20-38` |
| Form_FormAuditoriasGestion | Con la auditoría seleccionada asignada a `m_ObjAuditoriaActiva`, abre el formulario y escucha sus eventos | `Form_FormAuditoriasGestion.cls:105-126, 673, 696` |

Declara los eventos `Alta`, `CambioNombre` y `Eliminado` (`:17-19`); la gestión escucha `Alta` y `Eliminado`, no `CambioNombre`. Desde `Form_FormAuditoria`, tras registrar una auditoría nueva, se cierra esta ventana (`Form_FormAuditoria.cls:88-90`). «Salir» cierra con `acSaveNo` (`:66-69`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:415-552`), 3 secciones (`:573, 605, 985`) y los 11 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1` («DOCUMENTACIÓN DE AUDITORÍAS») | etiqueta | — | Título | `.form.txt:591-592` |
| `Etiqueta62` («Título») | etiqueta | — | Rótulo de `Documento` | `.form.txt:859-860` |
| `Documento` | cuadro de texto | — | Título (alta y renombrado) | `:232-237, 505, 339` |
| `URLDocumento` | cuadro de texto bloqueado | — | Muestra `..\` + nombre del fichero elegido | `:189`, `.form.txt:818-831` |
| `btnExaminar` (rótulo «Comando47») | botón | `OnClick` | Elige fichero | `:164-209`, `.form.txt:761-763` |
| `btnAceptar` | botón | `OnClick` | Guarda el documento (memoria o base de datos) | `:215-313`, `.form.txt:956-958` |
| `ListaDocumentos` | lista de valores | `OnClick`, `OnDblClick` | `IDDocumento;Documento;Archivo`; copia el título a `Documento` al seleccionar | `:490-529, 575-578`, `.form.txt:671-677` |
| `ImagenAnexo` | imagen, oculta de inicio | `OnClick` | Abre el anexo | `:531-573`, `.form.txt:698-706` |
| `ComandoCambiarNombre` | botón, oculto de inicio | `OnClick` | Renombra con el título del cuadro | `:319-410`, `.form.txt:908-921` |
| `ComandoEliminar` | botón, deshabilitado de inicio | `OnClick` | Elimina | `:416-484`, `.form.txt:618-631` |
| `cmdSalir` | botón | `OnClick` | Cierra | `.form.txt:1002-1004` |

## Flujos

### Apertura (`Form_Load`, `:25-64`)
Exige `m_ObjAuditoriaActiva` («abierto sin el contexto de Auditoría», `:39-42`). Si su `IDAuditoria` está vacío ⇒ `m_ParaAlta = Sí` (modo memoria); si no ⇒ `No` (`:45-49`). `EstablecerDatos` habilita `Documento`, `btnExaminar` y `btnAceptar`, oculta cambiar nombre y la imagen (`:75-96`). `EstablecerLista` (`:98-148`) vacía los campos y rellena la lista: en memoria, filas `0;título;nombre de fichero` desde la colección global (`:120-128`); en directo, desde `m_ObjAuditoriaActiva.Documentos` (`:131-141`). A diferencia de los formularios de AR, no selecciona la primera fila.

### Alta (Examinar → Aceptar)
1. `btnExaminar`: rechaza técnicos (`:171-174`), abre el selector (`:177`), valida existencia (`:180-183`), recuerda la ruta (`:186`), propone el nombre base como título si está vacío (`:192-194`) y deselecciona la lista (`:198-199`).
2. `btnAceptar`: rechaza técnicos (`:227-230`); exige título (`:234-237`) y fichero válido (`:240-243`).
   - **Memoria:** rechaza título repetido en la colección (sin distinguir mayúsculas, `:249-258`) y fichero ya añadido (`:261-266`); añade `título → ruta local` (`:269`) y recarga (`:271-272`).
   - **Directo:** crea `DocumentoAuditoria` con `IDAuditoria`, llama a `Registrar` (`:278-289`; reglas en Form_FormARAuditoriaDocumentos y `Ops:18-99`), lanza `Alta`, invalida la caché y recarga (`:291-297`).

### Cambiar título
Rechaza técnicos y título vacío (`:333-343`). En memoria: sale si no hay fila o no hay cambio, rechaza título existente y reconstruye la entrada con el nuevo nombre (`:346-364`). En directo: exige selección, sale si es igual, comprueba duplicados en la lista sin distinguir mayúsculas (`:369-388`), llama a `CambiarNombre`, lanza `CambioNombre` y recarga (`:391-400`).

### Eliminar y abrir
- Eliminar en memoria: quita la entrada, anula la colección si queda vacía y recarga, **sin confirmación** (`:429-441`). En directo: confirmación con «No» por defecto (`:454-456`), `Ops.Eliminar`, lanza `Eliminado`, invalida caché y recarga (`:461-470`).
- Abrir: en memoria con `ShellExecute` («open») sobre la ruta local, con error si el fichero ya no existe (`:539-552`); en directo con `AbrirEnLocal` sobre `URLAnexo` (`:561`).

### Errores
El error 1000 se muestra con `vbExclamation`; el resto con `vbCritical`. Solo `btnAceptar` (`:308`) y `ComandoEliminar` (`:479`) llaman a `CorreoAlAdministrador`. En `Form_Load` el error no cierra el formulario (`:57-63`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no pueden examinar, aceptar, renombrar ni eliminar (`:171, 227, 333, 423`).
- **Título único:** en memoria por clave de diccionario (`:251, 255`); en directo, por lista (`:382-386`) y por `MotivoNoOK`/`DocumentoRepetido` en `Registrar` (`Ops:114-139`).
- **Fichero único en memoria:** no se añade dos veces la misma ruta (`:261-266`).
- **Persistencia diferida:** en memoria nada se escribe hasta que `Form_FormAuditoria` ejecuta `AnexoMultiple` (`Form_FormAuditoria.cls:82-91`; `classes/AuditoriaOperaciones.cls:189-219`).
- **Ruta del anexo:** carpeta de la auditoría (`classes/DocumentoAuditoria.cls:352-361`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista en modo directo | `TbDocumentosAuditorias` (vía `m_ObjAuditoriaActiva.Documentos`) | Pendiente de detalle | `:131` |
| Leer un documento | `TbDocumentosAuditorias` | SQL dinámico concatenado | `modules/constructor.bas:586-596` |
| Alta (directo) | `TbDocumentosAuditorias` | Recordset DAO `AddNew/Update` | `Ops:64-84` |
| Cambio de título (directo) | `TbDocumentosAuditorias` | SQL dinámico concatenado | `Ops:238-240` |
| Baja (directo) | `TbDocumentosAuditorias` | SQL dinámico concatenado | `Ops:186-188` |
| Ficheros | Sistema de ficheros | `CopyFile`, `DeleteFile` | `Ops:91, 192` |
| Modo memoria | Ninguna | Colección global en proceso | `:122-125, 269` |

Sin transacción en modo directo (fila y fichero por separado).

## Coexistencia

- Mismo SQL concatenado y `DELETE *` de Form_FormARAuditoriaDocumentos: parametrizar y adaptar el dialecto.
- El modo memoria depende de estado global de proceso (`m_ObjColDocumentosAuditoriaParaAlta`, `modules/Variables Globales.bas:131`) y rutas locales del equipo del usuario; en la web equivale a documentos en borrador asociados a la sesión o a la operación de alta.
- Ficheros en una ruta de red configurada; no migran con la base.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual.** El mismo formulario sirve a dos mecánicas distintas; el modo se fija una sola vez al abrir (`:45-49`).

**Posibles defectos (a verificar).**
- En modo memoria, `ListaDocumentos_Click` no define `m_DocumentoSeleccionado` (`:517-520`) y `ComandoEliminar` usa en directo la comprobación de la línea 445, que puede dejar `m_DocumentoSeleccionado` obsoleto si se cambia de fila.
- Tras `Eliminar` en directo, no se comprueba el error de `getDocumentoAuditoria` en `ComandoEliminar_Click` (`:448`).
- `ImagenAnexo_Click` en memoria incluye la ruta local completa en el mensaje de error (`:547`).
- La colección global solo se vacía al registrar con éxito (`Form_FormAuditoria.cls:82-91`) y al eliminar la última entrada (`:435-437`); no se ve limpieza al salir.
- Renombrar en memoria con título vacío en la lista y selección sin fila sale en silencio (`:347-349`).

**Propuestas web (pendientes de decisión).** Un único componente de documentos para auditoría, NC y tareas; borrador de documentos en servidor en lugar de rutas locales; confirmación también al eliminar en memoria.

## Criterios de aceptación

1. Con una auditoría sin guardar, un documento aceptado aparece en la lista pero no se escribe nada hasta grabar la auditoría (`:246-272`).
2. Con una auditoría existente, aceptar crea la fila y copia el fichero y refresca la lista (`:278-297`).
3. En ambos modos, título vacío, fichero no elegido, título duplicado o usuario técnico se rechazan sin escribir (`:227-258`).
4. En memoria, el mismo fichero no puede añadirse dos veces (`:261-266`).
5. Renombrar a un título existente o vacío se rechaza (`:340-343, 353-356, 382-386`).
6. Eliminar en modo directo pide confirmación y borra fila y fichero (`:454-463`).
7. Cambiar nombre solo es visible y eliminar solo está activo con una fila seleccionada (`:495-514`).

## Pendientes

- `Ejecutar` (`ShellExecute`, `:551`): declaración en `modules/Variables Globales.bas:17, 45`.
- `Seleccionar`, `AbrirEnLocal`, `Documentos` de `Auditoria` (`:131`): no detallados.
- Existencia de `TbDocumentosAuditorias` en el esquema de main: el clon no incluye tablas; sujeto al ERD de origen.
