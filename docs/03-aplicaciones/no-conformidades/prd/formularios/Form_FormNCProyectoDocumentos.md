# Form_FormNCProyectoDocumentos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoDocumentos.cls` (534 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de anexos de una NC de proyecto: lista los documentos de la NC (o de su NC asociada), permite subir uno nuevo («Examinar» + «Aceptar»), cambiar el título, eliminarlo y abrir el archivo. Es un formulario **no enlazado** (sin `RecordSource`) con una lista de valores; los datos salen del objeto global `m_ObjNCProyectoActiva` (`Variables Globales.bas:117`) y de `DocumentoProyectoOperaciones`. Las citas son `Form_FormNCProyectoDocumentos.cls:línea` (abreviado `:línea`) salvo indicación; `Ops` abrevia `classes/DocumentoProyectoOperaciones.cls`.

## Entrada y salida

Se muestra como detalle o ventana de `FormNCProyecto`; el origen exacto de la apertura no está en este formulario (ver Pendientes). Requiere `m_ObjNCProyectoActiva` (`:34-37`). Declara tres eventos públicos `Alta`, `CambioNombre` y `Eliminado` (`:17-19`) que lanza tras cada operación (`:261`, `:414`, `:486`); en `src` no se ha rastreado quién los escucha.

## Controles y eventos

`.form.txt` declara 19 bloques `Begin`: el formulario (`:5`), 8 de valores por defecto (`:415-552`), la sección `Detalle` (`:573`; no hay encabezado ni pie) y los 9 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaDocumentos` | lista de valores, 3 columnas | `OnClick`, `OnDblClick` | Selecciona documento (carga título, activa botones); doble clic abre el archivo | `.form.txt:623-644`, `:289-332`, `:532-534` |
| `Documento` | cuadro de texto | — | Título del documento (alta o nuevo nombre) | `.form.txt:846-857` |
| `URLDocumento` | cuadro de texto bloqueado | — | Muestra «..\» + nombre del archivo elegido | `.form.txt:784-798`, `:199` |
| `btnExaminar` (rótulo «Comando47») | botón | `OnClick` | Elige archivo local | `.form.txt:721-730`, `:173-217` |
| `btnAceptar` | botón | `OnClick` | Registra el documento nuevo | `.form.txt:913-925`, `:223-283` |
| `ComandoCambiarNombre` | botón | `OnClick` | Renombra el seleccionado | `.form.txt:874-887`, `:338-434` |
| `ComandoEliminar` | botón | `OnClick` | Elimina el seleccionado | `.form.txt:585-598`, `:440-503` |
| `ImagenAnexo` | imagen, oculta por defecto | `OnClick` | Abre el anexo seleccionado | `.form.txt:664-673`, `:509-530` |
| `Etiqueta62` («Título») | etiqueta | — | Rótulo | `.form.txt:826-827` |

## Flujos

### Apertura (`Form_Load`, `:25-63`)

1. Exige NC activa. Trabaja sobre `NCProyectoAsociada` si existe; si no, sobre la propia NC (`:40-44`).
2. `EstablecerDatos` (`:80-108`): permite edición solo si no hay NC asociada y el usuario no es técnico (`:89-98`); `EstablecerLista` (`:110-167`) vuelve a fijar `btnExaminar`/`btnAceptar` según `Form_FormNCProyecto.m_PermitidoEditar` (`:129-140`), vacía `DocumentosCompletos` para forzar la recarga y rellena la lista con `IDDocumento;Documento;nombre de archivo` (`:143-157`). No se autoselecciona nada (`:159-160`).

### Alta (`btnExaminar` → `btnAceptar`)

1. Examinar: técnico ⇒ «No tiene autorización para esa acción» (`:180-183`); selector de archivo (`Seleccionar`, `FG.bas:764`); si no existe, aviso (`:190-193`). Guarda la ruta en memoria, sugiere el título con el nombre base si está vacío y deselecciona la lista (`:196-208`).
2. Aceptar: exige título y archivo elegido y existente (`:238-247`), crea `DocumentoProyecto` con `IDNoConformidad` y título y llama a `Registrar` (`:250-259`). Refresca colecciones y limpia (`:265-270`).

### Cambiar nombre y eliminar

- Cambiar nombre: rol; selección; título no vacío; distinto del actual (sin distinguir mayúsculas) y no repetido en la NC (`:351-398`); `CambiarNombre` (`:406`); refresco (`:411-419`).
- Eliminar: rol y `m_PermitidoEditar` (`:446-454`); selección; confirmación «¿Desea borrar el documento seleccionado?» con «No» por defecto (`:464-466`); `Eliminar`; recarga la NC con `constructor.getNCProyecto` y reasigna `m_NC` (`:470-484`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y aviso crítico (`:54-62`, `:274-282`, `:424-433`, `:494-502`). En Examinar, lista y abrir anexo solo hay `MsgBox` (`:211-216`, `:326-331`, `:527-529`).

## Reglas, validaciones y cálculos

- **Edición:** solo si no es técnico, la NC no tiene asociada y `m_PermitidoEditar` lo autoriza (`:89-98`, `:129-140`, `:314-319`).
- **Título único por NC:** `DocumentoRepetido` (`Ops:84-107`) y, al renombrar, comparación en memoria (`:383-398`).
- **Alta:** `MotivoNoOK` exige archivo alcanzable, título no repetido y que no esté abierto otro archivo homónimo en destino (`Ops:140-190`).
- **Baja:** rechaza si el anexo está abierto (`Ops:28-33`) y si es el único documento de una acción ya cerrada (`Ops:35-41`).
- **Lista:** los `;` del título se sustituyen por `:` (`:155`).

## Lecturas y escrituras

| Operación | Tabla o fichero | Tipo | Evidencia |
|---|---|---|---|
| Lista de documentos | `TbNCDocumentos` vía `NCProyecto.DocumentosCompletos` | no rastreado aquí (propiedad de la clase) | `:143-144` |
| Carga del seleccionado | `TbNCDocumentos` vía `constructor.getDocumentoProyecto` | SQL dinámico concatenado | `:322`, `constructor.bas:491-520` |
| Duplicados | `TbNCDocumentos` | SQL dinámico concatenado (título entre comillas) | `Ops:98-107` |
| Alta | `TbNCDocumentos` (`AddNew/Update`) + copia del archivo a la ubicación del anexo | DAO + `fso.CopyFile`; sin transacción | `Ops:220-256` |
| Renombrar | `TbNCDocumentos` | SQL dinámico concatenado (`UPDATE`) | `Ops:302-304` |
| Baja | `TbNCDocumentos` + borrado del fichero | SQL dinámico (`DELETE`) + `fso.DeleteFile`; sin transacción | `Ops:46-52` |
| Auditoría | `RegistrarLogProyecto` en alta y baja | ver ficha de la clase | `Ops:244-250`, `Ops:54-60` |

## Coexistencia

- Títulos concatenados en SQL con comillas simples: un título con `'` rompe `DocumentoRepetido` y `CambiarNombre`; en PostgreSQL debe parametrizarse (`Ops:98-107`, `Ops:302`).
- `DELETE *` y `getdb().Execute` son dialecto Jet/ACE (`Ops:46`); `IDDocumento` se calcula en cliente (`IDDocumentoCalculado`, `Ops:222`), incompatible sin adaptación con una secuencia.
- Los ficheros viven fuera de la base de datos (copia/borrado en recurso compartido configurado) y no entran en ninguna transacción: alta y baja pueden dejar fila sin fichero o fichero huérfano.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EstablecerLista` sobrescribe el permiso de `btnExaminar`/`btnAceptar` con `m_PermitidoEditar` (`:139-140`), anulando la restricción por técnico de `:96-98`; el bloqueo efectivo para técnicos depende de `:180`, `:233`.
- En alta, el archivo se copia después de insertar la fila y registrar el log (`Ops:225-256`); un fallo de copia deja la fila.
- En baja, la fila se borra antes que el fichero y sin transacción (`Ops:46-52`).
- `Form_Load` no deja el formulario inutilizable ante un error: sigue abierto con controles sin configurar (`:54-62`).
- El título del formulario conserva un texto fijo de diseño con rol (`.form.txt:40`) y el botón se rotula «Comando47» (`.form.txt:729`).
- `ListaDocumentos_Click` se invoca explícitamente con valor `Null` al limpiar (`:73`, `:208`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: almacenamiento del anexo con la fila en una sola operación; carga múltiple.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con una NC con asociada o un técnico, no se puede subir, renombrar ni eliminar.
2. Alta sin título o sin archivo muestra el aviso correspondiente y no registra (`:238-247`).
3. Un título ya existente en la NC se rechaza en alta y al renombrar, sin distinguir mayúsculas.
4. Renombrar con el mismo nombre avisa y no escribe (`:376-379`).
5. Eliminar pide confirmación y, con «No», no cambia nada (`:464-466`).
6. Doble clic en un documento seleccionado abre el archivo (`:532-534`).
7. Tras cada operación la lista se recarga con los documentos vigentes.

## Pendientes

- Quién abre este formulario y quién escucha los eventos `Alta`, `CambioNombre`, `Eliminado`: no se ha rastreado.
- Definición de `Seleccionar`, `AbrirEnLocal` (`FG.bas:682, 764`), `getURLAnexoFinal` e `IDDocumentoCalculado`: no leídas en detalle.
- Dónde se asigna `EsTecnico` y `m_PermitidoEditar` (`Form_FormNCProyecto.cls:137-141`) se cita solo como origen.
