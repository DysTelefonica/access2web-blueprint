# Form_FormNCAuditoriaDocumentos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaDocumentos.cls` (511 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Documentos» de la ficha de una NC de auditoría: lista los anexos de la NC, permite subir uno nuevo (Examinar + Aceptar), cambiar el título de uno existente, abrirlo y eliminarlo. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista es una lista de valores que se rellena por código desde `m_ObjNCAuditoriaActiva.DocumentosCompletos` (`:109-139`). Las citas son `Form_FormNCAuditoriaDocumentos.cls:línea` salvo indicación; `DocumentoAuditoriaOperaciones.cls` abrevia `classes/DocumentoAuditoriaOperaciones.cls`.

## Entrada y salida

Se muestra como subformulario de navegación de `FormNCAuditoria` mediante el botón de pestaña cuyo destino es `FormNCAuditoriaDocumentos` (`Form_FormNCAuditoria.form.txt:1197`). Sin objeto activo avisa «El formulario se ha abierto con parámetros insuficientes.» (`:34-37`). Declara tres eventos públicos `Alta`, `CambioNombre` y `Eliminado` (`:17-19`), que lanza con `RaiseEvent` (`:243, 394, 462`); en el código revisado no se ha localizado ningún suscriptor (ver Pendientes).

## Controles y eventos

`.form.txt` declara 19 bloques `Begin`: el formulario (`.form.txt:5`), 8 bloques de valores por defecto (`:416-553` y anteriores), la sección `Detalle` (`:577`) y los 9 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaDocumentos` | lista de 3 columnas (`IDDocumento;Documento;Archivo`) | `OnClick`, `OnDblClick` | Selecciona un anexo para editar; el doble clic lo abre | `:271-312, 508-511`, `.form.txt:625-647` |
| `Documento` | cuadro de texto («Título») | — | Título del anexo, nuevo o a renombrar | `.form.txt:859`, `:220, 346` |
| `URLDocumento` | cuadro de texto bloqueado | — | Muestra «..\» + nombre del fichero elegido | `:181`, `.form.txt:786-800` |
| `btnExaminar` | botón (rótulo `Comando47`) | `OnClick` | Elige el fichero local | `:155-199`, `.form.txt:730-732` |
| `btnAceptar` | botón «Aceptar» | `OnClick` | Guarda el anexo nuevo | `:205-265`, `.form.txt:924-926` |
| `ComandoCambiarNombre` | botón «Cambiar nombre» | `OnClick` | Renombra el anexo seleccionado | `:318-414`, `.form.txt:887-889` |
| `ComandoEliminar` | botón «Eliminar» | `OnClick` | Borra el anexo seleccionado | `:420-479`, `.form.txt:598-600` |
| `ImagenAnexo` | imagen con hipervínculo `#`, oculta | `OnClick` | Abre el anexo seleccionado | `:485-506`, `.form.txt:674-678` |
| etiqueta «Título» | etiqueta | — | Rótulo de `Documento` | `.form.txt:829` |

## Flujos

### Apertura (`Form_Load`, `:25-57`; `EstablecerDatos`, `:74-98`)

Oculta `ImagenAnexo`, habilita edición y activa `Documento`, `btnExaminar` y `btnAceptar` salvo para técnicos (`:78-88`). `EstablecerLista` pone `RowSource` a la cabecera, desactiva «Cambiar nombre» y «Eliminar», invalida la caché `DocumentosCompletos` y la recarga (`:109-126`); por cada anexo añade `id;título;nombre de fichero`, sustituyendo `;` del título por `:` (`:134-139`). No se preselecciona ningún elemento (`:141-142`).

### Alta (Examinar + Aceptar)

1. «Examinar»: técnico ⇒ rechazo (`:162-165`); abre el selector de ficheros (`:168`; `Funciones Generales.bas:764`), comprueba que existe (`:172-175`), guarda la ruta en memoria, muestra el nombre y, si el título está vacío, usa el nombre del fichero sin extensión; deselecciona la lista (`:178-190`).
2. «Aceptar»: técnico ⇒ rechazo; título obligatorio («Debe poner un título al documento.») y fichero elegido y existente (`:215-229`). Crea `DocumentoAuditoria` con `IDNoConformidad` y título y llama a `Registrar` (`:232-241`), lanza `Alta`, limpia cachés, resetea controles y recarga la lista (`:243-252`).

### Cambiar nombre (`:318-414`)

Rechazo a técnicos; exige selección y título no vacío (`:331-352`); rechaza nombre igual al actual sin distinguir mayúsculas (`:356-359`) o ya usado por otro anexo de la NC (`:363-378`); delega en `CambiarNombre` y refresca (`:383-397`).

### Eliminar (`:420-479`)

Rechazo a técnicos y si `Form_FormNCAuditoria.m_PermitidoEditar = No` («Operación no autorizada», `:426-434`); exige selección; confirma con «No» por defecto (`:444`); llama a `Eliminar`, recarga la NC activa desde `constructor.getNCAuditoria`, lanza `Eliminado` y recarga la lista (`:450-466`).

### Abrir anexo (`:485-511`)

Exige selección y llama a `AbrirEnLocal` con la ruta del anexo (`:498`); el doble clic en la lista lo dispara si la imagen está visible (`:510`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ mensaje y `CorreoAlAdministrador` salvo en `btnExaminar`, `ListaDocumentos_Click` y `ImagenAnexo_Click`, que solo muestran el mensaje (`:48-56, 193-198, 306-311, 503-506`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no pueden examinar, aceptar, renombrar ni eliminar (`:162, 215, 331, 426`). «Eliminar» y «Cambiar nombre» solo se habilitan al seleccionar si `m_PermitidoEditar = Sí` (`:296-299`); `btnAceptar` y `btnExaminar` no dependen de él (`:121-122`).
- **Alta** (`DocumentoAuditoriaOperaciones.cls:114-142`): el anexo debe pertenecer a una NC, AR o auditoría; título no vacío; fichero accesible; sin otro anexo del mismo título para la misma NC («Ya existe un documento para la misma No Conformidad de Auditoria»). Si el fichero de destino existe y está abierto: «Ya hay un archivo con el mismo nombre y parece estar abierto» (`DocumentoAuditoriaOperaciones.cls:46-51`).
- **Renombrado** (`DocumentoAuditoriaOperaciones.cls:221-236`): `IDDocumento` y nombre obligatorios; no se repite el nombre en la auditoría.
- **Borrado** (`DocumentoAuditoriaOperaciones.cls:180-193`): falla si el fichero está abierto; borra la fila y después el fichero físico.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de anexos de la NC | `constructor.getDocumentosCompletosAuditoria` (consulta no revisada) | pendiente | `NCAuditoria.cls:984-993`, `constructor.bas:4973` |
| Anexo seleccionado | `constructor.getDocumentoAuditoria` | pendiente | `:302`, `constructor.bas:560` |
| Alta de anexo | `TbDocumentosAuditorias` (`IDDocumento`, `IDNoConformidad`, `Documento`, `NombreAnexo`) | DAO `AddNew/Update` sobre la tabla | `DocumentoAuditoriaOperaciones.cls:56-86` |
| Copia física del fichero | ruta final de `getURLAnexoFinal` (carpeta de anexos configurada) | `fso.CopyFile` con sobrescritura | `DocumentoAuditoriaOperaciones.cls:42, 91` |
| Comprobación de duplicado | `TbDocumentosAuditorias` por `IDNoConformidad` y `Documento` | SQL dinámico concatenado | `DocumentoAuditoriaOperaciones.cls:253-289` |
| Renombrado | `TbDocumentosAuditorias.Documento` | SQL dinámico concatenado (`UPDATE`) | `DocumentoAuditoriaOperaciones.cls:238-240` |
| Borrado | `TbDocumentosAuditorias` + fichero | SQL dinámico concatenado (`DELETE`) + `fso.DeleteFile` | `DocumentoAuditoriaOperaciones.cls:186-193` |

Sin transacciones: la fila se inserta antes de copiar el fichero (`DocumentoAuditoriaOperaciones.cls:84-91`) y se borra antes de eliminarlo (`DocumentoAuditoriaOperaciones.cls:188-193`).

## Coexistencia

Pendiente de `../coexistencia.md`. A adaptar: `DELETE *` y títulos concatenados entre comillas simples (`DocumentoAuditoriaOperaciones.cls:186, 238, 278`; un título con apóstrofo rompe la consulta en cualquier dialecto), `IDDocumento` calculado en código (`DocumentoAuditoriaOperaciones.cls:60`) en lugar de secuencia, y el almacenamiento de ficheros en una ruta de red configurada con `FileSystemObject`, sin equivalente directo en web (almacén de objetos o similar).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `CambiarNombre` llama a `DocumentoRepetido(nombre, IDAuditoriaResultanteCalculado, p_Error)` (`DocumentoAuditoriaOperaciones.cls:229`): el tercer argumento posicional es `p_IDNoConformidad`, no `p_Error`, y la comprobación de repetido se hace por auditoría, no por NC.
- Un título con apóstrofo rompe `UPDATE` y la comprobación de repetido (`DocumentoAuditoriaOperaciones.cls:238, 278`).
- «Cambiar nombre» solo comprueba duplicados en la NC (`:363-378`), pero la regla de `DocumentoAuditoriaOperaciones.cls:234` habla de toda la auditoría: dos criterios distintos.
- Con `bPuedeEditar` falso, el bloque «heredado» fuerza `btnExaminar` y `btnAceptar` a habilitados (`:83-88` frente a `:118-122`); solo la comprobación `EsTecnico` los frena.
- `Registrar` inserta la fila antes de copiar el fichero (`DocumentoAuditoriaOperaciones.cls:84-91`): si la copia falla queda un anexo sin fichero.
- `Eliminar` no pide confirmación de seguridad adicional más allá del «Sí/No» y borra el fichero físico de forma irreversible (`DocumentoAuditoriaOperaciones.cls:191-193`).
- `ListaDocumentos.RowSource` usa `Documento` en minúsculas en código y `DOCUMENTO` en el `.form.txt` (`:110`, `.form.txt:642`).
- El texto del botón «Examinar» es `Comando47` (`.form.txt:731`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida directa con arrastrar y soltar; permitir renombrar sin recargar la lista; almacenamiento de anexos fuera del sistema de ficheros compartido.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico no puede añadir, renombrar ni eliminar anexos (`:162, 215, 331, 426`).
2. No se añade un anexo sin título ni sin fichero existente (`:221-229`).
3. No se añaden dos anexos con el mismo título a una NC (`DocumentoAuditoriaOperaciones.cls:134-136`).
4. El renombrado rechaza nombre vacío, igual al actual o ya usado en la NC (`:348-378`).
5. La eliminación pide confirmación y, si la NC no es editable, se rechaza (`:431-446`).
6. La lista muestra identificador, título y nombre de fichero de cada anexo (`:134-139`).
7. Abrir un anexo requiere seleccionarlo (`:490-498`).

## Pendientes

- Ruta de red del anexo (`DocumentoAuditoria.getURLAnexoFinal`, `classes/DocumentoAuditoria.cls:289`) y propiedad `URLAnexo`: no revisadas.
- Consultas de `constructor.getDocumentosCompletosAuditoria` y `getDocumentoAuditoria`.
- Suscriptores de los eventos `Alta`, `CambioNombre` y `Eliminado`: búsqueda por texto solo en esta ficha; resolución dinámica pendiente.
- `Seleccionar` y `AbrirEnLocal` (`Funciones Generales.bas:764, 682`): comportamiento exacto no revisado.
