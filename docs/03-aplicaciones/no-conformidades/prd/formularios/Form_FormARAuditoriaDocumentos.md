# Form_FormARAuditoriaDocumentos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormARAuditoriaDocumentos.cls` (461 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar para gestionar los documentos (evidencias) de una tarea de acción realizada (AR) de una NC de auditoría: lista los existentes, permite anexar uno nuevo con título, cambiar el título y eliminarlo. Es un formulario **no enlazado** (el `.form.txt` no declara `RecordSource`; la lista es de tipo «Value List», `.form.txt:431`) y trabaja sobre la AR activa global `m_ObjARAuditoriaActiva` (`modules/Variables Globales.bas:129`). Las citas son `Form_FormARAuditoriaDocumentos.cls:línea` salvo indicación; `Ops` abrevia `classes/DocumentoAuditoriaOperaciones.cls` y `Doc`, `classes/DocumentoAuditoria.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoriaAR | Botón de documentos: cierra la instancia previa y abre el formulario sin argumentos | `Form_FormNCAuditoriaAR.cls:343-357` |
| Form_FormNCAuditoriaAR | Al registrar una tarea con fecha de fin real: si el usuario acepta adjuntar evidencia, lo abre con `OpenArgs` = fecha de fin real | `Form_FormNCAuditoriaAR.cls:438-450` |
| Form_FormNCAuditoriaAcciones | Abre el formulario sin argumentos | `Form_FormNCAuditoriaAcciones.cls:159-165` |

El formulario declara los eventos `Alta`, `CambioNombre` y `Eliminado` (`:7-9`); los llamantes los escuchan con `WithEvents` (`Form_FormNCAuditoriaAR.cls:5, 564, 588`; `Form_FormNCAuditoriaAcciones.cls:11, 1039`). No existe handler de `CambioNombre` en esos dos llamantes (búsqueda de `m_FormAnexos_` en ambos). «Salir» cierra con `acSaveNo` (`:13-17`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:142-279`), 3 secciones (`:300, 334, 707`) y los 11 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («Documentos de tareas de NC de auditoría») | etiqueta | — | Título | `.form.txt:319-320` |
| `Etiqueta50` («Nuevo documento») | etiqueta | — | Rótulo del bloque de alta | `.form.txt:356-357` |
| `Etiqueta62` («Título») | etiqueta | — | Rótulo de `Documento` | `.form.txt:618-619` |
| `Documento` | cuadro de texto | — | Título del documento nuevo (obligatorio al anexar) | `:36-40`, `.form.txt:649` |
| `URLDocumento` | cuadro de texto bloqueado | — | Muestra `..\` + nombre del fichero elegido; se vacía tras el alta | `:49, 64`, `.form.txt:576-590` |
| `ComandoAnexar` (rótulo «Comando47») | botón | `OnClick` | Alta del documento | `:19-86`, `.form.txt:520-522` |
| `ListaDocumentos` | lista, 3 columnas (0; 9624; 2286), con cabecera | `OnClick`, `OnDblClick` | Lista `IDDocumento;Documento;Archivo`; selecciona y habilita acciones; doble clic abre el fichero | `:403-461`, `.form.txt:416-436` |
| `ImagenAnexo` | imagen, oculta de inicio | `OnClick` | Abre el anexo seleccionado con la aplicación local | `:365-401`, `.form.txt:457-465` |
| `ComandoCambiarNombre` | botón, deshabilitado de inicio | `OnClick` | Cambia el título | `:88-147`, `.form.txt:666-680` |
| `ComandoEliminar` | botón, deshabilitado de inicio | `OnClick` | Elimina el documento | `:149-218`, `.form.txt:376-390` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:13-17`, `.form.txt:724-726` |

## Flujos

### Apertura (`Form_Load`, `:258-290`)
Ajusta tamaño (`Ajustar Me`, `:265`), guarda `OpenArgs` en `m_FechaFinPropuesta` (`:266`), exige `m_ObjARAuditoriaActiva` («El formulario se ha abierto con parámetros insuficientes», `:267-270`) y carga la lista (`:271, 292-313`). `EstablecerLista` fija la lista a `IDDocumento;Documento;Archivo`, deshabilita cambiar/eliminar, habilita anexar y añade una fila por documento de `m_ObjARAuditoriaActiva.Documentos`, con el nombre del fichero y los `;` del título sustituidos por `:` (`:330-351`). Si hay filas, selecciona la primera y ejecuta `ListaDocumentos_Click` (`:352-357`).

### Alta (`ComandoAnexar_Click`)
1. Los técnicos no pueden: «No tiene autorización para esa acción» (`:32-35`). Exige título (`:36-40`).
2. Abre el selector de fichero (`Seleccionar`, `:41`; `modules/Funciones Generales.bas:764`); si el fichero no existe sale en silencio (`:42-48`).
3. Crea `DocumentoAuditoria` con `IDAccionRealizada` de la AR activa y el título, y llama a `Registrar` (`:51-57`), que valida (`Ops:34-41`), calcula la ruta (`Ops:42`), rechaza si el destino existe y está abierto (`Ops:46-51`), inserta la fila (`Ops:64-84`) y copia el fichero sobreescribiendo (`Ops:91`).
4. Lanza `Alta`, vacía los campos, invalida la caché `Documentos` y recarga la lista (`:61-66`).

### Cambiar título, eliminar, abrir
- Cambiar: pide el nuevo título con `InputBox`; si está vacío o igual sale sin cambios (`:109-115`); llama a `CambiarNombre`, invalida caché, lanza `CambioNombre` y recarga (`:120-130`).
- Eliminar: confirmación con «No» por defecto (`:174-178`); `Ops.Eliminar` rechaza si el fichero está abierto (`Ops:180-185`), borra la fila y el fichero (`Ops:186-193`); lanza `Eliminado`, recarga y, si `FormNCAuditoria` está abierto mostrando `FormNCAuditoriaAcciones`, relanza `ListaAC_Click` (`:191-203`).
- Abrir (icono o doble clic): `AbrirEnLocal` sobre `URLAnexo` (`:382`; `modules/Funciones Generales.bas:682`).

### Cierre (`Form_Close`, `:220-256`)
Solo si se abrió con una fecha válida, la AR activa existe y `FormNCAuditoriaAR` sigue abierto: invalida la caché, asigna `FechaFinReal = m_FechaFinPropuesta` y ejecuta `Form_FormNCAuditoriaAR.ComandoRegistrar_Click` (`:223-241`), es decir, completa el registro de la tarea iniciado en el formulario llamante.

### Errores
Patrón común: error 1000 ⇒ aviso con `m_Error`; otro error ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:77-85`, `:138-146`, `:209-217`, `:247-255`, `:281-289`, `:392-400`, `:443-451`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no pueden anexar, renombrar ni eliminar (`:32-35, 94-97, 160-163`). `EsTecnico` es una variable global `EnumSino` (`modules/Variables Globales.bas:98`).
- **Título obligatorio** y **no repetido** al anexar: `MotivoNoOK` rechaza por `IDAccionRealizada` + `Documento` ya existentes con «Ya existe un documento para la misma Tarea…» (`Ops:114-139, 285-288`).
- **Identificador:** `IDDocumento` = máximo + 1 sobre `TbDocumentosAuditorias` (`Doc:158`; `modules/constructor.bas:1165`).
- **Ruta del anexo:** directorio de la auditoría resultante + `NCnn\ARnn\`, creados si faltan (`Doc:377-391`); nombre calculado por `getNombreAnexoARCalculado` (`Doc:523`).
- **Selección:** cambiar/eliminar/icono solo se habilitan al seleccionar una fila con identificador (`:409-436`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de documentos de la AR | `TbDocumentosAuditorias` (vía `getDocumentosAuditoria`) | Pendiente de detalle (función fuera de la ficha) | `classes/ARAuditoria.cls:76` |
| Leer un documento | `TbDocumentosAuditorias` | SQL dinámico concatenado (`IDDocumento`) | `constructor.bas:586-596` |
| Siguiente `IDDocumento` | `TbDocumentosAuditorias` | SQL dinámico concatenado (`Max`) | `constructor.bas:1165-1167` |
| Alta | `TbDocumentosAuditorias` | Recordset DAO `AddNew/Update` | `Ops:64-84` |
| Cambio de título | `TbDocumentosAuditorias` | SQL dinámico concatenado (`UPDATE` con el título entre comillas simples) | `Ops:238-240` |
| Baja | `TbDocumentosAuditorias` | SQL dinámico concatenado (`DELETE`) | `Ops:186-188` |
| Comprobación de repetido | `TbDocumentosAuditorias` | SQL dinámico concatenado | `Ops:275-292` |
| Fichero anexo | Sistema de ficheros | `CopyFile` (sobrescribe) y `DeleteFile` | `Ops:91, 192` |

No hay transacción: fila y fichero se escriben en pasos separados (fila antes que copia en el alta, `Ops:64-91`; borrado de fila antes que fichero, `Ops:188-193`).

## Coexistencia

- Las sentencias concatenan el título sin escapar comillas (`Ops:238, 278`); en PostgreSQL vinculado hay que parametrizar o escapar.
- `DELETE * FROM` (`Ops:186`) es dialecto Jet; en PostgreSQL es `DELETE FROM`.
- `Max(...)`+1 para el identificador (`constructor.bas:1165`) no es seguro con concurrencia; con PostgreSQL convendría una secuencia (autonumérico): a decidir.
- Los ficheros residen en una ruta de red configurada y se manipulan con `Scripting.FileSystemObject`; no migran con la base de datos.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual.** El alta escribe la fila antes de copiar el fichero; si la copia falla, queda una fila sin fichero (`Ops:64-91`). Al cerrar con fecha propuesta se registra la tarea aunque no se haya anexado ningún documento (`:236-241`: solo se comprueba que `Documentos` no sea `Nothing`).

**Posibles defectos (a verificar).**
- `CambiarNombre` llama a `DocumentoRepetido(nombre, IDAuditoriaResultanteCalculado, p_Error)`; el tercer argumento posicional es `p_IDNoConformidad`, no el error (`Ops:229`, firma `Ops:253-259`), y la consulta filtra por `IDAuditoria`, mientras el alta de una AR guarda `IDAccionRealizada` (`Ops:72-74, 276-278`). La detección de duplicados al renombrar podría no funcionar para documentos de AR.
- `ComandoEliminar_Click` declara un `m_DocumentoSeleccionado` local que oculta el del módulo (`:152`), por lo que siempre relee el documento de la lista.
- Código inalcanzable tras `Exit Sub` en `Form_Close` (`:242-246`).
- `ListaDocumentos` evalúa `ListCount = 1` (`:352`) como lista vacía (solo cabecera); confirmar con la cabecera `ColumnHeads` (`.form.txt:416`).

**Propuestas web (pendientes de decisión).** Transaccionalidad entre fila y fichero; subida de ficheros en lugar de ruta de red; validar duplicados por tarea con restricción única.

## Criterios de aceptación

1. Un usuario no técnico con título y fichero válidos crea un documento que aparece en la lista con su título y nombre de fichero (`:36-66`, `Ops:64-91`).
2. Anexar sin título o como técnico se rechaza con el mensaje indicado y sin escribir (`:32-40`).
3. Anexar un título ya usado en la misma tarea se rechaza (`Ops:134-139`).
4. Cambiar el título a uno vacío o igual no modifica nada (`:110-115`).
5. Eliminar pide confirmación y, al aceptar, borra fila y fichero; si el fichero está abierto no borra (`:174-178`, `Ops:180-193`).
6. Con una fecha de fin propuesta, cerrar la ventana completa el registro de la tarea (`:223-241`).
7. Cambiar y eliminar solo están activos con una fila seleccionada (`:409-436`).

## Pendientes

- `Seleccionar`, `AbrirEnLocal` y `getDocumentosAuditoria`: ver `modules/Funciones Generales.bas:764, 682` y `classes/ARAuditoria.cls:76`; consulta exacta de la lista sin documentar aquí.
- Existencia de `TbDocumentosAuditorias` en el esquema de main: el clon no incluye tablas; queda sujeto al ERD de origen.
