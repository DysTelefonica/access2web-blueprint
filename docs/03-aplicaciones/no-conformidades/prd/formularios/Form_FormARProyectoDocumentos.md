# Form_FormARProyectoDocumentos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormARProyectoDocumentos.cls` (516 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar para gestionar los documentos de una tarea de acción realizada (AR) de una NC de **proyecto**: lista, alta en dos pasos (Examinar y Aceptar), cambio de título y eliminación. Es el equivalente del formulario de documentos de AR de auditoría ([Form_FormARAuditoriaDocumentos](Form_FormARAuditoriaDocumentos.md)), con otra mecánica de alta y otras reglas. Es un formulario **no enlazado** (lista «Value List», `.form.txt:672`) y trabaja sobre la AR activa global `m_ObjARProyectoActiva` (`modules/Variables Globales.bas:120`). Las citas son `Form_FormARProyectoDocumentos.cls:línea` salvo indicación; `Ops` abrevia `classes/DocumentoProyectoOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoAR | Botón de documentos: cierra la instancia previa y abre sin argumentos | `Form_FormNCProyectoAR.cls:362-368` |
| Form_FormNCProyectoAR | Al cerrar una tarea con fecha de fin real: abre con `OpenArgs` = `FechaFinReal` | `Form_FormNCProyectoAR.cls:455-460` |

Declara los eventos `Alta`, `CambioNombre` y `Eliminado` (`:7-11`) y el llamante escucha `Alta` y `Eliminado` (`Form_FormNCProyectoAR.cls:577, 604`); `CambioNombre` no tiene handler allí. «Salir» cierra con `acSaveNo` (`:21-24`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:414-551`), 3 secciones (`:572, 606, 985`) y los 11 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («Documentación de tareas de proyecto») | etiqueta | — | Título | `.form.txt:591-592` |
| `Etiqueta62` («Título») | etiqueta | — | Rótulo de `Documento` | `.form.txt:859-860` |
| `Documento` | cuadro de texto | — | Título del documento; sirve para alta (con selección vacía) y para renombrar (con fila seleccionada) | `:227-231, 294, 354` |
| `URLDocumento` | cuadro de texto bloqueado | — | Muestra `..\` + nombre del fichero elegido | `:188`, `.form.txt:818-831` |
| `btnExaminar` (rótulo «Comando47») | botón | `OnClick` | Elige fichero y pasa a modo alta | `:162-206`, `.form.txt:761-763` |
| `btnAceptar` | botón | `OnClick` | Guarda el documento nuevo | `:212-271`, `.form.txt:956-958` |
| `ListaDocumentos` | lista de valores | `OnClick`, `OnDblClick` | Lista `IDDocumento;Documento;Archivo`; al seleccionar copia el título a `Documento` | `:277-320, 514-517`, `.form.txt:671-677` |
| `ImagenAnexo` | imagen, oculta de inicio | `OnClick` | Abre el anexo con la aplicación local | `:491-512`, `.form.txt:697-706` |
| `ComandoCambiarNombre` | botón | `OnClick` | Renombra con el título del cuadro | `:326-422`, `.form.txt:918-920` |
| `ComandoEliminar` | botón, deshabilitado de inicio | `OnClick` | Elimina el documento | `:428-485`, `.form.txt:629-631` |
| `cmdSalir` | botón | `OnClick` | Cierra | `.form.txt:1002-1004` |

## Flujos

### Apertura (`Form_Load`, `:30-63`)
Ajusta tamaño, guarda `OpenArgs` en `m_FechaFinPropuesta` (`:37-38`), exige `m_ObjARProyectoActiva` («parámetros insuficientes», `:39-45`) y llama a `EstablecerDatos` (`:80-105`): oculta la imagen, y habilita `Documento`, `btnExaminar` y `btnAceptar` solo si `EsTecnico` no es `Sí` (`:88-94`). Cualquier error cierra el formulario (`:62`). `EstablecerLista` (`:107-156`) rellena la lista con una fila por documento (título sin `;`, nombre de fichero) y selecciona la primera si hay filas (`:139-150`).

### Alta (Examinar → Aceptar)
1. `btnExaminar`: rechaza a técnicos (`:169-172`); `Seleccionar` (`:175`); si se cancela, sale (`:177`); si el fichero no existe, avisa (`:179-182`). Guarda la ruta en memoria (`:185`), muestra el nombre, propone como título el nombre base del fichero si el título está vacío (`:191-193`) y deselecciona la lista (`:196-197`).
2. `btnAceptar`: exige no ser técnico (`:222`), título (`:228-231`) y fichero elegido y existente (`:233-236`). Crea `DocumentoProyecto` con `IDAccionRealizada` y título, y llama a `Registrar` (`:239-248`); lanza `Alta`, invalida la caché, limpia controles y recarga (`:250-258`).
3. `Registrar` valida con `MotivoNoOK` (`Ops:207-214`), calcula la ruta (`Ops:215`), inserta la fila por recordset (`Ops:227-243`), anota un log de proyecto (`Ops:244-248`) y copia el fichero (`Ops:256`).

### Cambiar título
Exige no ser técnico, selección y título no vacío (`:339-360`); rechaza título igual (sin distinguir mayúsculas) al actual (`:364-367`) y título ya usado por otro documento de la AR (`:371-386`); llama a `CambiarNombre`, invalida caché, lanza `CambioNombre`, recarga y limpia (`:391-405`).

### Eliminar y abrir
`ComandoEliminar_Click` exige no ser técnico, `Form_FormNCProyecto.m_PermitidoEditar` distinto de `No` y selección (`:434-449`); confirma con «No» por defecto (`:452-454`); `Ops.Eliminar` borra la fila y el fichero y anota el log (`Ops:46-62`); recarga `m_ObjARProyectoActiva` con `getARProyecto`, lanza `Eliminado` y recarga la lista (`:458-472`). Abrir (icono o doble clic): `AbrirEnLocal` sobre `URLAnexo` (`:504`, `:514-517`).

### Errores
Los avisos usan `MsgBox` con `vbExclamation` para el error 1000 y `vbCritical` para el resto; solo algunos handlers llaman a `CorreoAlAdministrador` (`:266, 416, 480`), no `btnExaminar`, `ListaDocumentos_Click` ni `ImagenAnexo_Click` (`:201-205, 315-319, 511`).

## Reglas, validaciones y cálculos

- **Rol y estado de la NC:** técnicos sin acceso de escritura (`:89, 169, 222, 339, 434`); `EstablecerLista` deshabilita `btnAceptar` si `Form_FormNCProyecto.m_PermitidoEditar = No` (`:122-128`); cambiar y eliminar solo se habilitan si la NC no tiene `NCProyectoAsociada` y `m_PermitidoEditar = Sí` (`:302-306`).
- **Unicidad del título:** por NC o por AR, `MotivoNoOK` y `DocumentoRepetido` (`Ops:97-105, 157`); en renombrado, además comprobación en memoria sin distinguir mayúsculas (`:364-386`).
- **Evidencia mínima:** no se elimina el único documento de una AR con `FechaFinReal` informada: «Sólo tiene un documento y está cerrada…» (`Ops:38-43`).
- **Aviso de repetidos:** el título del alta se propone del nombre del fichero (`:191-193`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de documentos de la AR | `TbNCDocumentos` (`WHERE IDAccionRealizada=…`) | SQL dinámico concatenado | `constructor.bas:621-651` |
| Leer un documento | `TbNCDocumentos` | SQL dinámico concatenado | `constructor.bas:491-520` (filtro por `IDDocumento`, ver Pendientes) |
| Alta | `TbNCDocumentos` | Recordset DAO `AddNew/Update` | `Ops:227-243` |
| Cambio de título | `TbNCDocumentos` | SQL dinámico concatenado | `Ops:302-304` |
| Baja | `TbNCDocumentos` | SQL dinámico concatenado | `Ops:46-49` |
| Comprobación de repetido | `TbNCDocumentos` | SQL dinámico concatenado | `Ops:97-107` |
| Log de proyecto | vía `RegistrarLogProyecto` | Pendiente de detalle | `Ops:54, 59, 245, 247` |
| Fichero anexo | Sistema de ficheros | `CopyFile`, `DeleteFile` | `Ops:256, 51` |

Sin transacción: fila, log y fichero se escriben en pasos separados (`Ops:227-256`).

## Coexistencia

- Concatenación del título sin escapar comillas (`Ops:100, 302`): parametrizar en PostgreSQL.
- `DELETE *` (`Ops:46`) es dialecto Jet.
- Identificador `IDDocumento` calculado por la clase antes del alta (`Ops:221`): con PostgreSQL, secuencia; mecánica del cálculo en Pendientes.
- Ficheros en una ruta de red configurada; no migran con la base.
- Acceso a la variable pública de otro formulario (`Form_FormNCProyecto.m_PermitidoEditar`, `:122, 303, 439`): en la web, estado de permisos explícito.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual.** Con una fila seleccionada, `Documento` edita el título; sin selección, prepara un alta. `m_FechaFinPropuesta` se guarda en `Form_Load` pero no se vuelve a leer en este fichero (`:13, 38`), y a diferencia de la ficha de auditoría no hay `Form_Close`.

**Posibles defectos (a verificar).**
- `EstablecerDatos` habilita `btnAceptar` (`:94`) y `EstablecerLista` lo vuelve a decidir (`:122-128`), pero nunca deshabilita `btnExaminar` ni `Documento` cuando `m_PermitidoEditar = No`.
- `ListaDocumentos_Click` consulta `m_ObjNCProyectoActiva.NCProyectoAsociada` (`:302`) sin comprobar que `m_ObjNCProyectoActiva` exista.
- Los errores de `ImagenAnexo_Click` muestran `Err.Description` aunque el error sea el controlado 1000 (`:509-511`).
- El borrado de fila ocurre antes que el del fichero y no hay marcha atrás (`Ops:46-52`).

**Propuestas web (pendientes de decisión).** Unificar con el formulario de auditoría un solo componente de documentos; subida de ficheros en lugar de ruta de red; borrar con transacción.

## Criterios de aceptación

1. Un técnico ve `Documento`, Examinar y Aceptar deshabilitados y no puede anexar (`:89-94, 169`).
2. Aceptar sin título o sin fichero elegido no escribe y avisa (`:228-236`).
3. Tras Examinar sin título, este se rellena con el nombre base del fichero (`:191-193`).
4. Un alta correcta añade la fila en la lista y limpia los controles (`:254-258`).
5. Renombrar a un título igual o ya existente se rechaza (`:364-386`).
6. No se puede eliminar el único documento de una AR cerrada (`Ops:38-43`).
7. Con la NC no editable no se puede eliminar ni dar de alta (`:122-128, 439-442`).

## Pendientes

- Filtro de `getDocumentoProyecto` (`constructor.bas:491-520`): la lectura cita `IDDocumento=` con el parámetro `p_IDNCResultante`; revisar la rama exacta.
- `RegistrarLogProyecto`, `IDDocumentoCalculado` de `DocumentoProyecto` y `getARProyecto`: no detallados en esta ficha.
- Existencia de `TbNCDocumentos` en el esquema de main: el clon no incluye tablas; sujeto al ERD de origen.
