# Form_FormAuditoria

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormAuditoria.cls` (187 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de alta y edición de una **auditoría** (cabecera): tipo (interna o externa), fecha de inicio y fecha de fin. Es **no enlazado** (sin `RecordSource` en el `.form.txt`; el tipo es una lista de valores, `.form.txt:305-306`) y escribe por la clase `AuditoriaOperaciones`. Desde él se pueden preparar documentos de la auditoría con `Form_FormAuditoriaDocumentos`. Las citas son `Form_FormAuditoria.cls:línea` salvo indicación; `Ops` abrevia `classes/AuditoriaOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesAuditorias | Alta: cierra la instancia previa, pone `m_ObjAuditoriaActiva = Nothing` y abre el formulario | `Form_Form0BDOpcionesAuditorias.cls:45-49` |
| Form_FormAuditoriasGestion | Edición: asigna la auditoría seleccionada a `m_ObjAuditoriaActiva`, abre el formulario y se suscribe a sus eventos | `Form_FormAuditoriasGestion.cls:166-172` |

Lanza `Alta` o `Editada` al registrar (`:92-96`); el llamante de gestión los escucha (`Form_FormAuditoriasGestion.cls:10, 604, 638`). Se cierra a sí mismo tras registrar (`:103`) y con «Salir» (`:13-17`). Abre `FormAuditoriaDocumentos` sin argumentos (`:35-38`).

## Controles y eventos

`.form.txt` declara 25 bloques `Begin`: el formulario (`.form.txt:4`), 9 de valores por defecto (`:127-186`, incluido un `CustomControl`), 3 secciones (`:191, 220, 776`) y los 12 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («REGISTRO DE AUDITORÍA») | etiqueta | — | Título; su texto pasa a la barra del formulario | `:165`, `.form.txt:209-210` |
| `Tipo` | lista desplegable «Value List» `Interna;Externa` | — | Tipo de auditoría | `.form.txt:304-306`, `:72, 135` |
| `FechaInicio`, `FechaFin` | cuadros de texto con formato «Short Date» | — | Fechas; se cargan con formato `dd/mm/yyyy` | `.form.txt:242-243, 359-360`, `:138, 141` |
| `Etiqueta251`, `Etiqueta253`, `Etiqueta255` | etiquetas «Fecha Inicio», «Tipo», «Fecha Final» | — | Rótulos | `.form.txt:273-274, 333-334, 390-391` |
| `Etiqueta256` («Adjuntar documentación») | etiqueta | — | Rótulo del botón de documentos | `.form.txt:411-412` |
| `ComandoAdjuntarDocumentacion` (rótulo «Grabar») | botón | `OnClick` | Abre la ventana de documentos | `:20-53`, `.form.txt:432-434` |
| `ComandoRegistrar` («Grabar») | botón | `OnClick` | Valida, graba, anexa y cierra | `:55-114`, `.form.txt:795-797` |
| `lblEstado` | etiqueta oculta | — | Destino de los mensajes de progreso (`AvanceCerrar`) | `.form.txt:833-843`, `:99, 171` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin grabar | `:13-17`, `.form.txt:864-866` |

Evento del formulario: `OnOpen` (`.form.txt:37`, `:154-186`).

## Flujos

### Apertura (`Form_Open`, `:154-186`)
Ajusta tamaño (`Ajustar Me`, `:163`), fija el título (`:165`) y llama a `EstablecerDatos` (`:167`). Si `m_ObjAuditoriaActiva` es `Nothing` (alta) crea una `Auditoria` vacía y no rellena nada (`:120-123`). Si existe (edición), la relee con `constructor.getAuditoria` y guarda la copia como `m_AuditoriaAlInicio` (`:125`); si no está registrada, error «La auditoría no parece estar registrada» (`:129-132`); vuelca tipo y fechas (`:133-143`). Cualquier error cierra el formulario (`:185`).

### Registrar (`ComandoRegistrar_Click`, `:55-114`)
1. Rechaza a técnicos (`:64-67`). Copia tipo y fechas a `m_ObjAuditoriaActiva` (`:71-75`) y llama a `AuditoriaOperaciones.Registrar` (`:76-78`).
2. `Registrar` valida: fecha inicial y final informadas, final no anterior a inicial, tipo «Interna» o «Externa» (`Ops:25-40`). En alta calcula `IDAuditoria` y añade la fila; en edición ejecuta `Edit` sobre la fila existente (`Ops:76-108`). Invalida las cachés `ColAuditorias` y `ColAuditoriasNombres` (`Ops:113-114`).
3. Solo en alta y si hay documentos preparados (`m_ObjColDocumentosAuditoriaParaAlta`), los registra con `AnexoMultiple` (uno por entrada, nombre → ruta local; `Ops:206-219`), vacía la colección y cierra la ventana de documentos (`:82-91`).
4. Lanza `Alta` (si no había auditoría previa) o `Editada` (`:92-96`), cierra el progreso y el formulario (`:99, 103`).

### Adjuntar documentación (`:20-53`)
Rechaza a técnicos (`:27-30`); si no hay auditoría activa crea una vacía (`:31-33`); cierra la ventana de documentos si ya estaba abierta y la abre (`:35-38`).

### Errores
Patrón común: error 1000 ⇒ aviso con `m_Error`; otro ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:46-52, 107-113, 178-184`).

## Reglas, validaciones y cálculos

- **Rol:** no se comprueba al abrir; los técnicos no pueden grabar ni adjuntar (`:27-30, 64-67`). `EsTecnico` es la variable global `modules/Variables Globales.bas:98`.
- **Fechas y tipo:** reglas de `Ops:25-40` (arriba). El tipo se valida también en código aunque el control sea una lista desplegable (`Ops:37-40`).
- **Identificador:** `IDAuditoria` lo calcula `IDAuditoriaCalculada` en alta (`Ops:78`).
- **Sin comprobación de solapes ni de duplicados** de auditorías en `Ops:18-50`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Cargar auditoría en edición | `TbAuditorias` (vía `getAuditoria`) | Pendiente de detalle | `:125` |
| Alta de auditoría | `TbAuditorias` | Recordset DAO `AddNew/Update` | `Ops:77, 90-99` |
| Edición de auditoría | `TbAuditorias` | SQL dinámico concatenado (`SELECT … WHERE IDAuditoria=`) + recordset `Edit/Update` | `Ops:83-84, 102-107` |
| Alta de documentos preparados | `TbDocumentosAuditorias` y ficheros | Ver Form_FormAuditoriaDocumentos | `Ops:206-219` |

Sin transacción: la fila de auditoría se confirma antes de registrar los documentos (`Ops:99` frente a `:82-91`), de modo que un fallo en un anexo deja la auditoría creada.

## Coexistencia

- Fechas gestionadas como texto del formulario con `Nz(…, "")` y `IsDate`/`CDate` (`:72-74`; `Ops:25-33`): en PostgreSQL, tipos `date` y formato explícito.
- `IDAuditoria` calculado en la aplicación (`Ops:78`): con PostgreSQL, secuencia o autonumérico; hay que decidir cómo convive con otras cargas.
- `DELETE *` y otras construcciones Jet de la clase no intervienen aquí (solo `Eliminar`, `Ops:173`, fuera del formulario).
- Las cachés de entorno (`ColAuditorias`) deben invalidarse tras escribir (`Ops:113-114`).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual.** Un solo formulario para alta y edición, distinguidos por la existencia de `m_ObjAuditoriaActiva`. Los documentos de una auditoría nueva se acumulan en una colección global hasta que se graba.

**Posibles defectos (a verificar).**
- `ComandoAdjuntarDocumentacion` tiene el rótulo «Grabar» (`.form.txt:433`) pese a abrir documentos.
- La barra de título del `.form.txt` es «Registro de NPedido» (`.form.txt:36`) y solo se corrige por código (`:165`).
- En edición, la colección `m_ObjColDocumentosAuditoriaParaAlta` no se procesa (`:82`); no se ve que `cmdSalir` ni `Form_Open` la vacíen, por lo que podrían quedar documentos de una sesión anterior.

**Propuestas web (pendientes de decisión).** Seleccionar el tipo con lista cerrada y fechas con selector; permitir adjuntos también en edición con el mismo flujo; alta de auditoría y documentos en una sola operación.

## Criterios de aceptación

1. Alta con tipo y fechas válidos crea la fila, lanza `Alta` y cierra el formulario (`:92-103`, `Ops:93-99`).
2. Fecha final anterior a inicial, fechas ausentes o tipo distinto de Interna/Externa se rechazan con el mensaje indicado sin escribir (`Ops:25-40`).
3. Edición carga tipo y fechas de la auditoría activa y, al grabar, modifica la misma fila y lanza `Editada` (`:133-143`, `Ops:102-107`, `:95`).
4. Un técnico no puede grabar ni adjuntar documentación (`:27-30, 64-67`).
5. Los documentos preparados en alta se registran tras crear la auditoría y la ventana de documentos se cierra (`:82-91`).
6. Tras grabar, la lista de auditorías del entorno refleja el cambio (`Ops:113-114`).

## Pendientes

- `getAuditoria`, `IDAuditoriaCalculada`, `AvanceCerrar`: no detallados en esta ficha.
- `m_ObjColDocumentosAuditoriaParaAlta` (`modules/Variables Globales.bas:131`): ciclo de vida completo en Form_FormAuditoriaDocumentos.
- Existencia de `TbAuditorias` en el esquema de main: el clon no incluye tablas; sujeto al ERD de origen.
