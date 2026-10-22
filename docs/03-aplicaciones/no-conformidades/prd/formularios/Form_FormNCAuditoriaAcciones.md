# Form_FormNCAuditoriaAcciones

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaAcciones.cls` (1397 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Acciones» de la ficha de NC de auditoría: lista las acciones correctivas (AC) de la NC activa en `ListaAC` y, debajo, las tareas (acciones realizadas, AR) de la AC seleccionada en `ListaAR`. Desde ella se dan de alta, editan, borran y consultan en historial AC y AR, se gestionan las evidencias de las tareas y se recalculan estados. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); las listas son de tipo «Value List» que el código rellena con `AddItem` (`:656, 673, 712, 732`; `.form.txt:359, 578`). Las citas son `Form_FormNCAuditoriaAcciones.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `ACAuditoriaOperaciones.cls` abrevia `classes/ACAuditoriaOperaciones.cls`, `ARAuditoriaOperaciones.cls` abrevia `classes/ARAuditoriaOperaciones.cls` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNCAuditoria` | Pestaña `NavAcciones` con destino `FormNCAuditoriaAcciones` (se muestra en `FrmDetalle`) | `Form_FormNCAuditoria.form.txt:1033` |
| `Form_FormNCAuditoriaAC`, `Form_FormNCAuditoriaAR` | Llaman a `ActualizarLineaListaAC` tras recalcular | `Form_FormNCAuditoriaAC.cls:196-197`, `Form_FormNCAuditoriaAR.cls:272-273` |

Usa `m_ObjNCAuditoriaActiva`; sin ella: «El formulario se ha abierto con parámetros insuficentes» (`:600-601`). Abre `FormNCAuditoriaAC` (`:37, 200`), `FormNCAuditoriaAR` (`:363, 499`) y `FormARAuditoriaDocumentos` (`:163`), cerrando antes la instancia abierta, y escucha sus eventos (`:11-13, 953-1180`). Los historiales se vuelcan a un fichero de texto con `HTMLENTXT` (`:236-240, 535-539`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | form | `OnLoad` | `EstablecerDatos` | `.form.txt:83`, `:563-589` |
| `ListaAC` | lista de valores | `OnClick` → `ListaAC_Click` | Selecciona la AC, carga el objeto y habilita sus botones y las tareas | `:825-880`, `.form.txt:342, 363` |
| `ListaAR` | lista de valores | `OnClick` → `ListaAR_Click` | Selecciona la tarea y habilita edición, historial y borrado | `:883-949`, `.form.txt:561, 582` |
| `OrdenAC`, `OrdenAR` | combos | `AfterUpdate` | Reordenan la lista conservando la selección | `:1359-1398, 1319-1358`, `.form.txt:1190-1287` |
| `ComandoACAlta` | botón | `OnClick` | Abre el alta de AC | `:16-55` |
| `ComandoACEdicion` | botón | `OnClick` | Abre la AC seleccionada | `:179-217` |
| `ComandoACBorrar` | botón | `OnClick` | Borra la AC (y sus tareas) previa confirmación | `:57-139` |
| `ComandoACHistorial` | botón | `OnClick` | Historial de cambios de la AC | `:219-258` |
| `ComandoActualizarACs` | botón | `OnClick` | Recalcula todas las AC de la NC | `:260-288` |
| `ComandoARAlta` | botón | `OnClick` | Abre el alta de tarea | `:333-380` |
| `ComandoAREdicion` | botón | `OnClick` | Abre la tarea seleccionada | `:466-516` |
| `ComandoARBorrar` | botón | `OnClick` | Borra la tarea previa confirmación | `:384-464` |
| `ComandoARHistorial` | botón | `OnClick` | Historial de cambios de la tarea | `:518-557` |
| `ComandoARDocumentos` | botón | `OnClick` | Abre la gestión de evidencias de la tarea | `:141-177` |
| `ComandoActualizarARs` | botón | `OnClick` | Recalcula las tareas de la AC seleccionada | `:290-331` |
| `ImagenAnexo` | imagen | `OnClick` | Si la tarea tiene un solo documento lo abre en local; si no, abre la gestión | `:752-809`, `.form.txt:750-759` |
| `ImagenNotaAC`, `ImagenNotaAR` | imágenes | `OnClick` | Indicadores de notas; abren la edición | `:811-823`, `.form.txt:944, 1056` |
| `Etiqueta50`, `Etiqueta64`, `Etiqueta75`, `Etiqueta246`, `Detalle` | rótulos y sección | — | Sin código | `.form.txt:299-1263` |

## Flujos

### Apertura (`EstablecerDatos`, `:590-634`)

Rellena los combos de orden con `EstablecerComboOrdenarPor` (`:1302-1306`), habilita el alta y los dos botones de recalcular, fija `Me.AllowEdits = True`, ordena por `PorNAccion` y carga `ListaAC` y `ListaAR` (`:604-623`).

### Cargar listas (`:635-749`)

`ListaAC` (cabecera «ID;Nº;Acción correctiva;Responsable;F.inicio;F.fin prev.;F.fin real») se llena con `constructor.getACsAuditoria` según el orden elegido; selecciona la primera fila y dispara `ListaAC_Click` (`:656-685`). `ListaAR` se llena con `constructor.getARsDeACAuditoria` de la AC seleccionada (`:712-743`). El punto y coma de los textos se sustituye por «:» para no romper la lista (`:673, 732`).

### Alta y edición de AC / AR

1. `ComandoACAlta` y `ComandoARAlta` rechazan a técnicos («No tiene autorización para esa acción») y exigen `DatosGeneralesOK = Sí`: «Debe regularizar primero los datos que faltan en la ficha general» (`:23-30, 340-357`).
2. Las ediciones (`ComandoACEdicion`, `ComandoAREdicion`) no repiten esa comprobación; piden seleccionar un elemento de la lista superior si no hay (`:186-194, 474-492`).
3. Al volver, `m_FormAC_Alta`, `m_FormAC_Editado`, `m_FormAR_Alta` y `m_FormAR_Editado` recargan la NC y las listas (`EstadoGrabar` para el alta de AC) y avisan si la NC pasa a «Cerrada» o deja de estarlo (`:953-1180`; textos «La NC de Auditoria pasa a estar cerrada» y «…ha dejado de estar cerrada»).

### Borrado

- **AC:** confirma con «¿Desea borrar la acción Correctiva seleccionada?» o, si tiene tareas, «…junto con sus N tareas asociadas?» (`:80-98`); llama a `ACAuditoriaOperaciones.Eliminar` (`:105-111`), invalida replanificaciones y ACs de la NC y recarga (`:113-124`).
- **AR:** confirma «¿Desea borrar la tarea seleccionada?» y llama a `ARAuditoriaOperaciones.Eliminar` (`:409-422`); recarga la NC, repinta pestañas y recalcula las AC (`:425-446`).

### Recalcular

`ComandoActualizarACs` ⇒ `ActualizarDatosACsAuditoria` sobre `m_NC.ACs` y refresco (`:267-274`); `ComandoActualizarARs` ⇒ `ActualizarDatosARsAuditoria` sobre las tareas de la AC (`:307-316`).

### Errores

Error 1000 ⇒ advertencia; otros ⇒ correo al administrador y `MsgBox` crítico (p. ej. `:48-54, 132-138`).

## Reglas, validaciones y cálculos

- **Permisos:** `EsTecnico = Sí` bloquea alta de AC/AR y borrado de AC/AR (`:23-26, 66-69, 340-343, 392-395`).
- **Prerrequisito de alta:** ficha general completa (`:27-30, 354-357`).
- **Borrado de AC:** elimina antes cada tarea (`ACAuditoriaOperaciones.cls:301-310`); después `DELETE` de la AC, recálculo de la NC e indicadores si la NC queda «Registrada» (`ACAuditoriaOperaciones.cls:314-326`).
- **Borrado de AR:** `DELETE` de la tarea, recálculo de AC y NC y reinicio de indicadores (`ARAuditoriaOperaciones.cls:334-349`).
- **Cierre/reapertura de NC:** compara el estado inicial con el final tras alta o edición (`:968-981, 1088-1101, 1149-1162`).
- **Refresco de una línea:** `ActualizarLineaListaAC/AR` sustituye en el `RowSource` la línea seleccionada por la recalculada (`:1182-1292`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Listar AC de la NC | `TbNCAuditoriaAccionCorrectivas` | SQL dinámico concatenado por `ID` (`getACsAuditoria`) | `constructor.bas:5146-5151` |
| Listar tareas de la AC | `TbNCAuditoriaAccionesRealizadas` | SQL dinámico concatenado por `IDAccionCorrectiva` (`getARsDeACAuditoria`) | `constructor.bas:5570-5577` |
| Borrar AC | `TbNCAuditoriaAccionCorrectivas` | SQL dinámico concatenado `DELETE`; sin transacción | `ACAuditoriaOperaciones.cls:314-317` |
| Borrar AR | `TbNCAuditoriaAccionesRealizadas` | SQL dinámico concatenado `DELETE`; sin transacción | `ARAuditoriaOperaciones.cls:334-337` |
| Cargar AC/AR/NC | `getACAuditoria`, `getARAuditoria`, `getNCAuditoria` | Objetos de dominio (SQL no revisado salvo `getACAuditoria`: `constructor.bas:5081-5085`) | `:71, 150, 433` |
| Recalcular | `ActualizarDatosACsAuditoria`, `ActualizarDatosARsAuditoria` | Escrituras de campos calculados (no revisadas) | `:267, 307` |

## Coexistencia

`DELETE *` con comodín (`ACAuditoriaOperaciones.cls:314`, `ARAuditoriaOperaciones.cls:334`) es dialecto Access: en PostgreSQL es `DELETE FROM`. Identificadores numéricos concatenados, listas de valores con `;` como separador y borrado en cascada manual sin transacción: con tablas vinculadas puede dejar tareas huérfanas ante un fallo. `Ordenar por` (`EnumOrden`) se aplica en el SQL de `constructor` (no revisado).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual:** la lista se mantiene en memoria y se refresca por sustitución de texto en `RowSource`; el estado de la NC se recalcula tras cada alta, edición o borrado.

**Posibles defectos (a verificar):**
- `AllowEdits = True` siempre (`:608`); la restricción a técnicos depende de cada botón. `ComandoACEdicion`, `ComandoAREdicion` y los historiales no la comprueban, aunque el guardado sí (`Form_FormNCAuditoriaAC.cls:243`, `Form_FormNCAuditoriaAR.cls:387`).
- `ActualizarLineaListaAC/AR` componen la línea previa con `j = 0 To 5` y dejan fuera la séptima columna (`:1217, 1275`).
- El mensaje de error de `ListaAR_Click` cita «ListaAC_Click» (`:943`).
- Tras borrar una tarea se recarga la AC con `Me.ListaAR.Column(0)` (id de tarea) en lugar del id de AC (`:428`).
- `m_FormAC_Editado` no recarga la NC ni repinta pestañas (`:998-1037`), a diferencia de `m_FormAR_Editado` (`:1128-1148`).
- No hay handler para el evento `Eliminado` de documentos; solo `Alta` muestra el icono (`:1039-1062`).

**Propuestas web pendientes de decisión:**
- Borrado en cascada en una única transacción.
- Sustituir las listas de valores por tablas con orden y selección propios.

## Criterios de aceptación

1. Al abrir, `ListaAC` muestra las AC de la NC ordenadas por número y la primera queda seleccionada con sus tareas (`:615-685`).
2. Seleccionar una AC habilita editar, borrar, historial y alta de tarea; seleccionar una tarea habilita editar, historial y borrar (`:854-858, 925-929`).
3. Alta de AC o tarea: técnico y ficha general incompleta se rechazan con los mensajes indicados (`:23-30, 340-357`).
4. Borrar una AC con tareas advierte de su número y las elimina (`:95-98`, `ACAuditoriaOperaciones.cls:301-310`); borrar una tarea recalcula AC y NC (`ARAuditoriaOperaciones.cls:340-349`).
5. Al cerrar o reabrir la NC por una alta o edición se muestra el aviso correspondiente (`:968-981`).
6. El orden elegido en `OrdenAC` y `OrdenAR` se aplica sin perder la selección (`:1370-1382, 1330-1342`).
7. Los historiales y las evidencias se abren desde los botones correspondientes (`:236-240, 163`).

## Pendientes

- Orden y SQL de `constructor.getACsAuditoria` y `getARsDeACAuditoria` (`:660, 719`), `getARAuditoria`, `getNCAuditoria` y `HTMHistorialAuditoria` (`:236, 535`).
- Cálculos de `ActualizarDatosACsAuditoria` y `ActualizarDatosARsAuditoria` (`:267, 307`) y `EstadoGrabar` (`:962`).
- Origen de los valores del combo `EnumOrden` (`:1302`).
