# Form_FormPlanificacionEquipos

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionEquipos.cls` (862 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla «Mantenimientos preventivos. Equipos asignados»: lista los equipos dados de alta en la planificación preventiva (`TbEquipos` unido a `TbPlanificacionEquipos`), con filtros por BUI, subsistema, texto de equipo y estado de sus planificaciones, y desde ella se da de alta o edita un equipo, se planifica un mantenimiento, se consultan los ya planificados, se saca o readmite el equipo de los preventivos y se exporta la lista a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se rellena asignando `RowSource` en ejecución (`Form_FormPlanificacionEquipos.cls:158`). Las citas son `Form_FormPlanificacionEquipos.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionGestion` | `DoCmd.OpenForm "FormPlanificacionEquipos"` sin argumentos | `Form_FormPlanificacionGestion.cls:69` |
| `Form_FormPlanificacionAlta` | si el formulario está abierto, llama a su `ActualizarLista` pública tras el alta | `Form_FormPlanificacionAlta.cls:104-105` |

Hacia otros formularios: `FormIDEquiposGestion` sin argumentos (`:40`); `FormPlanificacionEdicionEquipo` con `OpenArgs` = ID del equipo (`:210`); `FormPlanificacionAlta` con `OpenArgs` = `BUI|SUBSISTEMA|IDEquipo|Equipo` (`:445, 454`); `FormPlanificacionesBusqueda` con `OpenArgs` = `BUI|SUBSISTEMA|Equipo` (`:481, 490`). Antes de abrir los tres últimos cierra la instancia previa si existe (`:207-209, 451-453, 487-489`). Sale con `cmdSalir` (`:610-613`). Públicos: `ActualizarLista`, `ListaFiltrados_Click`, `Filtrar`, `LimpiarUltimos`, `ComandoActualizarTodo_Click`, `SacarEquipoDePlanificacion`, `MeterEquipoEnPlanificacion` (`:6, 12, 52, 128, 614, 700, 800`).

## Controles y eventos

Controles: 29 (37 `Begin` en el `.form.txt`: 29 controles más `Form`, `Section` y seis plantillas de estilo de `.form.txt:749-857`).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista 4 columnas con encabezado (ID oculto) | `OnClick` → `ListaFiltrados_Click` | Columnas `IDEquipo, Equipo, BUI, SUBSISTEMA`; al elegir fila fija la botonera según baja y nº de planificaciones | CAP-6 Planificación | `:614-699`, `.form.txt:914-937` |
| `ComboBUI` | combo (`LimitToList`) | `OnChange` y `OnExit` → `ComboBui_Change/_Exit` | Si cambia el BUI recarga subsistemas y equipos y la lista | CAP-12 Estructura y catálogos | `:502-549`, `.form.txt:1055-1079` |
| `ComboSubSistema` | combo | `OnChange` y `OnExit` | Recarga equipos y la lista | CAP-12 | `:550-595`, `.form.txt:1226-1251` |
| `ComboEquipos` | combo (`LimitToList`) | — | Texto de equipo para el filtro; su `RowSource` se reasigna en ejecución | CAP-5 Equipos | `:371-402`, `.form.txt:1121-1144` |
| `ComboFiltro` | combo lista de valores | — | Tres criterios de estado de planificaciones | CAP-6 | `:97-106`, `.form.txt:1434-1455` |
| `ComandoLimpiarBUI`, `ComandoLimpiarSubSistema`, `ComandoLimpiarIDEquipo`, `ComandoLimpiarFiltro` (4) | botones | `OnClick` | Vacían su combo y relanzan `ActualizarLista`; no hacen nada si el combo ya está vacío | CAP-6 | `:222-266, 356-370`, `.form.txt:1582-2097` |
| `ComandoActualizarTodo` (aviso «Actualizar Lista») | botón | `OnClick` | Ejecuta en cadena los cuatro «limpiar» | CAP-6 | `:12-32`, `.form.txt:2254-2266` |
| `ComandoEquipoAlta` («Alta Equipo») | botón | `OnClick` | Abre `FormIDEquiposGestion` | CAP-5 | `:34-51` |
| `ComandoEquipoEdicion` («Editar Equipo») | botón | `OnClick` | Exige ID numérico y abre `FormPlanificacionEdicionEquipo` | CAP-5 | `:191-221` |
| `ComandoSacarMeterEnPlanificacion` | botón; rótulo alterna | `OnClick` | Tras confirmar, da de baja o readmite el equipo | CAP-6 | `:267-333` |
| `ComandoPlanificarNuevoMantenimiento` («Nueva Planificación») | botón | `OnClick` | Abre `FormPlanificacionAlta` | CAP-6 | `:430-465` |
| `ComandoMantenimientosPlanificados` («Mtos. Planificados (n)») | botón | `OnClick` | Abre `FormPlanificacionesBusqueda` | CAP-6 | `:466-501` |
| `ComandoExportarExcel` | botón (deshabilitado de diseño) | `OnClick` | Exporta el `RowSource` actual con `GenerarConsultasSQL` | CAP-10 Informes | `:334-355`, `.form.txt:1159-1170` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | sin asignar | `:610-613` |
| `ComandoAyudaFormulario` | botón | `OnClick` declarado en el `.form.txt`; el `.cls` no contiene handler | Sin efecto aparente | sin asignar | `.form.txt:1543-1553` |
| `lblNotaAlPie` | etiqueta | — | «Equipo Fuera de mantenimiento desde : <fecha>» cuando hay baja | CAP-6 | `:643-656`, `.form.txt:1375-1387` |
| `Etiqueta142`, `Etiqueta162`, `Etiqueta187`, `Etiqueta208`, `lblTitulo1`, `lblTitulo2` (6) | etiquetas | — | Rótulos «BUI», «EQUIPOS», «SUBSISTEMA», «FILTRO» y título; `lblTitulo1` pasa a `Caption` del formulario | sin asignar | `:602`, `.form.txt:1028-1531` |
| `CuadroSuperior`, `Cuadro2`, `Marco120`, `Marco122` (4) | rectángulos y marcos | — | Decoración | sin asignar | `.form.txt:868-1005` |

## Flujos

### Apertura (`Form_Open`, `:596-609`)

1. `On Error Resume Next`, `AjustarTamaño Me` y `Caption` = título (`:597-602`); llama a `ActualizarLista` (`:603`).
2. Con los cuatro filtros vacíos, `Filtrar` devuelve una consulta con `BUI='-1'`: la lista se abre vacía (`:76-81`) y `ComandoExportarExcel` se deshabilita (`:162-163`).

### Filtrado (`ActualizarLista`, `:128-190`)

1. Deshabilita las cuatro acciones sobre equipo, restablece el rótulo «Mtos. Planificados» y oculta la nota (`:138-143`).
2. `Filtrar` compone el SQL; un valor con `|` se trata como error (`:149-156`). Asigna `RowSource` y `Requery` (`:157-159`).
3. Sin filas habilita `ComandoExportarExcel` a falso; con una sola, selecciona la fila 1 y lanza el clic; con varias, relanza el clic sobre la primera seleccionada (`:160-178`).
4. Error: `MsgBox` crítico y reloj de arena apagado (`:184-189`).

### Cascada de combos

`ComboBui_Change/_Exit`: si el BUI difiere de `strBUIAnterior`, ejecuta `CargarSubSistema` (vacía subsistema y equipo y recarga subsistemas del BUI, o `'-1'` sin BUI; `:403-417`), `CargarIDEquipo` (equipos del BUI y subsistema si ambos informados; si no, todos; `:380-388`) y `ActualizarLista` (`:508-513`). `ComboSubsistema` recarga equipos y lista (`:556-560`).

### Selección de equipo (`ListaFiltrados_Click`, `:614-699`)

1. Sin ID, deja todo deshabilitado (`:627-632`); con ID habilita edición y sacar/meter (`:633-634`).
2. Lee `FechabajaParaPlanificacion`: si es fecha, muestra la nota, deshabilita «Nueva Planificación» y rotula «Meter equipo en Preventivos»; si no, rotula «Sacar equipo de Preventivos» (`:640-656`).
3. Con 0 planificaciones «No Realizadas» deshabilita sacar/meter (`:662-671`). Con «Realizadas» = 0 deja «Mtos. Planificados (0)» deshabilitado; si no, lo habilita con el número (`:677-688`).

### Sacar o meter equipo (`:267-333`)

Exige ID numérico y rótulo reconocido; pide confirmación Sí/No (`:273-293`). **Sacar** (`SacarEquipoDePlanificacion`, `:700-799`): rechaza si ya tiene baja (`:737-740`), si `EquipoPlanificable` devuelve «No» (`:752-755`) o si tiene planificaciones no finalizadas (`:773-776`); actualiza `FechabajaParaPlanificacion` por DAO (`:777-786`). **Meter** (`MeterEquipoEnPlanificacion`, `:800-861`): si el equipo no está en `TbPlanificacionEquipos` lo da de alta con 12 meses (`:833-844`; `FG.bas:2781-2849`); si está, pone la baja a `Null` con `UPDATE` (`:846-851`). Termina con `ListaFiltrados_Click` (`:322`).

### Exportar a Excel (`:334-355`)

Pasa el `RowSource` a `GenerarConsultasSQL` (`:340-341`), que abre un recordset, crea un libro con nombre temporal en el directorio local configurado y lo rellena por Excel (`FG.bas:1584-1631`); «La consulta no contiene registros» si no hay filas (`FG.bas:1620-1623`).

## Reglas, validaciones y cálculos

- **Texto de equipo:** `Like '*texto*'`; BUI y subsistema, igualdad exacta (`:83-96`).
- **Criterios de `ComboFiltro`:** «Con Planificaciones pendientes», «Sin Planificaciones Pendientes» y «Sin Planificaciones» (`:97-102`); cualquier otro valor se ignora (`:103-104`). Orden por `TbEquipos.Equipo` (`:111, 116`).
- **Baja:** equipo fuera de preventivos si `FechabajaParaPlanificacion` es fecha (`:647`); solo se puede sacar sin planificaciones pendientes (`:773-776`).
- **Nº de planificaciones:** `NPlanificaciones` cuenta filas de `TbPlanificacion` según estado (`FG.bas:2604-2653`) y devuelve -1 ante error (`FG.bas:2660`).
- **Cierre de errores:** el error interno 1000 muestra el texto preparado; cualquier otro, número y descripción (`:27-31`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de equipos | `TbEquipos` ⋈ `TbPlanificacionEquipos` | SQL dinámico concatenado (BUI, subsistema y texto entre comillas) | — | `:74-118` |
| Filtros por estado | `TbPlanificacion`, `TbPlanificacionEquipos` | funciones VBA (`NPlanificaciones`, `EquipoSinPlanificaciones`) dentro del `WHERE` | — | `:98-102` |
| Combo de subsistemas | `TbSubsistemaBui` ⋈ `TbEquipos` | SQL dinámico concatenado | — | `:412-417` |
| Combo de equipos | `TbEquipos` | SQL dinámico concatenado | — | `:381-388` |
| Combo BUI (diseño) | `TbBUI` | `RowSource` guardado en el formulario | — | `.form.txt:1076` |
| Lectura de baja | `TbPlanificacionEquipos` | `Dame` (SQL concatenado, reintento con comillas) | — | `:640`; `FG.bas:300-311` |
| Baja del equipo | `TbPlanificacionEquipos` | DAO `Edit`/`Update` sobre SQL concatenado | no | `:777-786` |
| Readmisión | `TbPlanificacionEquipos` | `DoCmd.RunSQL` `UPDATE` con avisos apagados | no | `:846-851` |
| Alta en planificación | `TbPlanificacionEquipos` | DAO `AddNew`/`Update` | no | `FG.bas:2815-2834` |
| Exportación | la del `RowSource` | DAO `OpenRecordset` | — | `FG.bas:1619` |

## Coexistencia

Hay que adaptar a PostgreSQL vinculado: comodines `Like '*'` (`:83, 88, 93`), `Is Null`, funciones VBA llamadas desde el `WHERE` (`:98-102`, no existen en el servidor), `DateDiff('d',Now(),…)` y `IIf` de `NPlanificaciones` (`FG.bas:2625-2637`), `DoCmd.RunSQL` con `SetWarnings` (`:849-851`), y la conversión implícita de ID numérico a texto en `Dame`. La baja es una fecha en `TbPlanificacionEquipos`; el alta en planificación no usa autonumérico propio. Pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los filtros «Con/Sin Planificaciones pendientes» llaman `NPlanificaciones(0,0,[TbEquipos].[IDEquipo],'No Realizadas')` (`:98, 100`): por posición, el ID cae en `strEstadoPlanificacion` y el texto en `intDiasAviso` (`FG.bas:2577-2583`); con un estado no válido la función devuelve -1 (`FG.bas:2604-2607, 2660`), de modo que ninguno de los dos criterios devolvería filas.
- «Sin Planificaciones» llama `EquipoSinPlanificaciones(...,'No')` (`:102`), que en esa rama consulta `TbPlanificacionEquipos` (`FG.bas:778-788`); como la lista ya está unida a esa tabla, el criterio podría no filtrar nada.
- `ComboEquipos` y `ComboFiltro` no tienen evento en el `.form.txt`: elegir un valor no refresca la lista hasta que otra acción llame a `ActualizarLista`.
- `ListaFiltrados_Click` deshabilita sacar/meter cuando las planificaciones «No Realizadas» son 0 (`:667-668`), pero `SacarEquipoDePlanificacion` rechaza justo lo contrario (`:773-776`): sacar solo estaría habilitado cuando va a ser rechazado, y un equipo ya fuera, sin pendientes, no podría readmitirse desde la botonera.
- `SacarEquipoDePlanificacion` se llama solo con el ID (`:303`), por lo que `strFechaBaja` está vacío y se asigna a `FechabajaParaPlanificacion` (`:783`); no se asigna la fecha del día.
- `ComboSubsistema_Change/_Exit` comparan con `strBUIAnterior` (`:557, 580`) y `strSubSistemaAnterior` nunca se lee.
- `ComandoExportarExcel_Click` no apaga el reloj de arena en la ruta de error (`:350-354`) y `Form_Open` ignora errores (`:597`).
- SQL concatenado sin escapar comillas (`:85, 90, 95, 387, 415`); `ComandoAyudaFormulario` declara evento sin handler (`.form.txt:1543-1553`).
- `MeterEquipoEnPlanificacion` da de alta con 12 meses fijos (`:839`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro único por parámetros, refresco al cambiar cualquier filtro y baja con fecha explícita.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir sin filtros la lista aparece vacía y la exportación deshabilitada (`:76-81, 162-163`).
2. Elegir BUI recarga subsistemas y equipos y filtra la lista (`:508-513`).
3. Elegir un equipo muestra «Fuera de mantenimiento desde» y «Meter equipo en Preventivos» si tiene baja; si no, «Sacar equipo de Preventivos» (`:647-655`).
4. «Mtos. Planificados (n)» muestra las planificaciones realizadas y se deshabilita con 0 (`:677-688`).
5. Sacar con planificaciones pendientes muestra «Este equipo tiene al menos una planificación registrada no finalizada» (`:773-776`).
6. Exportar con la lista filtrada genera un libro con esas filas; sin filas, «La consulta no contiene registros» (`FG.bas:1620-1623`).
7. Tras crear una planificación desde `FormPlanificacionAlta` la lista se actualiza (`Form_FormPlanificacionAlta.cls:104-105`).

## Pendientes

- Contenido de `AjustarTamaño`, `FormularioAbierto` y del resto de `EquipoPlanificable` (`:207, 601`; `FG.bas:2138-2170`).
- Qué hace `FormIDEquiposGestion` al volver (no hay escucha ni refresco; `:40`).
- Comportamiento de asignar cadena vacía al campo de fecha de baja en el motor (`:783`).
- Cabecera de `SacarEquipoDePlanificacion` cita un `ComandoEquipoBaja_Click` inexistente en este formulario (`:714`).
