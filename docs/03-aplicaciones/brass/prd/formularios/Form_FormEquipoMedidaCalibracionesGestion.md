# Form_FormEquipoMedidaCalibracionesGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoMedidaCalibracionesGestion.cls` (330 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de las calibraciones del equipo de medida activo: lista (`ID`, fecha inicial, fecha final) con alta, edición, eliminación y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); toma el equipo de la global `m_ObjEquipoMedidaActivo` (`Variables Globales.bas:253`) y las calibraciones de su propiedad `Calibraciones` (`EquipoMedida.cls:208-225`). La edición de cada calibración se delega en `FormEquipoMedidaCalibracion`. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas son `Form_FormEquipoMedidaCalibracionesGestion.cls:línea` salvo indicación; aquí se abrevian a `:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquiposMedidaGestion` | `ComandoCalibraciones_Click`: fija `m_ObjEquipoMedidaActivo` con el equipo seleccionado, cierra el formulario si estaba abierto y lo abre | `Form_FormEquiposMedidaGestion.cls:93-125` |

Escribe la global `m_ObjCalibracionEquipoMedidaActiva` (`VG.bas:254`): `Nothing` en el alta (`:76`) y la calibración seleccionada en la edición (`:115`). Abre `FormEquipoMedidaCalibracion` (`:80, 119`) y escucha sus eventos `Alta` y `Edicion` (`:6, 283, 308`; declarados en `Form_FormEquipoMedidaCalibracion.cls:7-8`, lanzados en `:44, 46`). Sale con `cmdSalir` (`:63-65`).

## Controles y eventos

Controles: 21 (21 `Begin` con tipo en el `.form.txt`: formulario, 8 definiciones por defecto `.form.txt:723-852`, 3 secciones, 2 etiquetas, 1 lista, 4 botones de detalle, 1 texto y `cmdSalir`).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista 3 columnas con encabezados (ID oculto, `ColumnWidths ="0;1701;1701"`) | `OnClick` → `ListaFiltrados_Click` | Carga la calibración de la fila y habilita Editar y Eliminar | CAP-5 | `:235-280`, `.form.txt:931-966` |
| `ComandoAlta` («Alta») | botón | `OnClick` | Abre `FormEquipoMedidaCalibracion` en alta | CAP-5 | `:67-92`, `.form.txt:995-1020` |
| `ComandoEditar` («Editar») | botón | `OnClick` | Abre `FormEquipoMedidaCalibracion` con la calibración seleccionada | CAP-5 | `:94-131`, `.form.txt:968-993` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` | Borra la calibración seleccionada y repinta | CAP-5 | `:133-173`, `.form.txt:1022-1051` |
| `ComandoExportarExcel` (rótulo «Comando227», icono) | botón, `Enabled = NotDefault` en diseño | `OnClick` | Exporta la lista con `GenerarConsultas` | CAP-5 | `:176-200`, `.form.txt:1053-1091` |
| `Nombre` | texto bloqueado, no enlazado | — | Muestra `NombreParaCombo` del equipo | CAP-5 | `:216`, `.form.txt:1093-1122` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «GESTIÓN DE LAS CALIBRACIONES»; `lblTitulo1` pasa a `Caption` del formulario | CAP-5 | `:215`, `.form.txt:875-918` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra el formulario | CAP-5 | `:63-65`, `.form.txt:1135-1160` |

Además de los eventos de controles: `Form_Load` (`:202-233`), `m_FormCalibracion_Alta` (`:283-306`) y `m_FormCalibracion_Edicion` (`:308-331`). El `RowSource` de diseño de `ListaFiltrados` trae filas de ejemplo del diseño (`.form.txt:948`); en ejecución se sustituye por los encabezados (`:27`).

## Flujos

### Apertura (`Form_Load`, `:202-233`)

1. `AjustarTamaño Me` (`:210`). Sin equipo activo: «No se conoce el equipo» (`:211-214`).
2. Copia el título a `Caption` y `NombreParaCombo` a `Nombre` (`:215-216`; `EquipoMedida.cls:166-176`).
3. `Filtrar` (`:7-60`): reinicia el estado (sin selección, Alta habilitado, resto deshabilitado; `:19-23`), pone los encabezados (`:27`), lee `m_ObjEquipoMedidaActivo.Calibraciones` (`:28`) y añade una fila por calibración con `IDCalibracion;FechaCalibracion;FechaFinCalibracion` (`:36-40`). Si no hay calibraciones sale sin habilitar la exportación (`:33-35`); si las hay, la habilita (`:43`).
4. Con una única fila (`ListCount = 2` por el encabezado) la selecciona y lanza el clic (`:44-47`); si no, lanza igualmente `ListaFiltrados_Click` (`:49`).
5. Error: `MsgBox` crítico y cierre del formulario (`:226-233`).

### Selección (`ListaFiltrados_Click`, `:235-280`)

Habilita solo Alta (`:242-244`). Sin ID en la columna 0 termina (`:247-250`). Si no hay calibración cargada, o su ID difiere, la relee con `Constructor.getEquipoMedidaCalibracion` (`:251-263`); con calibración, habilita Editar y Eliminar (`:268-269`).

### Alta y edición (`:67-131`)

Cursor de espera; el alta pone la global de calibración a `Nothing`; la edición exige selección («Se ha de seleccionar la calibración de la lista primero», `:102-112`) y la fija en la global (`:115`). Ambas cierran `FormEquipoMedidaCalibracion` si estaba abierto sin guardar (`:77-79, 116-118`), lo abren y enganchan `m_FormCalibracion` (`:80-81, 119-120`). Al llegar `Alta` o `Edicion`, se anula la caché `Calibraciones` del equipo y se vuelve a ejecutar `Filtrar` (`:290-291, 315-316`).

### Eliminación (`ComandoEliminar_Click`, `:133-173`)

Resuelve la selección como en la edición (`:140-151`), ejecuta `m_CalibracionSeleccionada.Eliminar` (`:152`), limpia la selección y repinta con `Filtrar` (`:156-161`). No pide confirmación ni anula la caché `Calibraciones` antes de repintar (`:152-161`).

### Exportación (`ComandoExportarExcel_Click`, `:176-200`)

`GenerarConsultas("Lista de Calibraciones", Me.hWnd, , Me.ListaFiltrados, "No")` (`:184`; `FG.bas:1384-1390`). Si el resultado contiene `|`, se muestra el detalle como error (`:185-189`).

### Errores

Todos los handlers: error 1000 = mensaje ya preparado; otro número se compone con número, descripción y origen; `MsgBox` crítico (`:86-91`, análogo en `:125-130, 167-172, 194-199, 274-279`).

## Reglas, validaciones y cálculos

- La edición y la eliminación exigen una calibración seleccionada (`:102-112, 140-151`).
- Botones: Alta siempre habilitado tras `Filtrar`/clic; Editar y Eliminar solo con calibración cargada; exportar solo si hay calibraciones (`:20-23, 43, 242-244, 268-269`).
- Formato de fila: fechas tal como las devuelve la propiedad, sin formateo explícito (`:39`).
- Eliminación: `Delete` por `IDCalibracion`; exige ID conocido («No se conoce la IDCalibración», `EquipoMedidaCalibracion.cls:176-179`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de calibraciones del equipo | `TbEquiposMedidaCalibraciones` | SQL dinámico concatenado (`IDEquipoMedida=`) | — | `Constructor.bas:5063-5125`, `EquipoMedida.cls:208-225` |
| Lectura de una calibración | `TbEquiposMedidaCalibraciones` | SQL dinámico concatenado (`IDCalibracion=`) | — | `Constructor.bas:5020-5061` |
| Eliminación | `TbEquiposMedidaCalibraciones` | SQL dinámico concatenado, `Execute` | no | `EquipoMedidaCalibracion.cls:180-183` |
| Exportación a Excel | — | `GenerarConsultas` sobre el contenido de la lista | — | `:184` |

## Coexistencia

A adaptar a PostgreSQL vinculado: `Delete * FROM` (sintaxis Jet, `EquipoMedidaCalibracion.cls:180-182`) no es SQL estándar; SQL concatenado con IDs numéricos sin parámetros; `Nz(..., "")` aplicado a todos los campos, incluidas las fechas (`Constructor.bas:5046, 5107`), con lo que una fecha nula pasa a cadena vacía; el diccionario usa `TextCompare` (`Constructor.bas:5098`); el borrado se ejecuta fuera de transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoEliminar_Click` no anula `m_ObjEquipoMedidaActivo.Calibraciones` antes de `Filtrar` (`:152-161`; a diferencia de `:290, 315`): la lista podría seguir mostrando la calibración borrada si la propiedad devuelve su caché (`EquipoMedida.cls:213-215`).
- El borrado no pide confirmación (`:152`).
- `Filtrar` con equipo sin calibraciones sale en `:33-35` sin lanzar `ListaFiltrados_Click`; los botones quedan en el estado fijado en `:20-23`, coherente, pero `Nombre`/lista no se limpian de una selección anterior salvo por `:19`.
- Los errores de `Filtrar` con error 1000 se propagan vía `p_Error` pero el manejador no los vuelve a lanzar (`:54-58`); el llamador sí lo comprueba (`:217-220, 291-294`).
- El rótulo de `ComandoExportarExcel` es «Comando227» en diseño (`.form.txt:1053-1091`).
- La ficha `Eliminar` se declara `Function ... As String` sin asignar valor de retorno (`EquipoMedidaCalibracion.cls:169`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: confirmación antes de eliminar; consulta con parámetros; filtrado y ordenación de la lista por fecha.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir desde un equipo, el título muestra su nombre y la lista sus calibraciones con fecha inicial y final (`:215-216, 36-40`).
2. Sin equipo activo se muestra «No se conoce el equipo» y se cierra el formulario (`:211-214, 232`).
3. Sin selección, Editar y Eliminar están deshabilitados; con selección se habilitan (`:242-244, 268-269`).
4. Alta y edición reflejan el cambio en la lista al guardar la calibración (`:283-331`).
5. Eliminar quita la calibración de la base y de la lista (`:152-161`).
6. Exportar genera el Excel «Lista de Calibraciones» solo si hay filas (`:43, 184`).

## Pendientes

- Contenido de `AjustarTamaño` (`FG.bas:261`) y de `GenerarConsultas` más allá de la firma (`FG.bas:1384`).
- Si `TbEquiposMedidaCalibraciones` tiene clave ajena o restricciones al borrar (ERD de origen).
- Orden de las filas: depende del orden de la consulta, sin `ORDER BY` (`Constructor.bas:5086-5088`).
