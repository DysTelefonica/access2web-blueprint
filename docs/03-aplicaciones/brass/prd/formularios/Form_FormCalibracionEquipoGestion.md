# Form_FormCalibracionEquipoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionEquipoGestion.cls` (800 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de los equipos calibrables (`TbEquiposCalibrables`): lista filtrable por BUI, equipo, criterio de calibración y ámbito (vigentes/obsoletos), con la vigencia de la última calibración de cada equipo, y acceso a alta, edición, borrado, calibraciones y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena asignando un SQL dinámico a `ListaFiltrados.RowSource` (`:62`). Las citas son `Form_FormCalibracionEquipoGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todo el formulario: CAP-5 Equipos, equipos de medida y calibraciones.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | Clic en la etiqueta `lblCalibraciones`: `DoCmd.OpenForm "FormCalibracionEquipoGestion"` | `Form_Form0BDOpcionesOtras.cls:44-46` |
| `Form_FormCalibracionEquipoAlta`, `Form_FormCalibracionEquipoEdicion`, `Form_FormCalibracionAlta`, `Form_FormCalibracionGestion` | Si está abierto, llaman a su `ActualizarLista` pública tras guardar | `Form_FormCalibracionEquipoAlta.cls:33-35`; `Form_FormCalibracionEquipoEdicion.cls:158-160`; `Form_FormCalibracionAlta.cls:34-36`; `Form_FormCalibracionGestion.cls:225-227` |

Abre `FormCalibracionEquipoAlta` sin argumentos (`:138-141`), `FormCalibracionEquipoEdicion` con el ID del equipo en `OpenArgs` (`:160-163`) y `FormCalibracionGestion` con `ID|texto de estado` (`:111-117`); en los tres casos cierra antes la instancia abierta. Métodos públicos para otros formularios: `ActualizarLista` (`:33`), `ComandoActualizarTodo_Click` (`:12`), `LimpiarUltimos` (`:6`) y `ListaFiltrados_Click` (`:530`). Sale con `cmdSalir` (`:526-529`).

## Controles y eventos

Controles: 37 (37 Begin en el .form.txt: 29 en la sección `Detalle` — 12 botones, 4 combos, 1 lista, 7 etiquetas, 3 rectángulos, 2 grupos de opciones —, más `Form`, `Section` y 6 plantillas de controles predeterminados).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista de 6 columnas con encabezados (ID oculto) | `OnClick` → `ListaFiltrados_Click` | Guarda el ID, habilita editar/eliminar/calibraciones y rotula el estado | `:530-598`; `Form_FormCalibracionEquipoGestion.form.txt:934-953` |
| `ComboBUI` (código: `ComboBui`) | combo `Table/Query`, `LimitToList` | `OnChange`, `OnExit` | Filtra por BUI exacto | `:374-418`; `Form_FormCalibracionEquipoGestion.form.txt:1176-1182` |
| `ComboEquipos` | combo de 2 columnas (equipo, BUI) | `OnChange`, `OnExit` | Filtra por texto contenido en `Equipo` | `:466-510`; `Form_FormCalibracionEquipoGestion.form.txt:1243-1249` |
| `ComboCalibracion` | combo `Value List`, `LimitToList` | `OnChange`, `OnExit` | Criterio: con, sin, con vigente, con no vigente | `:420-464`; `Form_FormCalibracionEquipoGestion.form.txt:1410-1416` |
| `ComboAmbito` | combo `Value List`: Todos, Vigentes, Obsoletos | `OnChange`, `OnExit` | Filtra por fecha de baja nula o no | `:328-372`; `Form_FormCalibracionEquipoGestion.form.txt:1477-1482` |
| `ComandoLimpiarBUI`, `ComandoLimpiarEquipo`, `ComandoLimpiarCalibracion`, `ComandoLimpiarAmbito` | botones | `OnClick` | Vacían su combo y refrescan si tenía valor | `:268-327`; `Form_FormCalibracionEquipoGestion.form.txt:1716-2219` |
| `ComandoActualizarTodo` | botón | `OnClick` (público) | Ejecuta los cuatro «limpiar» | `:12-32`; `Form_FormCalibracionEquipoGestion.form.txt:1585-1587` |
| `ComandoAltaEquipo` («Alta Equipo») | botón | `OnClick` | Abre `FormCalibracionEquipoAlta` | `:129-142`; `Form_FormCalibracionEquipoGestion.form.txt:1008-1010` |
| `ComandoEditarEquipo` («Editar Equipo») | botón, deshabilitado de diseño | `OnClick` | Abre `FormCalibracionEquipoEdicion` | `:143-174`; `Form_FormCalibracionEquipoGestion.form.txt:972, 981-983` |
| `ComandoEliminarEquipo` («Eliminar Equipo») | botón, deshabilitado de diseño | `OnClick` | Confirma y borra el equipo | `:175-244`; `Form_FormCalibracionEquipoGestion.form.txt:1027, 1037-1039` |
| `ComandoCalibraciones` («Calibraciones (n)») | botón, deshabilitado de diseño | `OnClick` | Abre la gestión de calibraciones del equipo | `:95-128`; `Form_FormCalibracionEquipoGestion.form.txt:1265, 1274-1276` |
| `ComandoExportarExcel` | botón con imagen, deshabilitado de diseño | `OnClick` | Exporta el SQL de la lista a un `.xlsx` | `:245-267`; `Form_FormCalibracionEquipoGestion.form.txt:1325, 1334-1336` |
| `lblCalibraciones` | etiqueta | — | Texto de estado de calibración del equipo seleccionado | `:563-568`; `Form_FormCalibracionEquipoGestion.form.txt:1303-1304` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:526-529` |
| `ComandoAyudaFormulario` | botón con imagen | — | Sin evento en el `.form.txt` | `Form_FormCalibracionEquipoGestion.form.txt:1505-1506` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta142`, `Etiqueta162`, `Etiqueta187`, `Etiqueta0` | etiquetas | — | Título (`lblTitulo1` pasa a `Caption`) y rótulos BUI, EQUIPOS, CALIBRACIÓN, ÁMBITO | `:518`; `Form_FormCalibracionEquipoGestion.form.txt:1142-1565` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2`, `Marco120`, `Marco122` | rectángulos y grupos de opciones | — | Decoración | `Form_FormCalibracionEquipoGestion.form.txt:868, 891, 914, 1095, 1118` |
| `Form`, `Detalle` y 6 plantillas | formulario, sección, predeterminados | `OnOpen` → `Form_Open` | Ver Flujos | `:512-525`; `Form_FormCalibracionEquipoGestion.form.txt:57` |

## Flujos

### Apertura (`Form_Open`, `:512-525`)

Con `On Error Resume Next` (`:513`): `AjustarTamaño Me`, `Caption` = `lblTitulo1` y `ActualizarLista` (`:517-519`).

### Refresco (`ActualizarLista`, `:33-94`)

1. Deshabilita editar, eliminar, calibraciones y exportar, y oculta `lblCalibraciones` (`:43-47`).
2. `Filtrar` construye el SQL con los cuatro combos (`:53-60`); se asigna a `RowSource` y se hace `Requery` (`:61-63`).
3. Sin filas, exportar sigue deshabilitado; con una sola fila la selecciona y lanza `ListaFiltrados_Click` (`:66-72`); con varias, busca la fila ya seleccionada y la relanza (`:74-80`). Con filas habilita exportar (`:82`).

### Selección (`ListaFiltrados_Click`, `:530-598`)

Lee el ID de la columna 0, habilita los tres botones (`:536-550`), obtiene el texto con `DameTextoCalibracion` y lo muestra en `lblCalibraciones` (`:556-568`), cuenta las fechas de calibración con `DameNumeroRegistrosPorSQL` y rotula el botón «Calibraciones (n)» (`:569-587`).

### Eliminar (`ComandoEliminarEquipo_Click`, `:175-244`)

1. Cuenta las calibraciones del equipo en `TbEquiposCalibrablesFechas` (`:185-199`).
2. Pregunta con texto distinto para 0, 1 o varias, con «No» por defecto (`:204-211`).
3. Con «Sí», `EliminarEquipoACalibrar` abre el registro y hace `.Delete` (`:226, 776-785`) y refresca (`:232-233`).

### Exportar (`ComandoExportarExcel_Click`, `:245-267`)

`GenerarConsultasSQL` recibe el `RowSource` actual (`:251-252`); crea un `.xlsx` en el directorio local configurado, lo rellena por Excel y lo abre (`FG.bas:1606-1608, 1625-1670`). Errores controlados: «La consulta no contiene registros» (`FG.bas:1620-1623`).

## Reglas, validaciones y cálculos

- **Filtro BUI** igual exacto; **Equipo** `Like '*texto*'`; vacíos no filtran (`:625-635`).
- **Ámbito:** «Vigentes» = `FechaBajaParaCalibracion Is Null`; «Obsoletos» = no nula; «Todos» o vacío, sin filtro (`:636-642`).
- **Criterio de calibración** sobre `DameDiasVigenciaCalibracion` (`FG.bas:1247-1292`): «SC» si no hay calibración, si no días entre hoy y la última `FechaFinCalibrado` (negativo = no vigente). Con = distinto de «SC»; Sin = «SC»; no vigente = empieza por «-»; vigente = no empieza por «s» ni «-» (`:646-657`).
- **Texto de estado** (`DameTextoCalibracion`, `:676-752`): con fecha de baja, «dado de baja de calibración desde <fecha>»; sin calibraciones, «sin ninguna calibración registrada»; si no, días fuera de vigencia, vigente hasta hoy o vigente dentro de N días (`:713-739`).
- **Refresco por combo:** `Change` y `Exit` solo refrescan si el valor difiere del anterior guardado (`:335-338, 358-361`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista filtrada | `TbEquiposCalibrables` + `TbEquiposCalibrablesFechas` | SQL dinámico concatenado (`SELECT DISTINCT`, `LEFT JOIN`, función VBA `DameDiasVigenciaCalibracion` en columna y `WHERE`) | — | `:620-667` |
| Vigencia por equipo | `TbEquiposCalibrablesFechas` | SQL dinámico con ID concatenado, `ORDER BY ... DESC` | — | `FG.bas:1268-1272` |
| Texto de estado | ambas | `Dame` (`FG.bas:290-312`) y `SELECT Last(FechaFinCalibrado)` concatenado | — | `:706, 718-721` |
| Recuento de calibraciones | `TbEquiposCalibrablesFechas` | SQL dinámico con ID concatenado | — | `:185-187, 569-571` |
| Borrado de equipo | `TbEquiposCalibrables` | DAO `.Delete` | no | `:776-785` |
| Origen de combos | `TbBUI`, `TbEquiposCalibrables` | `RowSource` de diseño | — | `Form_FormCalibracionEquipoGestion.form.txt:1178, 1245` |
| Exportación | resultado del SQL de la lista | DAO + automatización de Excel | — | `FG.bas:1616-1670` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `Like '*'` y `Is Null` en dialecto Jet (`:626-641`); funciones VBA dentro de `SELECT`/`WHERE` (`:622, 644-659`) que no se pueden delegar al servidor PostgreSQL y obligarían a leer la tabla en cliente; comparaciones de texto que dependen de que Jet no distinga mayúsculas (`'s'` frente a «SC», `:655`); `Last()` sin orden (`:718`); `RecordCount` tras `MoveLast` (`FG.bas:136-141`); borrado DAO sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El aviso de borrado dice «borrar al equipo junto con su calibración» (`:207, 209`), pero el código solo borra la fila del equipo (`:776-785`); depende de integridad referencial en cascada (no visible en `src`).
- `ComandoActualizarTodo` no refresca si los cuatro combos ya están vacíos: cada «limpiar» sale sin llamar a `ActualizarLista` (`:18-21, 271-273`).
- `Filtrar` concatena `BUI` y `Equipo` sin escapar (`:628, 634`); un apóstrofo rompería el SQL.
- Si `strAmbito` no fuera «Todos», «Vigentes» ni «Obsoletos», `strParteWhereAmbito` queda vacío y el `WHERE` es inválido (`:636-666`); lo mitiga `LimitToList` (`Form_FormCalibracionEquipoGestion.form.txt:1459`).
- `DameTextoCalibracion` usa `Last()` sin `ORDER BY` (`:718`), mientras que la vigencia ordena por fecha descendente (`FG.bas:1268-1276`): el texto y la columna `Vigencia` podrían discrepar.
- `ActualizarLista` calcula `strIDEquipoSeleccionadoActualmente` pero no lo usa para reseleccionar (`:42, 70, 76`); tras un `Requery` la selección podría perderse.
- Cada fila ejecuta `DameDiasVigenciaCalibracion` una vez en columna y hasta otra en el `WHERE` (`:622, 644-659`): coste por fila.
- `On Error Resume Next` en `Form_Open`, alta y los cuatro «limpiar» (`:130, 269, 284, 299, 314, 513`) oculta errores.
- El código del formulario no consulta permisos de usuario (`:1-800`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: vigencia calculada en una vista o consulta única; filtros parametrizados; borrado con recuento y cascada explícitos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir se lista el conjunto sin filtros, con editar, eliminar y calibraciones deshabilitados (`:43-47, 519`).
2. Cada combo filtra la lista y su botón de limpiar la restablece (`:625-642, 268-327`).
3. «Con/Sin Calibración» y «vigente/no vigente» filtran según `DameDiasVigenciaCalibracion` (`:646-657`).
4. Seleccionar un equipo habilita los tres botones y muestra el texto de estado y «Calibraciones (n)» (`:548-587`).
5. Eliminar pide confirmación con el número de calibraciones asociadas y, con «Sí», quita el equipo y refresca (`:204-233`).
6. Exportar genera un Excel con las columnas de la lista; sin filas, avisa (`FG.bas:1620-1623`).
7. Tras alta o edición de un equipo, la lista abierta se refresca (`Form_FormCalibracionEquipoEdicion.cls:158-160`).

## Pendientes

- Relación y borrado en cascada entre `TbEquiposCalibrables` y `TbEquiposCalibrablesFechas` (ERD de origen).
- Qué hace `Ejecutar 1, "open", ...` (`FG.bas:1670`) y su comportamiento en la web.
- Destino de `OpenArgs` con `ID|texto` en `FormCalibracionGestion` (`:117`).
- `RowSource` completo de diseño de `ListaFiltrados` y `ComboEquipos` (`Form_FormCalibracionEquipoGestion.form.txt:945, 1245`).
