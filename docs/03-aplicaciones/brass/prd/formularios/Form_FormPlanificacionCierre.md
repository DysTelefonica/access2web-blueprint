# Form_FormPlanificacionCierre

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionCierre.cls` (242 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Cierre de una planificación (`TbPlanificacion`) asociándole el trabajo que la cumple: un evento de tipo `MANT. PROGRAMADO` ya cerrado, o una planificación registrada (`TbPlanificacionRegistrada`). Es un formulario **no enlazado** (sin `RecordSource`): recibe `IDPlanificacion` en `OpenArgs`, vuelca los datos en controles bloqueados y delega el cierre en `FG.bas:CerrarPlanificacionConTrabajos` (`FG.bas` = `modules/Funciones Generales.bas`). Las citas son `Form_FormPlanificacionCierre.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionesBusqueda` | Cierra el formulario si estaba abierto y lo abre con la planificación seleccionada (dos puntos de entrada) | `Form_FormPlanificacionesBusqueda.cls:412-415, 612-615` |
| `Form_FormEventoEdicion` | Lo abre con `strIDMantenimiento` si es numérico | `Form_FormEventoEdicion.cls:925-937` |
| El propio formulario (`ComandoVerProgramacion`) | Se cierra y se reabre con otra planificación | `:154-157` |

`OpenArgs` debe ser numérico; si no, «No se ha abierto el formulario con los parámetros adecuados» y cierre (`:205-209, 227-233`). Tras cerrar con éxito se cierra solo (`:28`); no refresca al formulario que lo abrió. `cmdSalir` cierra sin guardar (`:40-43`).

## Controles y eventos

Controles: 43 `Begin` en el `.form.txt` = 1 `Form` + 1 `Section` (`Detalle`) + 7 plantillas por defecto (`.form.txt:169-293`) + 34 controles reales (tabla). Cuadre: 1 + 1 + 7 + 34 = 43.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Form` | formulario | `OnLoad` → `Form_Load` | Guarda `OpenArgs` en `strIDPlanificacion`, llama a `EstablecerDatos` | CAP-6 | `:195-234`, `.form.txt:110` |
| `ComandoCierre` («Cierre») | botón, deshabilitado de diseño | `OnClick` | Llama a `CerrarPlanificacionConTrabajos` con los combos de evento y no-evento y cierra | CAP-6 | `:5-39`, `.form.txt:350-362` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-6 | `:40-43`, `.form.txt:386-388` |
| `ComandoVerProgramacion` («Ver Prog.») | botón, deshabilitado de diseño | `OnClick` | Reabre el formulario con la planificación elegida en `ComboIDEventosUsados` (columna 1) | CAP-6 | `:146-168`, `.form.txt:1123, 1131-1133` |
| `ComandoAyudaFormulario` | botón (rótulo «Comando46») | — (sin evento) | Sin handler | sin asignar | `.form.txt:1157-1158` |
| `ComboIDEvento` | combo | `OnExit` → `ComboIDEvento_Exit` | Eventos `MANT. PROGRAMADO` no usados; al salir habilita `ComandoCierre` si hay alguno de los dos combos informado | CAP-6 | `:113-121, 169-176`, `.form.txt:902-911` |
| `ComboIDNoEvento` | combo | `OnExit` → `ComboIDNoEvento_Exit` | Planificaciones registradas del equipo; misma regla de habilitación | CAP-6 | `:122-126, 187-194`, `.form.txt:973-980` |
| `ComboIDEventosUsados` | combo | `OnChange` → `ComboIDEventosUsados_Change` | Eventos ya usados en otras planificaciones del equipo; habilita `ComandoVerProgramacion` si la columna 1 es numérica | CAP-6 | `:128-133, 177-186`, `.form.txt:1099-1108` |
| `BUI`, `SubSistema`, `EQUIPO` | 3 textos bloqueados y deshabilitados | — | Se asignan antes de leerse (ver defectos) | CAP-6 | `:54-56`, `.form.txt:503-518, 534-549, 592-607` |
| `AÑO`, `Mes`, `Semana`, `FechaPrevistaCierre`, `FechaCierre` | 5 textos bloqueados | — | Datos de la planificación; `FechaPrevistaCierre` y `FechaCierre` con formato `dd/mm/yyyy` | CAP-6 | `:83-95, 103-106`, `.form.txt:664, 749, 781, 839, 1035` |
| `lblFechaCierre` + `FechaCierre` | etiqueta y texto | — | Visibles solo si la planificación ya tiene fecha de cierre | CAP-6 | `:103-112` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | sin asignar | `.form.txt:313, 336` |
| `Marco175`, `Marco189`, `Marco3` | 3 grupos de opciones | — | Marcos visuales | sin asignar | `.form.txt:414, 436, 1265` |
| 14 etiquetas (`Etiqueta191, 192, 197, 242, 245, 249, 258, 261, 264, 0`, `lblFechaCierre`, `lblTitulo1`, `lblTitulo2`, `lblListaCandidatosEventos`) | etiquetas | — | Rótulos; `lblTitulo1` alimenta `Me.Caption` | sin asignar | `:204`, `.form.txt:1217, 1242` |

Desglose de los 34: 4 botones, 3 combos, 8 textos, 2 rectángulos, 3 marcos y 14 etiquetas (`lblFechaCierre` cuenta una vez, entre las etiquetas).

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:44-145`)
1. Lee la planificación (`Anio`, `Mes`, `Semana`, `FechaCierre`, `FechaPrevistaCierre`, `IDEvento`, `IDNoEvento`, `IDEquipo`) y exige año, mes y semana numéricos (`:57-93`).
2. Vuelca fecha prevista y, si existen, el evento o no-evento ya asociados en los combos (`:94-102`).
3. Con `FechaCierre` válida muestra `lblFechaCierre`/`FechaCierre` y deshabilita `ComandoCierre`; si no, los oculta y lo habilita (`:103-112`).
4. Reasigna el `RowSource` de los tres combos (`:113-133`) y **deshabilita `ComandoCierre` siempre** (`:127`).

### Cierre (`ComandoCierre_Click`, `:5-39`)
1. `strIDPlanificacion` debe ser numérico (`:11-14`). Toma `ComboIDEvento` y `ComboIDNoEvento` (`:15-16`).
2. `CerrarPlanificacionConTrabajos` (`:22`): ver reglas. Un resultado con `\|` es error (`:23-27`).
3. Éxito: cierra el formulario; el texto descriptivo devuelto no se muestra (`:28`).
4. Error: `MsgBox` crítico y el formulario sigue abierto (`:33-38`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Planificación numérica | función | `FG.bas:3223-3226` |
| Debe informarse uno y solo uno de evento / no-evento | función | `FG.bas:3227-3236` |
| Evento: no usado en otra planificación (`Dame` sobre `TbPlanificacion.IDEvento`) | función | `FG.bas:3243-3252` |
| Evento: `TIPOEVENTO` debe ser `MANT. PROGRAMADO` | función | `FG.bas:3258-3267` |
| Evento: debe tener fecha fin (`DameFechaFinEvento` en `FG.bas:798`), que pasa a ser `FechaCierre` | función | `FG.bas:3273-3282` |
| No-evento: numérico, no usado en otra planificación y registrado en `TbPlanificacionRegistrada` para el equipo de la planificación; su `FechaMantoRealizado` pasa a ser `FechaCierre` | función | `FG.bas:3284-3333` |
| Cierre: `MotivoCierre = "Trabajo Realizado"`, `FechaCierre` y `IDEvento` o `IDNoEvento` | función | `FG.bas:3341, 3351-3359` |
| `ComandoCierre` se habilita al salir de cualquiera de los dos combos con valor | formulario | `:169-176, 187-194` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos de la planificación | `TbPlanificacion`, `TbEquipos` | DAO sobre SQL concatenado (`IDPlanificacion`) | — | `:57-62` |
| Eventos candidatos (`DameFechaFinEvento` como columna) | `TbEventos`, `TbPlanificacion`, `TbEquipos` | SQL dinámico concatenado como `RowSource` | — | `:113-121` |
| Planificaciones registradas del equipo | `TbPlanificacionRegistrada`, `TbPlanificacion` | SQL dinámico concatenado | — | `:122-126` |
| Eventos usados | `TbEventos`, `TbPlanificacion` | SQL dinámico concatenado | — | `:128-133` |
| Comprobaciones de uso y tipo | `TbPlanificacion`, `TbEventos` | función `Dame` (implementación no leída) | — | `FG.bas:3243, 3258` |
| Registro del no-evento | `TbPlanificacionRegistrada` | DAO `OpenRecordset` sobre SQL concatenado | no | `FG.bas:3319-3333` |
| Cierre | `TbPlanificacion` | DAO `Edit`/`Update` | no | `FG.bas:3342-3359` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DameFechaFinEvento` es función VBA invocada desde una consulta del `RowSource` (`:113`), sin equivalente en PostgreSQL vinculado; SQL concatenado con textos entre comillas; fechas tratadas como cadenas con `Format(..., "dd/mm/yyyy")` (`:95, 106`) y asignadas a campos fecha (`FG.bas:3354`); cierre en un `Edit`/`Update` sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strBUI`, `strSUBSISTEMA` y `strEquipo` se vuelcan en los controles (`:54-56`) y se usan en el `WHERE` de `ComboIDEvento` (`:116-118`) sin que se les asigne valor en `EstablecerDatos` (`:46-48, 57-112`): los textos quedarían vacíos y el combo de eventos candidatos no devolvería filas.
- `ComandoCierre.Enabled = False` incondicional (`:127`) anula la habilitación de `:111`; la rama de la planificación aún abierta solo se habilita tras `Exit` de un combo (`:169-176`).
- El mensaje de éxito de la función (`FG.bas:3363`) no llega al usuario (`:22-28`).
- La función admite un solo trabajo (`FG.bas:3232-3236`), pero el formulario permite informar ambos combos y solo lo detecta al guardar.
- Mensajes de la función hablan de «IDActividad» donde el parámetro es el no-evento (`FG.bas:3227-3231`) y de «Sismeta» (`FG.bas:3344`).
- `RowSource` de diseño de los combos (`.form.txt:904-909, 975-978, 1101-1106`) referencia `TbIDEquipos` y `TbPlanificacionAñadida` con valores fijos; el código los sustituye (`:120, 125, 132`), pero el diseño no refleja el origen real.
- Un apóstrofo en BUI, subsistema o equipo rompería el SQL concatenado (`:116-118`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cierre en una sola transacción; exclusión mutua evento/no-evento en la interfaz; mostrar el resultado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con un identificador no numérico muestra el error y cierra (`:205-209`).
2. Una planificación ya cerrada muestra su fecha de cierre y no permite cerrar de nuevo (`:103-107`).
3. Cerrar sin evento ni no-evento muestra el error de la función y no escribe (`FG.bas:3227-3231`).
4. Cerrar con ambos informados se rechaza (`FG.bas:3232-3236`).
5. Un evento ya usado o de otro tipo se rechaza (`FG.bas:3252, 3267`).
6. Un evento válido deja `FechaCierre` igual a su fecha fin y `MotivoCierre` «Trabajo Realizado» (`FG.bas:3273-3282, 3351-3359`).
7. «Ver Prog.» se habilita solo con una planificación numérica en `ComboIDEventosUsados` (`:177-186`).

## Pendientes

- Implementación de `Dame` (`FG.bas:3243, 3258, 3293, 3308`) y de `AjustarTamaño` (`:202`).
- Si `AllowEdits`/`Locked` impiden cambiar los textos bloqueados y `Enabled` los tiñe en gris (`.form.txt:504-505`).
- Tipos y claves de `TbPlanificacion` y `TbPlanificacionRegistrada` (ERD de origen).
