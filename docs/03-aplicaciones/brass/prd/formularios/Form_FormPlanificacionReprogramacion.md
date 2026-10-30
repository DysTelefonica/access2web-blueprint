# Form_FormPlanificacionReprogramacion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionReprogramacion.cls` (278 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Reprograma una planificación abierta: la cierra con motivo «Reprogramación» y crea otra para el mismo equipo en el año, mes y semana elegidos. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe `IDPlanificacion` por `OpenArgs`, lee los datos con DAO y delega el cambio en `CerrarPlanificacionSinTrabajos` (`FG.bas:2982-3190`). Las citas son `Form_FormPlanificacionReprogramacion.cls:línea` (aquí `:línea`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionesBusqueda` (`ComandoReprogramar_Click`) | Exige planificación seleccionada, cierra el formulario si estaba abierto y lo abre con `OpenArgs` = `IDPlanificacion` | `Form_FormPlanificacionesBusqueda.cls:424-435` |

Salida: tras reprogramar refresca `ListaFiltrados` de `FormPlanificacionesBusqueda` si está abierto y se cierra sin guardar (`:109-112`). `cmdSalir` cierra sin más (`:124-127`). No emite eventos ni deja globales.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `EQUIPO`, `SUBSISTEMA`, `BUI` | textos (3), deshabilitados y bloqueados | — | Identifican el equipo; se rellenan en `EstablecerDatos` | CAP-6 Planificación | `:167-177`, `.form.txt:715-717, 752-753, 821-822` |
| `AÑO`, `Mes`, `Semana` | textos (3) bloqueados | — | Planificación actual | CAP-6 Planificación | `:182-192`, `.form.txt:890, 1022, 1058` |
| `FechaPrevistaCierre` | texto bloqueado, `Short Date` | — | Fecha prevista actual (`dd/mm/yyyy`) si es fecha | CAP-6 Planificación | `:193-195`, `.form.txt:1128, 1142` |
| `MotivoReplanificacion` | texto, `Tag` OBLIGATORIO | `OnExit` → `MotivoReplanificacion_Exit` | Motivo; recalcula la botonera | CAP-6 Planificación | `:269-278`, `.form.txt:1205-1207` |
| `AñoNuevo` | texto, `Tag` OBLIGATORIO | `OnExit` → `AñoNuevo_Exit` | Año nuevo; recarga semanas y botonera | CAP-6 Planificación | `:65-69`, `.form.txt:1273-1275` |
| `ComboMesNuevo` | combo lista 1-12, `LimitToList`, OBLIGATORIO | `OnChange`, `OnExit` | Mes nuevo; recarga semanas (y botonera al salir) | CAP-6 Planificación | `:209-229`, `.form.txt:1329, 1348-1354` |
| `ComboSemanaNueva` | combo lista de valores, `LimitToList`, OBLIGATORIO | `OnExit` → `ComboSemanaNueva_Exit` | Semana nueva; su lista se reconstruye con año y mes | CAP-6 Planificación | `:22-64, 230-239`, `.form.txt:1407, 1427-1431` |
| `ComandoReplanificar` | botón «Replanificar», deshabilitado en diseño | `OnClick` → `ComandoReplanificar_Click` | Valida y reprograma | CAP-6 Planificación | `:70-123`, `.form.txt:375-386` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra | CAP-6 Planificación | `:124-127`, `.form.txt:427-429` |
| `lblTitulo` y 15 `Etiqueta…` | etiquetas (16) | — | Rótulos; `lblTitulo` da el `Caption` de la ventana | CAP-6 Planificación | `:247`, `.form.txt:330, 555-1546` |
| `Cuadro69`, `CuadroSuperior`, `Cuadro2`, `Marco175`, `Marco189`, `Imagen116`, `Imagen255`, `Imagen283` | rectángulos, marcos e imágenes (8) | — | Decoración | sin asignar | `.form.txt:284-360, 467, 583-604, 1456, 1572` |

Controles: 46 (46 Begin en el .form.txt): 1 `Form`, 7 plantillas de tipo por defecto (`.form.txt:138-241`), 1 `Section` y 37 controles con nombre = 7 + 2 + 2 + 2 + 16 + 8 de la tabla.

## Flujos

### Apertura (`Form_Load`, `:240-268`)

1. `AjustarTamaño Me`, `Caption` desde `lblTitulo` y `strIDPlanificacion = OpenArgs`, que debe ser numérico (`:246-253`); si no, «No se ha abierto el formulario con los parámetros adecuados», aviso y cierre (`:261-267`).
2. `EstablecerDatos` (`:128-208`) lee equipo y planificación (`TbPlanificacion` ⋈ `TbEquipos`, `:140-145`), exige `BUI`, `SUBSISTEMA`, `Equipo`, `Anio`, `Mes` y `Semana` no vacíos o numéricos (`:163-192`) y deja el botón deshabilitado (`:196`). Devuelve `#ERR|…` en fallo.

### Reprogramación (`ComandoReplanificar_Click`, `:70-123`)

1. Toma motivo, año, mes y semana; la semana debe tener al menos 3 caracteres y empezar por 2 dígitos, de los que usa los 2 primeros (`:80-92`).
2. `CerrarPlanificacionSinTrabajos(id, motivo, semana, mes, año)` (`:98`). Un retorno con `\|` es un error y se muestra (`:99-103`, `:122`).
3. Refresca la búsqueda y cierra (`:109-112`). Con error permanece abierto.

### Lista de semanas (`EstablecerListaSemana`, `:22-64`)

Con año y mes numéricos llama a `FechaViernes(mes, año)`, que devuelve «NN fecha» por semana con viernes (`FG.bas:589-680`); sustituye los saltos de línea por `;` y lo asigna como `RowSource` (`:41-55`). Si no son numéricos asigna `#ERR` (`:28-30`). Errores: aviso modal (`:59-63`).

## Reglas, validaciones y cálculos

- **Botonera:** `ComandoReplanificar` solo se habilita cuando los controles con `Tag` OBLIGATORIO tienen valor (`:5-21`); se evalúa en las salidas de campo (`:65-69, 219-239, 269-278`).
- **Motivo obligatorio** (`FG.bas:3023-3026`).
- **Estado:** la planificación debe estar «Abierta» (`FG.bas:3043-3052`; `EstadoPlanificacion` lee `FechaCierre` y `MotivoCierre`, `FG.bas:2059-2069`).
- **Semana:** entre 1 y las semanas del mes (`FG.bas:3070-3078`) y con viernes (`FG.bas:3089-3098`).
- **Duplicado:** no puede existir otra planificación del equipo en esa semana, mes y año (`FG.bas:3106-3115`).
- **Resultado:** crea la nueva (`FG.bas:3121-3136`), cierra la antigua con `MotivoCierre` «Reprogramación», `FechaCierre` = `Now()`, `MotivacionReprogramacion` e `IDNuevaPlanificacion` (`FG.bas:3138-3150`).
- **Compensación:** si falla tras crear la nueva y antes de cerrar la antigua, borra la nueva (`FG.bas:3180-3186`; `EliminarPlanificacion`, `FG.bas:2966-2971`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos del equipo y planificación | `TbPlanificacion`, `TbEquipos` | SQL dinámico concatenado (`IDPlanificacion`) | — | `:140-145` |
| Estado y equipo | `TbPlanificacion` | SQL dinámico concatenado; `Dame` | — | `FG.bas:2059-2062, 3058` |
| Comprobación de duplicado | `TbPlanificacion` | SQL dinámico concatenado | — | `FG.bas:2210-2217` |
| Alta de la nueva | `TbPlanificacion` | DAO `AddNew`/`Update`; ID por `DameID` | no | `FG.bas:2352, 2372-2381` |
| Cierre de la antigua | `TbPlanificacion` | DAO `Edit`/`Update` | no | `FG.bas:3140-3150` |
| Compensación | `TbPlanificacion` | `DELETE` con `RunSQL` y avisos desactivados | no | `FG.bas:2966-2971` |

## Coexistencia

Adaptar a PostgreSQL vinculado: SQL concatenado con `IDPlanificacion`, `IDEquipo` y semana/mes/año; identificador nuevo calculado en VBA mediante `DameID` y escrito con `AddNew` (`FG.bas:2352, 2373`); `DELETE … FROM` con `DoCmd.RunSQL` y `SetWarnings` (`FG.bas:2966-2971`); fechas como texto `dd/mm/yyyy` y `Now()` en `FechaCierre` (`:194`, `FG.bas:3138`); dos escrituras sin transacción con compensación manual.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El resultado de `EstablecerDatos` no se comprueba (`:254`): un `#ERR|…` dejaría el formulario abierto sin datos y sin aviso.
- `ComboMesNuevo_Change` lee `Me.ComboMesNuevo` mientras se edita (`:209-218`, `:27`): podría usar aún el valor anterior.
- Sin transacción: si la nueva se crea y la antigua no se cierra, la compensación puede fallar y dejar una planificación duplicada (`FG.bas:3121-3150, 3180-3186`).
- El parámetro opcional `strURLAnexo` no se pasa desde este formulario (`:98`), por lo que la rama de anexos de `CerrarPlanificacionSinTrabajos` no se ejecuta desde aquí (`FG.bas:3155-3168`).
- Mensajes imprecisos: «No se ha introducido una válida» (`:85, 89`), «CerrarPlanificacionConTrabajos» en el aviso de otra función (`:101`).
- Si cambia el año y no se sale de `AñoNuevo` antes de pulsar, la lista de semanas podría quedar antigua; sin comprobación al pulsar (`:65-69`, `:76-92`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: operación atómica en el servidor; selector de semana calculado en servidor; exigir motivo antes de habilitar la acción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin identificador numérico, o con una planificación sin equipo o sin año/mes/semana, muestra el aviso y cierra.
2. Equipo, subsistema, BUI, año, mes, semana y fecha prevista actuales se ven pero no se editan.
3. «Replanificar» no se habilita hasta rellenar motivo, año, mes y semana.
4. Al elegir año y mes, la lista de semanas ofrece solo semanas con viernes.
5. Reprogramar a una semana ya planificada para el equipo muestra «La semana … ya está planificada para este equipo.»
6. Reprogramar una planificación no abierta muestra su estado y no escribe.
7. Un resultado correcto cierra la antigua con motivo «Reprogramación», crea la nueva y refresca la búsqueda.

## Pendientes

- Contenido de `NumeroSemanasTieneElMes`, `DameID` y `AnexarAPlanificacion` (`FG.bas:2662`, `FG.bas:2352`, `FG.bas:3161`).
- Si `AltaPlanificacion` deja `FechaPrevistaCierre` como el viernes calculado (`FG.bas:2379`).
- La lista de semanas literal del `.form.txt` (`.form.txt:1427`) se sustituye siempre en tiempo de ejecución.
