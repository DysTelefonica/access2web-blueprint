# Form_FormNCProyectoAC

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoAC.cls` (339 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana de alta y edición de una acción correctiva (AC) de una NC de proyecto. Es un formulario **no enlazado** (sin `RecordSource`; la única referencia de origen es el `RowSource` de lista de valores del combo, que el código vacía y rellena en ejecución, `Funciones Generales.bas:1816`). Trabaja sobre los objetos globales `m_ObjNCProyectoActiva` y `m_ObjACProyectoActiva` (`Variables Globales.bas:117,119`). Las citas son `Form_FormNCProyectoAC.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `VG.bas` abrevia `modules/Variables Globales.bas` y `ACProyectoOperaciones.cls` abrevia `classes/ACProyectoOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Pestaña de acciones de la NC de proyecto (Form_FormNCProyectoAcciones) | Abre el formulario con `OpenArgs` = estado de la NC; si no hay AC activa es un alta | `:316`, `:55` |

Se cierra a sí mismo con «Salir» (`:171-179`), tras registrar (`:291`) y ante cualquier error de apertura (`:336`). Al registrar levanta los eventos públicos `Alta(estado)` o `Editado` (`:8-9,276-281`) y, si está abierto `FormNCProyectoGestion`, refresca su línea (`:284-286`). La ficha no resuelve quién escucha esos eventos: **pendiente** (referencia dinámica, `:8-9`).

## Controles y eventos

`.form.txt` declara 34 bloques `Begin`: el formulario (`.form.txt:4`), 10 de valores por defecto (`:406-468`), 3 secciones (`:475,512,1073`) y los 20 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | «Alta de Acción Correctiva» o «Edición de acción correctiva» y mismo texto en la barra de título | `:56-57,101-103` |
| `NoConformidad` | cuadro de texto bloqueado | — | Descripción de la NC o «Sin definir» | `:44-49`, `.form.txt:611-612` |
| `AccionCorrectiva` | cuadro de texto, `Tag` `OBLIGATORIO` | `AfterUpdate` | Recalcula el borde de obligatorios si la NC es editable | `:162-169`, `.form.txt:534-537` |
| `Notas` | cuadro de texto | — | Notas libres; `Null` si vacío | `:107-111` |
| `Responsable` | combo de 2 columnas, `Tag` `DATO` | — | Lista de jefes de proyecto (nombre; usuario de red); se muestra el nombre y se graba la columna 2 | `:21,131,266`; `FG.bas:1822-1829`; `.form.txt:957,1000` |
| `FechaInicialMinima`, `FechaFinPrevistaUltima`, `FechaFinalUltima`, `Estado` | 3 fechas + 1 texto, bloqueados | `FechaInicialMinima` declara `AfterUpdate` sin handler | Valores calculados, solo visibles en edición | `:62-69,115-129`, `.form.txt:553-554,570` |
| `lblfecha1`, `lblfecha2`, `lblfecha3`, `lblEstado1`, `Etiqueta52`, `Etiqueta31`, `Etiqueta61`, `Etiqueta38` | etiquetas | — | Rótulos («F.inicio», «F. fIn real», «F.fin prevista», «Estado», «NC», «Acción», «Notas», «Resp.») | `.form.txt:590-1050` |
| `ComandoActualizar` (rótulo «Contador») | botón | `OnClick` | Recalcula estado y fechas de la AC y refresca el formulario | `:181-235`, `.form.txt:802-815` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:171-179` |
| `ComandoRegistrar` | botón | `OnClick` | Valida y graba la AC | `:237-303` |

## Flujos

### Apertura (`Form_Open`, `:307-337`)

1. Reloj de arena, `Ajustar Me`, lee `OpenArgs` y llama a `EstablecerDatos` (`:311-318`).
2. `EstablecerDatos` falla con «Se ha abierto el formulario con parámetros insuficientes» si no hay NC activa (`:40-43`), rellena el combo (`:51`) y, **sin AC activa**, oculta calculados y `ComandoActualizar`, preselecciona como responsable al de la NC si lo tiene y pinta obligatorios (`:55-72`).
3. **Con AC activa** relee la AC con `constructor.getACProyecto` (`:84`), vuelca textos, fechas, estado y responsable (`:99-135`).
4. Edición permitida solo si la NC no está borrada, no tiene NC vinculada y `Form_FormNCProyecto.m_PermitidoEditar` no es «No»; en otro caso `AllowEdits = False` y `ComandoRegistrar` deshabilitado (`:136-153`). Con NC vinculada también se oculta `ComandoActualizar` (`:92-98`).

### Registrar (`:237-303`)

1. Si `EsTecnico = Sí` ⇒ «No tiene autorización para esa acción» (`:255-258`; global `VG.bas:98`).
2. Alta: crea `ACProyecto` con el `IDNoConformidad` de la NC (`:259-262`); copia texto, notas y la columna 2 del responsable (`:263-267`).
3. `ACProyectoOperaciones.Registrar`: sin NC vinculada valida con `MotivoNoOK` (`ACProyectoOperaciones.cls:123-132`); sustituye `'` por vacío y `;` por `:` en el texto (`:133-134`); alta ⇒ `AddNew` en `TbNCAccionCorrectivas` con ID calculado, `NAccion`, fecha `Now()` y log «Alta de Acción Correctiva» (`:137-206`); edición ⇒ `Edit` de IDNC, texto, notas, responsable y `NAccion` si falta, con log «Edición de Acción Correctiva» (`:211-229`); invalida la caché de la NC y, sin NC vinculada, recalcula (`:238-246`).
4. Levanta `Alta`/`Editado`, refresca la gestión y cierra (`:276-291`).

### Actualizar (`:181-235`)

Vacía fechas y estado, llama a `ActualizarDatosACProyecto` (`FG.bas:2471-2490`), que recalcula estado y las tres fechas (`ACProyectoOperaciones.cls:359-375`), y repinta. Si cambió el estado, refresca la línea de gestión, recarga la NC y, si el detalle activo es `FormNCProyectoAcciones`, su línea (`:201-216`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con el motivo; otro error ⇒ correo al administrador y `MsgBox` crítico (`:226-234,293-301`). En `Form_Open` cualquier error cierra el formulario (`:327-336`).

## Reglas, validaciones y cálculos

- Obligatorio: texto de la AC (`ACProyectoOperaciones.cls:71-74`); la acción no puede repetirse en la misma NC si es alta o cambia el texto (`:75-82,89-97`).
- Responsable informado debe existir en la lista de usuarios (`ACProyectoOperaciones.cls:83-88`).
- Los técnicos no pueden registrar (`:255-258`).
- Estado y fechas son derivados, no editables (`.form.txt:554,612,688,745,1015`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta de AC | `TbNCAccionCorrectivas` | DAO `OpenRecordset` + `AddNew`/`Update`; SQL = nombre de tabla | `ACProyectoOperaciones.cls:137,157,168-201` |
| Edición de AC | `TbNCAccionCorrectivas` | DAO con SQL dinámico concatenado por `IDAccionCorrectiva` | `ACProyectoOperaciones.cls:154-155,211-228` |
| Relectura / cálculos | vía `constructor.getACProyecto`, `ActualizarDatosCalculados` | no resuelto en esta ficha | `:84`; `ACProyectoOperaciones.cls:359` |

Sin transacción explícita en `Registrar`; el alta y el recálculo son operaciones DAO separadas (`ACProyectoOperaciones.cls:201,241`).

## Coexistencia

- Dialecto Jet: el SQL de edición concatena un identificador numérico (`ACProyectoOperaciones.cls:154-155`); en PostgreSQL pasar a parámetros.
- Alta con `AddNew` sobre ID calculado por la clase (`:138`), no autonumérico: la generación del ID debe mantenerse atómica con el vínculo.
- `Now()` como fecha de alta (`:171`) y fechas con `IsDate`/`Null` (`:115-129`): verificar tipo y zona horaria.
- Sin transacción entre alta, log y recálculo: definir límite transaccional.

## Comportamiento actual · Posibles defectos · Propuestas

**Comportamiento actual.** El responsable se muestra por nombre y se graba por la columna 2 del combo (`:131,266`). Los registros se sanean eliminando `'` y `;` del texto, pero no de las notas (`ACProyectoOperaciones.cls:133-134`).

**Posibles defectos (a verificar).**
- Si el responsable se escribe a mano y no coincide con una fila, `Column(1)` es nulo y se graba sin responsable (`:266`).
- `FechaInicialMinima` declara `AfterUpdate` sin procedimiento (`.form.txt:570`); en una edición con la AC sin cambios de texto, la edición no incluye estado ni fechas (`ACProyectoOperaciones.cls:211-228`).
- Botón con rótulo «Contador» para la acción de recalcular (`.form.txt:814`).
- `ComandoActualizar_Click` declara `m_EstadoAlInicio` local que oculta la variable del módulo (`:183` frente a `:6`); el evento `Alta` recibe la del módulo (`:278`).

**Propuestas web pendientes de decisión.**
- Sustituir el sanitizado por parámetros; mostrar estado y fechas como solo lectura calculada.

## Criterios de aceptación

1. Sin NC activa, el formulario muestra el error de parámetros y se cierra (`:40-43,336`).
2. Alta: título «Alta de Acción Correctiva», calculados ocultos, responsable preseleccionado si la NC lo tiene (`:55-72`).
3. Registrar con texto vacío o repetido en la misma NC se rechaza con su motivo (`ACProyectoOperaciones.cls:71-97`).
4. Un técnico no puede registrar (`:255-258`).
5. Con NC borrada, vinculada o `m_PermitidoEditar` = No, el formulario es solo lectura (`:136-153`).
6. «Actualizar» recalcula estado y fechas y propaga el cambio de estado a la gestión (`:195-216`).

## Pendientes

- Quién escucha `Alta` y `Editado` (`:8-9,278,280`).
- Cuerpos de `constructor.getACProyecto` (`:84`), `AccionRepetida`, `IDAccionCorrectivaCalculada` y `EstadoGrabar` (`ACProyecto.cls:228`).
- Existencia en `main` de la tabla `TbNCAccionCorrectivas`: aparece en `Instalador.bas` y `Variables Globales.bas`, no hay definición de tablas en `src`.
