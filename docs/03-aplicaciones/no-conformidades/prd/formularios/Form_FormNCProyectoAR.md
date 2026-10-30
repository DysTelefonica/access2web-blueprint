# Form_FormNCProyectoAR

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoAR.cls` (674 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana de alta, detalle y edición de una **tarea** (acción realizada, AR) de una acción correctiva (AC) de una NC de proyecto, con acceso a sus anexos de evidencia. Es un formulario **no enlazado** (sin `RecordSource`) que trabaja sobre los objetos globales `m_ObjNCProyectoActiva`, `m_ObjACProyectoActiva` y `m_ObjARProyectoActiva` (`Variables Globales.bas:117-119`). Las citas son `Form_FormNCProyectoAR.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `ARProyectoOperaciones.cls`, `classes/ARProyectoOperaciones.cls`. Hermano de [Form_FormNCProyectoAC](Form_FormNCProyectoAC.md).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Listado de tareas de la AC (Form_FormNCProyectoAcciones) | Abre el formulario con `OpenArgs` = estado de la NC; sin AR activa es un alta | `:556`, `:60` |

Abre `FormARProyectoDocumentos` (ficha de anexos) con «Documentos» (`:366`), con el icono de anexo (`:657`) y al intentar cerrar sin evidencia (`:459`). Se cierra con «Salir» (`:239-242`), tras registrar (`:517`) y ante error de apertura (`:575`). Levanta `Alta`/`Editado(estado)`, `AltaAnexo` y `EliminadoAnexo` (`:6-9,500-502,591,618`); quién los escucha es **pendiente** (referencia dinámica).

## Controles y eventos

`.form.txt` declara 36 bloques `Begin`: el formulario, 10 de valores por defecto (`.form.txt:406-468`), 3 secciones (`:475,512,1129`) y los 22 controles de la tabla (9 etiquetas, 7 cuadros de texto, 1 combo, 1 imagen y 4 botones, agrupados en las filas siguientes).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | «Alta de tarea de AC de proyecto» / «Detalle de tarea de AC de proyecto» (también en la barra de título) | `:61-62,93-94` |
| `AC` | texto bloqueado | — | Texto de la AC padre | `:57`, `.form.txt:755` |
| `AccionRealizada` (`OBLIGATORIO`) | texto | `AfterUpdate` | Descripción de la tarea; repinta bordes | `:233-237`, `.form.txt:535-537` |
| `FechaInicio`, `FechaFinPrevista` (`OBLIGATORIO`), `FechaFinReal` (`DATOS`) | texto | `AfterUpdate` cada una | Fechas de la tarea; repintan bordes | `:530-546`, `.form.txt:591-713` |
| `Notas` | texto | — | Notas libres | `:102-106` |
| `Responsable` | combo de 2 columnas (`DATO`) | — | Jefes de proyecto; muestra nombre y se graba la columna 2 | `:26,126,329,438`; `.form.txt:1070,1113` |
| `ESTADOAR` y `lblEstadoAR` | texto bloqueado y etiqueta | — | Estado de la tarea; ocultos en el alta | `:48-49,80-81,116-118` |
| `ImagenAnexo` | imagen | `OnClick` | Visible si la tarea tiene documentos; con un anexo lo abre en local y después abre la ficha de anexos | `:120-124,631-674` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Recalcula el estado de la tarea y, si cambia, el de la AC | `:244-310` |
| `ComandoDocumentos` | botón | `OnClick` | Abre la ficha de anexos; deshabilitado en el alta | `:67,355-381` |
| `cmdSalir`, `ComandoRegistrar` | botones | `OnClick` | Cerrar sin guardar / grabar | `:239,383` |
| Etiquetas `Etiqueta78/80/82/55/52/61/38` | etiquetas | — | Rótulos («F. inicio», «F. fin prevista», «F. fin real», «Tarea», «A.correc.», «Notas», «Resp.») | `.form.txt:563-1053` |

## Flujos

### Apertura (`:548-576`)
`Ajustar Me`, lee `OpenArgs` y llama a `EstablecerDatos`. Exige AC activa («El formulario se ha abierto con parámetros insuficientes», `:44-47`). Alta: precarga el responsable de la AC, deshabilita «Actualizar» y «Documentos» y pinta bordes (`:60-69`). Edición: relee la AR con `constructor.getARProyecto` (`:71`), vuelca valores, muestra el icono de anexo si hay documentos (`:95-130`) y deja solo lectura si la NC está borrada, tiene NC vinculada o `m_PermitidoEditar` = No (`:132-151`).

### Registrar (`:383-528`)
1. Rechaza a los técnicos (`:399-402`) y si no hay cambios: «No hay cambios que guardar» (`:403-410`).
2. Copia campos al objeto AR; en edición, si se informa `FechaFinReal` por primera vez y no hay documentos, ofrece adjuntar evidencia y **sale sin grabar** (`:445-466`).
3. Si cambia la fecha fin prevista pide por `InputBox` el motivo de replanificación; vacío ⇒ «No se ha grabado el cambio» (`:467-475`).
4. `ARProyectoOperaciones.Registrar` valida (`ARProyectoOperaciones.cls:133-142`), saneando `'` y `;` (`:146-147`); alta con `AddNew` (`:178-220`) y log «Alta de Acción Realizada» (`:222`); edición con `Edit`, recalculando estado (`:232-263`), log «Edición de Acción Realizada» (`:267`) y, si cambió la fecha prevista, alta de una replanificación (`:271-289`); invalida la caché (`:301`).
5. Aviso «Registrado correctamente», eventos y refresco de la gestión si cambió estado o texto; cierra (`:495-517`).

### Actualizar (`:244-310`)
Recalcula con `ActualizarDatosARProyecto` (`FG.bas:2609`); si cambió el estado, repinta y recalcula la AC con `ActualizarDatosACProyecto` (`:269`), refresca la gestión, recarga la NC y la línea de acciones (`:263-295`).

### Anexos
Los eventos `Alta`/`Eliminado` de la ficha de anexos anulan la colección, releen la AR y reabren el formulario con `Form_Open False` (`:577-629`).

### Errores
Error 1000 ⇒ aviso con el motivo; otro ⇒ correo al administrador y `MsgBox` crítico (`:519-527`); en `Form_Open` cierra el formulario (`:566-575`).

## Reglas, validaciones y cálculos

- Obligatorios: AC y descripción (`ARProyectoOperaciones.cls:34-41`); fin real exige fin prevista, fin prevista exige inicio, inicio exige fin prevista (`:42-59`); coherencia de fechas en `BuenaFecha` (`:62,379`, no desglosada).
- Cerrar una tarea (fin real) exige al menos un documento (`:68-87`); acción repetida en la misma AC se rechaza (`:88-95`).
- Bordes: se pintan en rojo según faltas, fin prevista anterior al inicio o a hoy sin fin real, o inicio posterior al fin real (`:183-225`).
- Responsable vacío ⇒ hereda el de la AC (`ARProyectoOperaciones.cls:149-151`). Estado derivado (`:262`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta de tarea | `TbNCAccionesRealizadas` | DAO `AddNew`; SQL = nombre de tabla | `ARProyectoOperaciones.cls:153,170,179-220` |
| Edición de tarea | `TbNCAccionesRealizadas` | SQL dinámico concatenado por `IDAccionRealizada` + `Edit` | `ARProyectoOperaciones.cls:167-168,232-265` |
| Replanificación | clase `ReplanificacionesProyectoOperaciones` | no resuelto en esta ficha | `ARProyectoOperaciones.cls:281-284` |

Sin transacción explícita entre tarea, log y replanificación.

## Coexistencia

SQL concatenado con identificador (`ARProyectoOperaciones.cls:167`) ⇒ parámetros. IDs calculados por la clase (`:154`) y no autonumérico. `Now()` y campos fecha/`Null` (`:183,186-202`). Definir transacción única para tarea + log + replanificación + recálculo.

## Comportamiento actual · Posibles defectos · Propuestas

**Comportamiento actual.** En el alta, `HaHabidoCambios` siempre devuelve verdadero (`:316-318`). Al grabar con fin real, la fecha de alta de la tarea toma el fin real (`ARProyectoOperaciones.cls:197-199,245`).

**Posibles defectos (a verificar).**
- Alta con fin real: «Documentos» está deshabilitado en el alta (`:67`) pero la validación exige documentos (`ARProyectoOperaciones.cls:80-85`); no se puede cerrar una tarea al darla de alta.
- `HaHabidoCambios` compara fechas y texto como variantes (`:333-341`); un formato distinto podría producir cambios falsos.
- `Alta` y `Editado` reciben `m_EstadoAlInicio` del módulo (`:500-502`), no el local de `ComandoActualizar` (`:245`).
- Responsable escrito a mano fuera de la lista da `Column(1)` nulo (`:438`).

**Propuestas web pendientes de decisión.** Pedir motivo de replanificación en el propio formulario; permitir subir evidencia en el alta.

## Criterios de aceptación

1. Sin AC activa, aviso y cierre (`:44-47,575`).
2. Un técnico no puede registrar (`:399-402`); sin cambios no se graba (`:407-410`).
3. Cambiar la fecha fin prevista exige motivo; sin él no se graba (`:467-475`).
4. Informar fin real sin documentos propone adjuntar y no graba (`:445-466`).
5. Descripción repetida en la misma AC se rechaza (`ARProyectoOperaciones.cls:88-95`).
6. NC borrada, vinculada o no editable ⇒ solo lectura (`:132-151`).

## Pendientes

- Quién escucha `Alta`, `Editado`, `AltaAnexo` y `EliminadoAnexo` (`:6-9`).
- Cuerpos de `BuenaFecha` (`ARProyectoOperaciones.cls:379`), `ActualizarDatosCalculados` (`ARProyectoOperaciones.cls:459`), `constructor.getARProyecto` (`:71`) y `RegistrarLogProyecto`.
- Tras `Registrar`, tramo `ARProyectoOperaciones.cls:303-343` (recálculo de AC y NC) no leído en esta ficha.
