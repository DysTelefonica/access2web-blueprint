# Form_FormUsuario02DatosObservacion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02DatosObservacion.cls` (194 líneas) y `.form.txt` (1375 líneas) · **PRD:** [README](../README.md)

Alta y edición de una observación del usuario **actual**: un formulario modal-ish (sin `PopUp`/`Modal`) con tres campos (`Fecha`, `Observacion`, `TIPO`) y dos botones (`ComandoGrabar`, `cmdSalir`). Decide entre alta nueva y edición consultando si el global `m_ObjObservacionActiva` está vacío; en edición detecta cambios con `HaHabidoCambios` y, si los hay, pregunta al cerrar. Publica dos eventos (`AltaRegistro`, `EdicionRegistro`) que el contenedor (típicamente `Form_FormUsuario02Edicion`) escucha para refrescar colecciones. Las citas son `Form_FormUsuario02DatosObservacion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario02Edicion`, `Form_FormUsuario02HPS` u otro subform de detalle | `DoCmd.OpenForm "FormUsuario02DatosObservacion", , , , , , m_Argumento`, pasando el tipo en `OpenArgs` | no hay `DoCmd.OpenForm` en `src` con ese nombre; `OpenArgs` se lee en `:120` |
| `m_ObjUsuarioActivo` (global) | Necesario para `ID` y `NombreCompleto` en el título | `:127-130, 142, 148` |
| `m_ObjObservacionActiva` (global) | Si no es `Nothing`, se entra en edición; si es `Nothing`, en alta | `:131-167` |
| Eventos `AltaRegistro`, `EdicionRegistro` | El propio `ComandoGrabar_Click` (`:79, 102`) y `CmdSalir_Click` (`:38`) lanzan `RaiseEvent` con `m_ObjObservacionActiva.TIPO` | `:9-10` |

Al cerrarse no se ejecuta código de salida: `DoCmd.Close` se limita al método de Access. Si el llamador tiene `WithEvents`, la notificación llega por evento, no por retorno.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Título centrado, reescrito en `Form_Load` con «ALTA DE OBSERVACIÓN …», «EDICIÓN …» o «DETALLE …» | `:169-170`, `.form.txt:864, 873` |
| `Fecha` | `TextBox` (`Format = "Short Date"`, `Tag = "DATOS"`) | — | Fecha de la observación; precargado con `Date` en alta | `:145`, `.form.txt:927, 938` |
| `Observacion` | `TextBox` multilínea (`Tag = "DATOS"`, `TabIndex = 1`) | — | Texto de la observación; recibe foco al abrir | `:147, 167`, `.form.txt:985, 997` |
| `TIPO` | `TextBox` (`Locked = NotDefault`, `TabStop = NotDefault`, `Tag = "DATOS"`) | — | Categoría («General» / «Curso»); bloqueado tras la asignación inicial | `:146`, `.form.txt:1043, 1057` |
| `Etiqueta14` | etiqueta «Fecha» | — | Rótulo | `.form.txt:898, 912` |
| `Etiqueta3` | etiqueta «OBSERVACIÓN» | — | Rótulo | `.form.txt:956, 970` |
| `Etiqueta16` | etiqueta «TIPO» | — | Rótulo | `.form.txt:1014, 1028` |
| `ComandoGrabar` («Grabar») | botón (`OnClick`) | `Click` → `ComandoGrabar_Click` | Crea o actualiza la observación y cierra el formulario | `:53-109`, `.form.txt:1087, 1094` |
| `cmdSalir` («Salir») | botón (`OnClick`) | `Click` → `CmdSalir_Click` | Cierre con pregunta de guardar | `:14-51`, `.form.txt:1130, 1138` |
| (formulario) | — | `OnLoad` → `Form_Load` | Decide alta o edición, fija el título y aplica estilo | `:111-194`, `.form.txt:372` |
| `EncabezadoDelFormulario`, `PieDelFormulario` | secciones | — | Cabecera y pie; `BackColor = m_ObjEntorno.CabeceraBackColor` | `:179-180`, `.form.txt:1076, 1079` |

Sin ribbon de Access; `NavigationButtons = NotDefault`, `CloseButton = NotDefault`, `ControlBox = NotDefault` (`.form.txt:9-13`). `Me.ScrollBars = 0` y `Me.DividingLines = NotDefault` (`.form.txt:15-16`).

## Flujos

### Apertura (`Form_Load`, `:111-194`)

1. `AjustarTamaño Me` (`:118`; helper en `FG.bas:901`).
2. Lee `OpenArgs` con `Nz`; cadena vacía si falta (`:120`).
3. Si `Application.TempVars("SoloLectura") = "Sí"`, `Me.AllowEdits = False`; en caso contrario, editable (`:122-126`).
4. Si `m_ObjUsuarioActivo` es `Nothing`, eleva «No se conoce el usuario activo» (`:127-130`).
5. **Si `m_ObjObservacionActiva` es `Nothing` (alta nueva):**
   1. Si `AllowEdits = False`, eleva «El Sistema está en modo Sólo lectura» (`:131-135`); un formulario abierto en modo lectura nunca da de alta.
   2. Valida `m_Argumento`: solo «General» o «Curso»; cualquier otro valor eleva «Formulario abierto con parámetros insuficientes» (`:136-139`).
   3. Crea un `Observacion` nuevo, le asigna `ID = m_ObjUsuarioActivo.ID` y `ParaAlta = EnumSiNo.Sí`, y rellena los controles: `Me.TIPO = m_Argumento`, `Me.Fecha = Date`, `Me.TIPO.Enabled = False` (`:141-146`).
   4. Pone el foco en `Observacion` (`:147`) y fija el título «ALTA DE OBSERVACIÓN <TIPO>: <nombre>» (`:148`); `ComandoGrabar.Enabled = True` (`:149`).
6. **Si `m_ObjObservacionActiva` no es `Nothing` (edición o solo lectura):**
   1. Mira `AllowEdits`: si verdadero, título «EDICIÓN DE OBSERVACIÓN …» y `ComandoGrabar.Enabled = True`; si falso, «DETALLE DE OBSERVACIÓN …» y `ComandoGrabar.Enabled = False` (`:151-157`).
   2. Carga `m_ObjObservacionAlInicio` con `getObservacion(m_ObjObservacionActiva.IDObservacion, m_Error)` (`:158-161`; SQL en `modules/Constructor.bas:657`): si `m_Error` es no vacío, eleva `1000`.
   3. Pasa los valores del objeto al formulario con `DatosDelObjetoAlFormulario` (helper en `FG.bas:4180`, `:163-166`).
   4. Pone el foco en `Observacion` (`:167`).
7. Aplica `m_Titulo` a `lblTitulo.Caption` y `Caption` (`:169-170`).
8. Aplica estilo a los botones (`CopiarEstiloBoton`, `FG.bas:2435`, firma con `ByRef p_Error`, dos bloques, `:171-178`), repinta cabecera y pie con `m_ObjEntorno.CabeceraBackColor` (`:179-180`) y termina con `EstablecerControlesDatos` (`FG.bas:1005`, `:181-184`).
9. En error: «Al abrir el formulario se ha producido el error: …», `CorreoAlAdministrador` si no es `1000`, `MsgBox` y **cierra el formulario sin guardar** (`:188-193`).

### Guardar (`ComandoGrabar_Click`, `:53-109`)

1. `RellenarObjetoConDatos m_ObjObservacionActiva, Me, m_Error` (`:60`; helper en `FG.bas:4124`). Error ⇒ `1000`.
2. **Si `m_ObjObservacionAlInicio` es `Nothing` (alta):**
   1. `m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí` y `m_ObjObservacionActiva.ID = m_ObjUsuarioActivo.ID` (`:67-68`).
   2. `m_ObjObservacionActiva.Registrar m_Error` (`:69`; clase `Observacion`, método `Registrar` en `classes/Observacion.cls:152`). Error ⇒ `1000`.
   3. Invalida la caché del usuario activo: si `Me.TIPO = "General"` vacía `m_ObjUsuarioActivo.ColObservaciones`; si «Curso», `ColObservacionesCurso` (`:73-77`).
   4. Lanza `RaiseEvent AltaRegistro(m_ObjObservacionActiva.TIPO)` (`:79`).
3. **Si `m_ObjObservacionAlInicio` no es `Nothing` (edición):**
   1. `HaHabidoCambios(m_ObjObservacionAlInicio, Me, m_Error)` (`:81`; helper en `FG.bas:4147`). Error ⇒ `1000`.
   2. Si **no** hay cambios, eleva «No hay cambios que guardar» y sale (`:85-91`): no graba, no cierra. Es la única condición de salida sin error que aborta la operación.
   3. `m_ObjObservacionActiva.ParaAlta = EnumSiNo.No` y vuelve a `Registrar` (`:92-96`).
   4. Lanza `RaiseEvent EdicionRegistro(m_ObjObservacionActiva.TIPO)` (`:102`).
4. `DoCmd.Close acForm, Me.Name, acSaveNo` (`:104`).
5. En error: «Al ComandoGrabar_Click se ha producido el error: …», `CorreoAlAdministrador`, `MsgBox` **sin cerrar** (`:106-109`).

### Salir (`CmdSalir_Click`, `:14-51`)

1. Si `m_ObjObservacionActiva.ParaAlta = EnumSiNo.No` (es edición) (`:22`):
   1. `HaHabidoCambios` (`:23`).
   2. Si `m_HaHabidoCambios`, pregunta «¿Desea guardar los cambios antes de cerrar el formulario?» (`:31`). Si la respuesta es Sí, `m_ObjObservacionActiva.Registrar m_Error` (`:33`) y `RaiseEvent EdicionRegistro` con el tipo (`:38`). Si es No, simplemente sigue.
2. `DoCmd.Close acForm, Me.Name, acSaveNo` (`:42`).
3. En error: «Al intentar salir se ha producido el error: …», `CorreoAlAdministrador`, `MsgBox`, **sin cerrar** (`:45-49`). El formulario se queda abierto si algo falla en la pregunta.

### Errores

Tres manejadores con el mismo patrón, todos dependientes del global `pregunta` (`modules/VariablesEntorno.bas:220`):

- `Form_Load:188-193`: compone «Al abrir el formulario se ha producido el error: …», eleva `1000` si procede, `CorreoAlAdministrador`, `MsgBox`, cierra.
- `ComandoGrabar_Click:103-108`: «Al ComandoGrabar_Click …», no cierra.
- `CmdSalir_Click:44-49`: «Al intentar salir …», no cierra.

## Reglas, validaciones y cálculos

- **Tipo abierto en `OpenArgs`.** Solo acepta «General» o «Curso» (`:136-139`); cualquier otro valor aborta el alta antes de tocar la base de datos. La edición no valida `OpenArgs`, porque la categorización ya viene del objeto existente.
- **Invalidez en modo lectura.** Si `SoloLectura="Sí"`, el alta se aborta (`:131-135`) y la edición se queda con `ComandoGrabar` deshabilitado (`:156`); el cierre sigue permitido. Es el camino sin escritura, no la cancelación.
- **`m_ObjObservacionAlInicio` como snapshot.** La copia se hace en alta solo si llega un `m_ObjObservacionActiva` no nulo (`:158-161`); en alta nueva, `m_ObjObservacionAlInicio` sigue `Nothing` y el `ComandoGrabar_Click` entra por la rama alta (`:66`). Esto es lo que decide sin un parámetro «modo» adicional.
- **Detección de cambios por `HaHabidoCambios`.** Compara el snapshot contra el estado actual del formulario (helper en `FG.bas:4147`). Si la edición «no tiene cambios», no se graba: lanza `1000` con «No hay cambios que guardar» (`:85-91`), no un mensaje de usuario.
- **Invalidación de caché post-alta.** `Set m_ObjUsuarioActivo.ColObservaciones = Nothing` o `ColObservacionesCurso` (`:74, 76`) fuerza que el siguiente consumidor (lista de observaciones del contenedor) relea; sin esto, la caché seguiría mostrando el conjunto previo.
- **`TIPO` se queda `Enabled = False` en alta** (`:146`). En edición el control no se bloquea explícitamente, así que depende del flag global `AllowEdits`. Si la edición se abre en modo lectura, la propiedad `Locked = NotDefault` del control (`.form.txt:1044`) es la última garantía.
- **`Observacion` recibe foco al abrir.** Tanto en alta como en edición (`:147, 167`); el usuario puede teclear inmediatamente sin pasar por `Tab`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga de la observación a editar | `getObservacion` (Constructor) | SQL en `modules/Constructor.bas:657` | `:158-161` |
| Alta o edición en base | `Observacion.Registrar` | `INSERT`/`UPDATE` en `classes/Observacion.cls:152` | `:69, 93` |
| Invalidación de caché | `m_ObjUsuarioActivo.ColObservaciones` / `ColObservacionesCurso` | `Set … = Nothing` | `:74, 76` |
| Notificación al contenedor | `RaiseEvent AltaRegistro` / `EdicionRegistro` | evento VBA | `:9-10, 38, 79, 102` |

Sin transacciones explícitas. El alta/edición se delega en `Observacion.Registrar` (clase); errores se propagan por `m_Error`. Caché del usuario activo se invalida solo cuando se añade (no en edición).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Cosas que ya se ven: `TIPO.Enabled = False` se aplica tras un `Me.TIPO = m_Argumento` (`:146`, dentro del bloque `:141-146`), así que en web debe traducirse a un campo no editable desde el inicio del alta; la pregunta «¿Desea guardar los cambios antes de cerrar?» es exclusiva de este formulario del PRD: `Form_FormUsuario02Edicion` y `Form_FormUsuario02HPS` aplican patrón similar (a verificar) pero no idéntico; el evento `EdicionRegistro` se lanza tanto desde `ComandoGrabar_Click` (`:102`) como desde `CmdSalir_Click` cuando el usuario eligió «Sí» en la pregunta (`:38`), así que el contenedor puede recibir la notificación dos veces si no deduplica.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En `CmdSalir_Click` no se mide `SoloLectura` (`:22-41`): si el sistema está en modo lectura con cambios en memoria (no debería, pero `AllowEdits = False` no impide `Me.Observacion.Text = …` en tiempo de diseño ni cancela eventos sueltos), la pregunta de guardar se dispara igualmente.
- `m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí` se fuerza en alta (`:67`) y luego `Registrar` decide qué hacer (clase `Observacion.cls:177` lee el flag). En edición (`:92`) se fuerza `No`. Pero la **primera apertura** con `m_ObjObservacionActiva` ya rellenado puede venir con `ParaAlta` en cualquier estado heredado del constructor; `Form_Load` no lo normaliza.
- `HaHabidoCambios` se evalúa dos veces seguidas en `ComandoGrabar_Click`: una vez dentro del `If` (`:81`) y otra en una rama más interna del estilo «If m_Error <> "" Then» (a verificar el patrón duplicado en `:82-88`). El segundo `If m_Error <> "" Then` con `Err.Raise 1000` está duplicado sin consumir el valor.
- `m_Titulo` se concatena con `UCase(m_Argumento)` (`:148, 152, 155`). Si `m_Argumento` viene vacío (lo que no debería pasar porque `Form_Load` lo aborta), el título queda con «OBSERVACIÓN : …».
- `TIPO.Enabled = False` se aplica dentro de la rama alta (`:146`), pero en la rama edición no se restablece (`:151-167`). Si la edición abre con `AllowEdits = True` y luego se cambia a `AllowEdits = False`, el campo sigue editable.
- `CmdSalir_Click` no cancela el evento si `Registrar` falla (`:33-37`): eleva `1000`, muestra `MsgBox` y **no cierra**, pero la pregunta ya se hizo y se eligió «Sí»; el usuario verá el mensaje y verá el formulario seguir abierto con datos posiblemente sin guardar.
- `RaiseEvent EdicionRegistro` puede dispararse dos veces (`:38` desde el `cmdSalir`, `:102` desde `comandoGrabar`) si el usuario llega a pulsar ambos en secuencia; los contenedores con `WithEvents` deben deduplicar.
- `ComandoGrabar.Enabled = False` solo aplica al modo detalle (`:156`); en cualquier otro modo editable el botón siempre está activo, incluso si al pulsar no hay cambios (el `1000` lo descubre en `:85-91` con el mensaje «No hay cambios que guardar»).
- El título «DETALLE DE OBSERVACIÓN …» (`:155`) concatenado sin espacio entre «OBSERVACIÓN» y `<TIPO>` es lo que muestra en la versión del binario: queda como `… OBSERVACIÓNGeneral: …`. A corregir en web.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir el contrato del alta (tipo en `OpenArgs`, en ruta, en store) en lugar de `OpenArgs`; normalizar `m_ObjObservacionActiva.ParaAlta` en `Form_Load` antes de la lógica; deduplicar la notificación de `EdicionRegistro` con un guard; reemplazar los `MsgBox` por notificaciones no bloqueantes; validar `TIPO` desde una tabla de tipos en lugar de un literal `If/ElseIf`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con `OpenArgs` vacío o distinto de «General» / «Curso», el alta se cancela con el aviso «Formulario abierto con parámetros insuficientes».
2. Con `SoloLectura="Sí"`, abrir en alta aborta; abrir en edición entra en modo detalle con «ComandoGrabar» deshabilitado.
3. Tras guardar una nueva observación, la colección correspondiente del usuario activo queda invalidada (`Set … = Nothing`).
4. Si en edición no hay cambios, «Grabar» no graba y se ve el aviso «No hay cambios que guardar».
5. Al pulsar «Salir» con cambios pendientes, se pregunta y, según la respuesta, se graba o no antes de cerrar.
6. El contenedor que abrió el formulario recibe `AltaRegistro` o `EdicionRegistro` con el `TIPO` correspondiente.

## Pendientes

- Sentencias SQL exactas de `Observacion.Registrar` (`classes/Observacion.cls:152`) y de `getObservacion` (`modules/Constructor.bas:657`): no se han reproducido aquí para evitar duplicar la fuente.
- Firma exacta de `CopiarEstiloBoton` (`FG.bas:2435`): la versión con `ByRef p_Error` es la invocada en `:171-178`; hay otras firmas con un solo argumento en otras llamadas del PRD (a verificar).
- Quién escucha los eventos `AltaRegistro`/`EdicionRegistro`: no hay `WithEvents` en este formulario; buscar `WithEvents m_Obj…` y `clsTestDoubleHpsEditListener` en `classes/clsTestDoubleHpsEditListener.cls:25`.
- Si `m_ObjObservacionActiva` puede llegar con `ParaAlta = Sí` desde un flujo que no sea el alta de este formulario.
- `m_ObjObservacionAlInicio` nunca se libera; queda con la última referencia hasta la próxima apertura.
