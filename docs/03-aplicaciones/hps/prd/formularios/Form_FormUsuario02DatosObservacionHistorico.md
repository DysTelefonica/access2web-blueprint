# Form_FormUsuario02DatosObservacionHistorico

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02DatosObservacionHistorico.cls` (60 líneas) y `.form.txt` (1197 líneas) · **PRD:** [README](../README.md)

Visor de solo lectura de una observación del usuario **histórico**. Se construye pasando el `IDObservacion` (cadena codificada en `OpenArgs` con el formato interno de `Constructor.getObservacionHistorica`) y se mapea a un objeto `ObservacionHistorica` cuyo `Fecha`, `TIPO` y `Observacion` se vuelcan en los tres `TextBox` del formulario. **No permite editar**: `Me.AllowEdits = False` se fuerza en `Form_Load:27` y el único botón es «Salir». Las citas son `Form_FormUsuario02DatosObservacionHistorico.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioHistorico02Edicion`, `Form_FormUsuarioHistorico02Anexos` u otro subform de histórico | `DoCmd.OpenForm "FormUsuario02DatosObservacionHistorico", , , , , , m_Argumento` | no hay `DoCmd.OpenForm` en `src` con ese nombre; `OpenArgs` se lee en `:28` |
| `m_ObjEntorno` (global) | Lectura para colorear la cabecera y el pie | `:44-45` |

No abre otros formularios. No publica eventos: la notificación del visor al contenedor es por cierre (`DoCmd.Close`), no por `RaiseEvent`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Título centrado, fijo «Observación Histórica» | `:51`, `.form.txt:864, 873` |
| `Fecha` | `TextBox` (`Format = "Short Date"`, `Tag = "DATOS"`) | — | Fecha de la observación histórica | `.form.txt:927, 938` |
| `Observacion` | `TextBox` multilínea (`Tag = "DATOS"`, `TabIndex = 1`) | — | Texto de la observación | `.form.txt:985, 997` |
| `TIPO` | `TextBox` (`Locked = NotDefault`, `TabStop = NotDefault`, `Tag = "DATOS"`) | — | Categoría («General» / «Curso» o equivalente histórico) | `.form.txt:1043, 1057` |
| `Etiqueta14` («Fecha») | etiqueta | — | Rótulo | `.form.txt:898, 912` |
| `Etiqueta3` («OBSERVACIÓN») | etiqueta | — | Rótulo | `.form.txt:956, 970` |
| `Etiqueta16` («TIPO») | etiqueta | — | Rótulo | `.form.txt:1014, 1028` |
| `cmdSalir` («Salir») | botón | `OnClick` → `CmdSalir_Click` | Cierra el visor | `:10-16`, `.form.txt:1087, 1138` |
| (formulario) | — | `OnLoad` → `Form_Load` | Carga los datos desde `Constructor.getObservacionHistorica` y aplica estilo | `:18-59`, `.form.txt:372` |
| `EncabezadoDelFormulario`, `PieDelFormulario` | secciones | — | Cabecera y pie; `BackColor = m_ObjEntorno.CabeceraBackColor` | `:44-45`, `.form.txt:1076, 1079` |

A diferencia de `Form_FormUsuario02DatosObservacion` (mismo PRD), este visor no declara `ComandoGrabar` ni eventos `AltaRegistro`/`EdicionRegistro`; toda la sección de detalle se reduce al `cmdSalir`.

## Flujos

### Apertura (`Form_Load`, `:18-59`)

1. `AjustarTamaño Me` (`:26`; helper en `FG.bas:901`).
2. `Me.AllowEdits = False` siempre — el visor es de solo lectura por diseño (`:27`). No consulta `SoloLectura` ni `TempVars`.
3. `m_Argumento = Nz(Me.OpenArgs(), "")` (`:28`).
4. `Constructor.getObservacionHistorica(m_Argumento, m_Error)` (`:29`; factory en `modules/Constructor.bas:700`). Error ⇒ `Err.Raise 1000`.
5. Si `m_ObsHistorica Is Nothing`, eleva «Se ha abierto con parámetros insuficientes» (`:32-35`). Este es el guard principal del visor.
6. Vuelca los campos al formulario, **solo si son utilizables**: `IsDate(m_ObsHistorica.Fecha)` (`:36`), `m_ObsHistorica.TIPO <> ""` (`:37`) y `m_ObsHistorica.Observacion <> ""` (`:38`). Si una propiedad viene vacía o no es fecha, el `TextBox` correspondiente queda en blanco.
7. Aplica `CopiarEstiloBoton` (helper en `FG.bas:2435`, firma `ByRef p_Error`, `:40-43`).
8. Pinta cabecera y pie con `m_ObjEntorno.CabeceraBackColor` (`:44-45`).
9. Aplica estilo vía `EstablecerControlesDatos` (`FG.bas:1005`, `:46-49`).
10. `Me.Caption = "Observación Histórica"` y `Me.lblTitulo.Caption = Me.Caption` (`:50-51`).
11. En error: «Al abrir el formulario se ha producido el error: …», `CorreoAlAdministrador` si no es `1000`, `MsgBox` y **cierra el visor** sin guardar (`:53-58`).

### Cierre (`CmdSalir_Click`, `:10-16`)

`DoCmd.Close acForm, Me.Name, acSaveNo`, sin más lógica. No pregunta nada: al ser solo lectura no hay datos modificados que salvar.

### Errores

Un único manejador en `:53-58`. Patrón: `Err.Number <> 1000` ⇒ compone mensaje y `CorreoAlAdministrador`; muestra `MsgBox`; siempre cierra.

## Reglas, validaciones y cálculos

- **`OpenArgs` obligatorio.** Sin un argumento que `Constructor.getObservacionHistorica` sepa resolver, el visor aborta en `:32-35` con «Se ha abierto con parámetros insuficientes». El visor no intenta decodificar el argumento por su cuenta.
- **Tolerancia a datos incompletos.** Cada uno de los tres campos se vuelca con guarda propia (`:36-38`); si la observación histórica tiene `Fecha` no fecha, `TIPO` vacío u `Observacion` vacía, el `TextBox` queda en blanco en vez de elevar error.
- **`AllowEdits = False` incondicional.** Se aplica antes de la carga de datos (`:27`); cualquier intento posterior de escribir en los `TextBox` no se ejecuta a nivel de VB. La propiedad `Locked = NotDefault` del control `TIPO` (`.form.txt:1044`) refuerza esa garantía solo en ese campo.
- **Sin eventos propios.** No hay `RaiseEvent` ni `WithEvents`: el contenedor que abrió el visor no recibe notificación de cierre. La coordinación se hace por código del contenedor tras `DoCmd.Close` del formulario hijo.
- **Sin caché invalidada.** El visor no modifica `m_ObjUsuarioActivo` ni `m_ObjObservacionActiva`; los globales quedan con su valor previo.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga de la observación histórica | `Constructor.getObservacionHistorica` | query/lectura en `modules/Constructor.bas:700` | `:29` |

Sin escritura. Sin transacciones. Sin correo.

## Coexistencia

Pendiente de `../coexistencia.md`. Cosas que ya se ven: el visor reutiliza el layout y los nombres de control de `Form_FormUsuario02DatosObservacion` (mismo PRD), así que tres campos `Fecha`/`Observacion`/`TIPO` se duplican entre los dos formularios; ambos comparten `EncabezadoDelFormulario` y `PieDelFormulario` con la misma estructura. En web conviene no doblar la maquetación: un mismo componente `ObservacionLectura` con una prop `editable` cubriría las dos vistas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_Titulo` se declara con `Dim` (`:21`) pero nunca se le asigna valor: ni en alta (no hay alta aquí) ni en el flujo de apertura. La variable queda con `""` durante toda la vida del formulario, sin afectar porque el título fijo lo escriben `:50-51` directamente.
- `Constructor.getObservacionHistorica(m_Argumento, m_Error)` recibe `m_Argumento` sin normalizar; si la codificación histórica del `IDObservacion` cambia, este visor queda ligado al formato actual. A verificar el contrato estable de la factory (`Constructor.bas:700`).
- No se llama a `DatosDelObjetoAlFormulario` (helper de `FG.bas:4180`); el volcado es manual (`:36-38`). Si se añaden campos a `ObservacionHistorica` (p. ej. `AltaPor`, `FechaAlta`), el visor los ignora.
- `Me.Caption = "Observación Histórica"` (`:50`) **no se internacionaliza** y deja el título de la ventana fijo; el `lblTitulo.Caption = Me.Caption` (`:51`) hereda ese literal.
- El campo `Fecha` se asigna solo si `IsDate(...)` (`:36`); si no, queda vacío y el usuario no ve motivo. Mismo problema para `TIPO` (`:37`) y `Observacion` (`:38`).
- Si `Constructor.getObservacionHistorica` devuelve `m_Error <> ""` pero el objeto no es `Nothing`, el código sigue (`:29-30`); el guard `Is Nothing` (`:32`) no se cruza con el error y el visor entra en `:36-38` con un objeto posiblemente a medias.
- `Me.AllowEdits = False` se aplica antes de la carga de datos (`:27`); si la carga deja los campos vacíos, el usuario no ve por qué y no hay forma de pedir el dato (no hay `Tag` ni `ControlTipText` específico).
- No hay `MensajeInfo` ni `MsgBox` posterior a la carga: si los tres campos quedan en blanco por tolerar vacío, no se informa al usuario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: componente reutilizable `ObservacionLectura` para actual e histórico; normalizar el `IDObservacion` en el contrato del servicio; mostrar un mensaje si la observación histórica no se pudo resolver por completo; traducir el literal del título al locale de la sesión.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs`, o con un argumento que `Constructor.getObservacionHistorica` no pueda resolver, muestra «Se ha abierto con parámetros insuficientes» y cierra el visor sin pintar los campos.
2. Los tres campos se llenan solo si su dato es utilizable (`IsDate`, no vacío); cualquier otro caso deja el `TextBox` en blanco.
3. El visor nunca escribe: `Me.AllowEdits = False`, el botón «Salir» no pregunta nada y no se llama a `Registrar`.
4. El título y la barra de título muestran «Observación Histórica».
5. La cabecera y el pie se pintan con `m_ObjEntorno.CabeceraBackColor`.
6. Un error durante la apertura se traduce en un `MsgBox` con prefijo «Al abrir el formulario …» y cierra el visor.

## Pendientes

- Sentencia exacta que ejecuta `Constructor.getObservacionHistorica` (`modules/Constructor.bas:700`) y formato del `IDObservacion` que espera.
- Forma del argumento `OpenArgs` en el flujo real: el visor confía en la factory, pero no la verifica.
- Si existe alguna ruta en el binario que abra este formulario sin pasar por `Constructor` (a verificar).
- Firma exacta de `CopiarEstiloBoton` (`FG.bas:2435`) en esta llamada: se invoca con `ByRef p_Error` (`:40`), pero el `Dim` lo llama con la firma de un solo argumento en otros formularios del PRD (a verificar).
- `m_Titulo` declarado (`:21`) y nunca usado: rastro a eliminar o reutilizar.
