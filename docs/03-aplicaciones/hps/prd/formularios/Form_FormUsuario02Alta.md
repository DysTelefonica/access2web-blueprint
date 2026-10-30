# Form_FormUsuario02Alta

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario02Alta.cls` (423 líneas) y `.form.txt` (2744 líneas) · **PRD:** [README](../README.md)

Subpantalla de alta de un usuario: crea un `UsuarioHPS` nuevo con `ParaAlta = Sí`, edita los campos en pantalla, registra el alta en `TbUsuarios`, crea hasta dos observaciones (`General` y `Curso`), regenera la caché y reentra en `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` para abrir el detalle del recién creado. Es **no enlazado** (el `.form.txt` no declara `RecordSource` textual; los `Tag = "DATOS"` de `.form.txt:700-2166` son los que `RellenarObjetoConDatos`/`RellenarCombo*`/`FG.bas:4124-4180` leen como puente entre controles y propiedades). Sister de `Form_FormUsuario02HPS` (`src/forms/Form_FormUsuario02HPS.cls`, 988 líneas, edición con foco HPS) y `Form_FormUsuario02Edicion` (`src/forms/Form_FormUsuario02Edicion.cls`, 977 líneas, edición general), todavía sin ficha propia. Las citas son `Form_FormUsuario02Alta.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `Constructor.bas` es `modules/Constructor.bas` y `VariablesEntorno.bas` es `modules/VariablesEntorno.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal.ComandoNuevoUsuario_Click` | `DoCmd.OpenForm "FormUsuario02Alta"` (sin `OpenArgs`) tras vaciar `m_ObjUsuarioActivo` | `Form_FormInicial02ConsultasPrincipal.cls:141-160` |
| `m_ObjUsuarioActivo` (global) | `Form_Load` lo crea con `New UsuarioHPS` y `ParaAlta = EnumSiNo.Sí` | `:347-348`, `VariablesEntorno.bas:225` |
| `m_ObjUsuarioAlInicio` (global) | Lo lee `DNI_Exit` para distinguir el DNI original del modificado (no se declara en este `.cls`: es el global `VariablesEntorno.bas:245`) | `:300-301`, `VariablesEntorno.bas:245` |
| `m_FormExpedientes.Seleccionado` (contenedor `Form_FormExpedientesBusqueda`) | `WithEvents` que captura la fila elegida y vuelca los datos al formulario | `:388-422`, `:8` |

Al guardar cierra con `DoCmd.Close acForm, Me.Name, acSaveNo` y dispara `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`:255`), que abre el recién creado como edición. El formulario **no** publica eventos propios; toda la cadena inversa viaja por el global `m_ObjUsuarioActivo`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Título centrado, fijado en diseño a «ALTA DE USUARIO» | `:334`, `.form.txt:597-603` |
| `Nombre`, `DNI`, `Correo_e`, `F_Nacimiento`, `F_Curso`, `Telefono`, `Apellido_1`, `Apellido_2`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `Motivo_HPS`, `CursoEnVigor`, `FechaPrimeraConvocatoria`, `FechaSegundaConvocatoria`, `FechaCorreoNoCurso`, `Observaciones`, `Obs_Curso`, `PalabraClave`, `IDExpediente`, `Titulo`, `LugarNacimiento`, `CadenaContratistas` | `TextBox` y `ComboBox` con `Tag = "DATOS"` (salvo `PalabraClave` / `Titulo` / `CadenaContratistas`) | — | Datos del usuario + HPS; `RellenarObjetoConDatos m_ObjUsuarioActivo, Me` (`:206`; `FG.bas:4124`) los escribe al pulsar «Grabar» | `:206, :347-380`, `.form.txt:700-2166` (controles `Tag = "DATOS"` en `:701, :761, :821, :911, :971, :1030, :1119, :1148, :1244, :1284, :1380, :1513, :1575, :1636, :1697, :1314, :1343, :2077, :2167`) |
| `LBLDNI`, `lblExpedienteActividad` y resto de `Etiqueta*` | etiquetas | — | Rótulos («Empresa_Usuario», «Nombre», «DNI», …) | `.form.txt:672, 730, 642-1475` |
| `DNI` | `TextBox` | `OnExit` → `DNI_Exit` | Valida longitud 9, formato y unicidad contra `TbUsuarios` y `TbUsuariosHistoricos` | `:293-332`, `.form.txt:759` |
| `Correo_e` | `TextBox` | `OnExit` → `Correo_e_Exit` | Valida formato de correo cuando se rellena con algo | `:269-291`, `.form.txt:819` |
| `PalabraClave` | `TextBox` | `OnKeyDown` → `PalabraClave_KeyDown` | `Enter` dispara `ComandoBuscarExpediente_Click` | `:144-161`, `.form.txt:1794` |
| `ComandoBuscarExpediente` | botón (icono «Lupa») | `OnClick` → `ComandoBuscarExpediente_Click` | Llama a `Constructor.getExpedientesBusqueda` con la palabra clave | `:14-103`, `.form.txt:1820` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` → `ComandoIrABusquedaExpedientes_Click` | Abre `FormExpedientesBusqueda` con la palabra clave actual y aplica `AvanceCerrar` | `:105-142`, `.form.txt:1908` |
| `ComandoActualizar` | botón (icono «Actualizar») | `OnClick` → `ComandoActualizar_Click` | Vacía `m_ObjEntorno.Expedientes` para forzar recarga | `:10-12`, `.form.txt:1977` |
| `ComandoEmpresas1` | botón «Editar» junto a `IDEmpresaUsuario` | — | Sin handler declarado en `.cls`; declarado en `.form.txt:1406` | `.form.txt:1406-1475` |
| `ComandoMotivos` | botón «Editar» junto a `Motivo_HPS` | `OnClick` → `ComandoMotivos_Click` | Pide un motivo nuevo por `InputBox` y lo inserta en la lista | `:167-189`, `.form.txt:1724` |
| `ComandoGrabar` («Grabar») | botón | `OnClick` → `ComandoGrabar_Click` | Da de alta al usuario + observaciones y cierra | `:194-267`, `.form.txt:2233` |
| `cmdSalir` («Salir») | botón | `OnClick` → `CmdSalir_Click` | `DoCmd.Close acForm, Me.Name, acSaveYes` (sin manejador de errores) | `:163-165`, `.form.txt:2276` |
| (formulario) | — | `OnLoad` → `Form_Load` | Decide editable, instancia el `UsuarioHPS`, rellena combos y aplica estilo | `:334-383`, `.form.txt` |

Sin ribbon; `NavigationButtons = NotDefault`, `RecordSelectors = NotDefault`, `ControlBox = NotDefault`, `CloseButton = NotDefault` (`.form.txt:5-9`). Cabecera (`EncabezadoDelFormulario`) y pie (`PieDelFormulario`) sin lógica (`.form.txt:579, 2217`); `lblTitulo` con `Caption = "ALTA DE USUARIO"` (`.form.txt:600`).

## Flujos

### Apertura (`Form_Load`, `:334-383`)

1. `m_Error = ""` y `AjustarTamaño Me` (`:336-339`; helper en `FG.bas:901`).
2. Si `Application.TempVars("SoloLectura") = "Sí"`, `Me.AllowEdits = False`; si no, `Me.AllowEdits = True` (`:341-346`).
3. `Set m_ObjUsuarioActivo = New UsuarioHPS` y `m_ObjUsuarioActivo.ParaAlta = EnumSiNo.Sí` (`:347-348`): instancia el global como `UsuarioHPS` en modo alta.
4. `EstablecerControlesDatos Me, m_Error` (`:351`; `FG.bas:1005-1055`) con `Err.Raise 1000` encadenado (`:352-354`).
5. Sincroniza `ComandoGrabar.Enabled = Me.AllowEdits` (`:357-361`): «Grabar» sigue el modo lectura/edición del formulario.
6. `RellenarComboEmpresas Me.IDEmpresaUsuario, EnumSiNo.No, , m_Error` y `RellenarComboEmpresas Me.IDEmpresaHPS, EnumSiNo.Sí, , m_Error` (`:363-371`; `FG.bas:1100-1180`) con raise encadenado en cada uno.
7. `RellenarComboMotivos Me.Motivo_HPS, m_Error` (`:373-378`; `FG.bas:1181-1205`) con raise encadenado.
8. `CopiarEstiloBoton Me.ComandoGrabar` y `CopiarEstiloBoton Me.cmdSalir` (`:380-381`; `FG.bas:2435`): **llamada con un solo argumento**, a diferencia del patrón `ByRef p_Error` del PRD vecino (`Form_FormInicial02ConsultasPrincipal.cls:193-204`).
9. En error: «Al abrir el formulario se ha producido el error: …», `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox` y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:376-382`).

### Buscar expediente (`ComandoBuscarExpediente_Click`, `:14-103`)

1. `m_PalabraClave = Nz(Me.PalabraClave, "")` (`:24`); vacío ⇒ sale con `Hourglass False` (`:25-31`). Es la única validación de entrada: no exige longitud mínima.
2. `Set m_Col = Constructor.getExpedientesBusqueda(p_PalabraClave:=m_PalabraClave, p_Error:=m_Error)` (`:34`; SQL en `Constructor.bas:3340-3396`); raise 1000 si dejó error.
3. **Rama `m_Col Is Nothing`** (`:38-56`): vacía `IDExpediente` y `Titulo`, ajusta `ComandoActualizar.Enabled` según el estado y llama a `ComandoIrABusquedaExpedientes_Click` (que abre el buscador avanzado).
4. **Rama `m_Col.Count = 1`** (`:58-83`): para el único expediente, concatena `CodExp & " " & Titulo` (si existe) o `CodExp & " " & TextoExpediente` (si no) en `Me.Titulo`, copia `CadenaContratistas` y `IDExpediente`, y vacía `PalabraClave`.
5. **Rama más de un resultado** (`:86-95`): abre `FormExpedientesBusqueda`, le pasa la `PalabraClave` por `m_FormExpedientes.PalabraClave = …` y dispara `Form_FormExpedientesBusqueda.Filtrar` (vía referencia estática al formulario). El resultado lo devuelve el evento `m_FormExpedientes_Seleccionado` (`:388-422`).
6. En error: `Hourglass False`, mensaje, `CorreoAlAdministrador` y `MsgBox` (vbCritical para errores no 1000; vbExclamation para 1000) (`:93-103`).

### Ir al buscador (`ComandoIrABusquedaExpedientes_Click`, `:105-142`)

Cierra `FormExpedientesBusqueda` si está abierto, lo vuelve a abrir, le pasa la `PalabraClave` actual y aplica `AvanceCerrar` (`:129`; helper en `FG.bas:4415-4421` que oculta `lblEstado` del formulario activo). El `WithEvents` se reapunta aquí (`:126`). Misma estructura de errores que `ComandoBuscarExpediente_Click`.

### Enter sobre la palabra clave (`PalabraClave_KeyDown`, `:144-161`)

`KeyCode = vbKeyReturn` ⇒ foco a `ComandoBuscarExpediente` y de vuelta a `PalabraClave` (efecto nulo para el usuario, ver `:151-152`), `ComandoBuscarExpediente_Click` y `KeyCode = 0` para anular el pitido de Access. Si el botón «Buscar» está deshabilitado (modo lectura, por ejemplo), sale sin tocar nada (`:146-149`).

### Motivo nuevo (`ComandoMotivos_Click`, `:167-189`)

`InputBox("Introduzca el motivo si no existe en la lista anterior", "Motivo nuevo", "")`; cadena vacía ⇒ sale. Si no, `MeterUnMotivoEnListaMOtivos m_Motivo, m_Error` (`:177`; `FG.bas:4234-4268`), `Me.Motivo_HPS.Requery` y `Me.Motivo_HPS = m_Motivo` (`:180-183`).

### Salir (`CmdSalir_Click`, `:163-165`)

Una sola línea: `DoCmd.Close acForm, Me.Name, acSaveYes`. **Sin** manejador de errores ni validación de cambios pendientes — peculiar frente a `Form_FormUsuario02DatosObservacion` (`Form_FormUsuario02DatosObservacion.cls:14-51`), que pregunta «¿Desea guardar los cambios antes de cerrar?».

### Validar correo (`Correo_e_Exit`, `:269-291`)

Si `Nz(Me.Correo_e, "") <> ""` y `FormatoDatoValido(UCase(m_Valor), "email", m_Error)` devuelve `False` (helper en `FG.bas:2238-2262`), muestra `vbCritical` con el texto «Formato de correo electrónico incorrecto» y sale sin modificar nada (`:277-285`). Si el control queda vacío, sale en silencio (sin validación).

### Validar DNI (`DNI_Exit`, `:293-332`)

1. Si `m_ObjUsuarioAlInicio Is Nothing` (alta pura, sin snapshot), no valida nada (`:300`).
2. Si `m_ObjUsuarioAlInicio.DNI <> m_Valor`, exige longitud 9 y `FormatoDatoValido(UCase(...), "DNI", …)` (`:301-309`; helper en `FG.bas:2238-2262`).
3. Comprueba colisión con usuarios activos: `m_UsuarioConDNI = DNIEnActivo(m_Valor, m_Error)` (`:311`; `FG.bas:2737-2769`); si no vacío, `Me.DNI = Null` y eleva «Hay un usuario registrado con ese DNI …» (`:312-318`).
4. Repite con históricos: `DNIEnHistorico(m_Valor, m_Error)` (`:319`; `FG.bas:2772-2802`); mismo manejo.

### Guardar (`ComandoGrabar_Click`, `:194-267`)

1. `RellenarObjetoConDatos m_ObjUsuarioActivo, Me, m_Error` (`:206`; `FG.bas:4124-4146`) con raise 1000.
2. `m_ObjUsuarioActivo.Registrar , m_Error` (`:212`; `UsuarioHPS.cls:1254-1486`), **sin** pasar `p_ObjUsuarioHPSAlInicio`. En alta la firma lo sobreentiende: si `ParaAlta = Sí` no se usa el snapshot.
3. `m_ID = m_ObjUsuarioActivo.ID` (`:215`), el ID que acaba de asignar `Registrar`.
4. Si `Nz(Me.Observaciones, "") <> ""` (`:218`): crea un `Observacion` nuevo, `ParaAlta = Sí`, `ID = m_ID`, `TIPO = "General"`, `Observacion = m_Observacion`, `Fecha = Date` y `Registrar m_Error` (`:219-234`).
5. Si `Nz(Me.Obs_Curso, "") <> ""` (`:236`): mismo patrón con `TIPO = "Curso"` y `Observacion = m_ObservacionCurso` (`:237-249`).
6. `cache_usuario_regenerar p_ID:=m_ObjUsuarioActivo.ID, p_Error:=m_Error` (`:248`; `modules/cacheUsuario.bas:65-228`): si falla, raise 1000.
7. `pregunta = MsgBox("Alta correcta", vbInformation, "Alta de Usuario")` (`:251`) — el retorno se guarda en el global `VariablesEntorno.bas:220` y no se usa.
8. `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`:255`; `Form_FormInicial02ConsultasPrincipal.cls:44-68`): abre el recién creado.
9. `DoCmd.Close acForm, Me.Name, acSaveNo` (`:258`).
10. En error: `Hourglass False`, mensaje, `CorreoAlAdministrador`, `MsgBox` **sin cerrar** (`:260-266`).

### Recibir selección del buscador (`m_FormExpedientes_Seleccionado`, `:388-422`)

Vuelca `m_Exp.IDExpediente`, `m_Exp.CodExp & " " & m_Exp.Titulo` (o `.TextoExpediente` si aquél está vacío), `m_Exp.CadenaContratistas` a los controles de la derecha y ajusta `ComandoActualizar.Enabled` (`:393-414`). La firma del evento la publica `Form_FormExpedientesBusqueda` (no en `src/`, a verificar). En error: `Hourglass False`, mensaje, `CorreoAlAdministrador`, `MsgBox` (vbCritical/vbExclamation, igual que en `:93-103`).

## Reglas, validaciones y cálculos

- **Alta se decide en `Form_Load`.** `Set m_ObjUsuarioActivo = New UsuarioHPS` + `ParaAlta = Sí` (`:347-348`); no hay parámetro `OpenArgs`, modo ni ruta: el formulario siempre arranca como alta cuando el contenedor lo invoca.
- **`ComandoGrabar.Enabled` sigue a `AllowEdits`.** Con `SoloLectura="Sí"`, `Form_Load` deja `ComandoGrabar.Enabled = False` (`:357-361`); el resto de controles siguen editables salvo los que `Me.AllowEdits = False` deshabilita a nivel de Access.
- **`SoloLectura` por `TempVars`.** `Application.TempVars("SoloLectura") = "Sí"` (`:341-345`): global de sesión; misma fuente que `Form_FormInicial02ConsultasPrincipal.cls:181` y `Form_FormUsuario02DatosObservacion.cls:122-126`. Sin equivalente directo en web — se traduce a un parámetro de ruta o un claim de rol.
- **DNI solo se valida si el alta ya tenía uno.** El snapshot vive en el global `m_ObjUsuarioAlInicio` (`VariablesEntorno.bas:245`), no en un campo de este `.cls`: si el flujo lo dejó `Nothing` (alta pura sin edición previa), `DNI_Exit` no ejecuta ninguna validación (`:300`).
- **DNI único contra activos e históricos.** `DNIEnActivo` y `DNIEnHistorico` (`FG.bas:2737-2802`) hacen `SELECT` sobre `TbUsuarios` y `TbUsuariosHistoricos` y devuelven `Nombre & Apellido_1 & Apellido_2` del conflicto, que se reusa como cuerpo del mensaje. El DNI escrito por el usuario se borra (`Me.DNI = Null`) antes de elevar el error (`:316, :323`).
- **`Observaciones` crea SIEMPRE un `Observacion` con `Fecha = Date`.** No hay fecha editable: el `.form.txt` no expone un `Fecha` para la observación desde este formulario; `Fecha = Date` se fija en código (`:229`). Si un alta queda para grabar al día siguiente, la observación lleva la fecha del día del alta, no la fecha real del hecho.
- **Caché tras alta.** `cache_usuario_regenerar` (`:248`; `cacheUsuario.bas:65-228`) reescribe `TbUsuariosEntidades` y `TbDatosLocal` desde el `DatosLocal` reconstruido por `getDatosLocalDeUsuario`. La invalidación cubre las cachés que consumen los listados del menú izquierdo y los filtros posteriores; no hay un `m_ObjColParaFiltros = Nothing` aquí.
- **Vuelta a edición al guardar.** `ComandoGrabar_Click` invoca `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`:255`), que cierra y reabre `FormUsuario00Principal` con `m_ObjUsuarioActivo` ya rellenado por `Registrar`. El usuario recién dado de alta entra como edición sin que el operador tenga que seleccionar nada.
- **`CopiarEstiloBoton` llamado con un solo argumento.** A diferencia del patrón `ByRef p_Error` del PRD vecino, este `.cls` llama con un solo argumento (`:380-381`): legal en VB, pero el control sobre errores del helper no se propaga al formulario. La versión con dos argumentos está documentada en `FG.bas:2435-2499`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Relleno del combo de empresa del usuario | `getExpedientes`/`TbSuministradores` (vía `RellenarComboEmpresas`) | `SELECT` (a verificar el `SQL` exacto en `FG.bas:1100-1180`) | `:363-365`, `FG.bas:1100-1180` |
| Relleno del combo de empresa HPS | `TbSuministradores` con `EnumSiNo.Sí` (tramitadoras) | `SELECT` | `:368-370`, `FG.bas:1100-1180` |
| Relleno del combo de motivos | `m_ObjEntorno.ColMotivos` (en memoria) | colección | `:373-375`, `FG.bas:1181-1205` |
| Búsqueda de expedientes por texto | `Constructor.getExpedientesBusqueda` | `SELECT` con `LIKE` en `Constructor.bas:3340-3396` | `:34`, `Constructor.bas:3340-3396` |
| Volcado de la selección del buscador | `m_FormExpedientes.Seleccionado` | evento VBA | `:388-422` |
| Validación de unicidad del DNI | `TbUsuarios` y `TbUsuariosHistoricos` (vía `DNIEnActivo`, `DNIEnHistorico`) | `SELECT` | `:311, :319`, `FG.bas:2737-2802` |
| Alta del usuario | `UsuarioHPS.Registrar` | `INSERT` en `TbUsuarios` (`UsuarioHPS.cls:1254-1486`) | `:212`, `UsuarioHPS.cls:1254-1486` |
| Alta de la observación «General» | `Observacion.Registrar` | `INSERT` en `TbObservaciones` | `:219-234` |
| Alta de la observación «Curso» | `Observacion.Registrar` | `INSERT` en `TbObservaciones` | `:237-249` |
| Regeneración de caché del usuario | `cacheUsuario.cache_usuario_regenerar` | `INSERT`/`UPDATE` en `TbUsuariosEntidades` y `TbDatosLocal` | `:248`, `cacheUsuario.bas:65-228` |

Sin transacciones explícitas: el alta es una secuencia de tres `Registrar` + un `cache_usuario_regenerar`; un fallo en cualquiera de los tres primeros deja la mitad del trabajo hecho. `cache_usuario_regenerar` siempre corre tras los dos `Registrar` de observación (la condición de `:218` y `:236` no eleva hasta `cache_usuario_regenerar` falle).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Cosas que ya se ven: `Tag = "DATOS"` aparece en 19 controles de los 22 útiles del `.form.txt` (lista de líneas en la tabla de controles): cualquier migrador debe respetar esa marca como contrato del puente `RellenarObjetoConDatos`/`RellenarCombo*` (`FG.bas:4124-4180`); `DNI_Exit` confía en un global (`m_ObjUsuarioAlInicio`) que este `.cls` no declara — el alta pura entra con `Nothing` y se salta la validación por construcción; `ComandoEmpresas1` (`.form.txt:1406-1475`) está declarado como botón «Editar» sin handler en el `.cls`: código muerto a borrar en la migración; los combos `IDEmpresaUsuario` e `IDEmpresaHPS` se rellenan con la misma colección de proveedores pero parametrizada por `p_SoloTramitadoras` (`FG.bas:1100`), y la diferencia operativa entre los dos reside solo en ese flag — la forma canónica de empresa tramitadora en `TbDatosLocal` debe confirmar si se persiste por identificador o por nombre. La pareja `IDExpediente` (oculto, `.form.txt:2076`) + `Titulo` (visible, `.form.txt:2107`) sigue el patrón «id invisible + texto concatenado para el usuario», sin equivalente directo: conviene devolver un objeto `{id, label}` en lugar de una concatenación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_ObjUsuarioAlInicio` (`:300-301`) **no** se declara en este `.cls`: depende del global `VariablesEntorno.bas:245`. Si otro consumidor lo deja `Nothing`, la validación se salta enteramente (riesgo de duplicados de DNI en una misma sesión). No hay `Set m_ObjUsuarioAlInicio = New UsuarioHPS` aquí: el alta no toma snapshot.
- `m_ObjUsuarioActivo` se crea en `Form_Load` (`:347`), pero si otro consumidor (p. ej. `Form_FormInicial02ConsultasPrincipal.ComandoNuevoUsuario_Click`) ya tenía algo, este `New` lo pisa sin avisar. El intent del contenedor (`Form_FormInicial02ConsultasPrincipal.cls:148`) sí lo vacía, pero solo cuando se llega al botón «Nuevo».
- `PalabraClave_KeyDown` (`:144-161`) salta el foco a `ComandoBuscarExpediente` y de vuelta a `PalabraClave`: efecto visual nulo, y al final hace `ComandoBuscarExpediente_Click`. Si «Buscar» estuviera habilitado por error mientras se escribe, cada `Enter` relanzaría la búsqueda; sin protección, el spinner de `Hourglass` parpadea.
- `CmdSalir_Click` (`:163-165`) usa `DoCmd.Close … acSaveYes` sin `On Error GoTo`: una `acSaveYes` sobre un formulario no enlazado no tiene mucho sentido y, además, **no** pregunta por cambios pendientes — un operador puede teclear todo el alta y «Salir» la descarta sin avisar.
- `CopiarEstiloBoton` se llama con un solo argumento (`:380-381`); el patrón de dos (`ByRef p_Error`) lo aplica `Form_FormInicial02ConsultasPrincipal.cls:193-204` y `Form_FormUsuario00Principal.cls:53`. La firma con un argumento es legal en VB pero silencia el canal de error: si `CopiarEstiloBoton` rellena `p_Error`, este formulario no se entera.
- `DNI_Exit` exige `Len(m_Valor) <> 9` antes de validar el formato (`:303`); un DNI con letra se rechaza antes de pasar por `FormatoDatoValido`, lo cual es esperable, pero el orden de las validaciones (longitud, formato, activos, históricos) hace que un DNI vacío de longitud correcta falle en `:307` antes de llegar a `DNIEnActivo`.
- `ComandoBuscarExpediente_Click:38-56` cae en la rama «no hay resultados» sin distinguir «cero resultados» de «fallo del helper»: `getExpedientesBusqueda` que devuelve `Nothing` por error se trata idénticamente a un `WHERE` que no casó. La verificación depende del contrato del helper (a verificar en `Constructor.bas:3340-3396`).
- `Observaciones` y `Obs_Curso` se crean con `Fecha = Date` (`:229, :245`): si el alta queda para grabar a las 23:59 de un día, la observación queda con la fecha de hoy, no la fecha real del hecho.
- `cache_usuario_regenerar` (`cacheUsuario.bas:65-228`) abre `getdb()` y `CurrentDb()` sin cerrar explícitamente (`:96, :125-126` no se han reproducido aquí): la sesión de Access puede acumular conexiones si las altas son consecutivas.
- `pregunta = MsgBox("Alta correcta", …)` (`:251`) guarda el retorno en el global `VariablesEntorno.bas:220` y nunca se lee aquí; el operador puede pulsar «Aceptar» o «Cancel» y el formulario sigue adelante con el `ComandoEditarUsuario_Click`.
- `m_ObjUsuarioActivo.Registrar , m_Error` (`:212`) pasa un `,` vacío como `p_ObjUsuarioHPSAlInicio`: la firma `Registrar(Optional ByRef p_ObjUsuarioHPSAlInicio As UsuarioHPS, …)` (`UsuarioHPS.cls:1254`) lo tolera, pero en código nuevo conviene pasar un argumento explícito o un comentario.
- `Me.Titulo.ForeColor = Me.Correo_e.ForeColor` (`:79, :412`) fija el color del texto del expediente seleccionado al color del correo. Es una decisión de UX que confunde el `ForeColor` con un canal de estilo compartido.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tomar el snapshot del usuario en `Form_Load` y declararlo en el `.cls` para que la validación de DNI no dependa de un global; sustituir `palabraClave_KeyDown` por un `@submit` y eliminar el ping-pong de focos; cambiar `CmdSalir_Click` por un flujo con `dirty state` y confirmación («¿Desea descartar el alta?»); reemplazar la concatenación `CodExp & " " & Titulo` por un objeto devuelto por backend; convertir `Observaciones.Fecha = Date` en un campo editable con valor por defecto; sustituir `cache_usuario_regenerar` por una suscripción al alta vía el backend en lugar de un helper VBA; pasar la observación «General» / «Curso» como un array de items (`[{tipo, texto, fecha}]`) en lugar de dos campos separados; reemplazar `MsgBox` por toasts no bloqueantes; traducir `SoloLectura` por un claim de sesión; consolidar `CopiarEstiloBoton` en una sola firma (`p_Error` por defecto) para evitar la doble convención.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario con `m_ObjUsuarioActivo` vacío instancia un `UsuarioHPS` nuevo con `ParaAlta = Sí` y `ComandoGrabar.Enabled = Me.AllowEdits`.
2. Con `SoloLectura="Sí"`, `ComandoGrabar.Enabled = False` y el resto de controles no admiten edición.
3. Pulsar «Buscar» sobre `PalabraClave` no vacío llama a `Constructor.getExpedientesBusqueda` y, con cero resultados, abre `FormExpedientesBusqueda`; con un resultado, lo vuelca a `IDExpediente`/`Titulo`/`CadenaContratistas`; con varios, abre el buscador y deja la selección por evento.
4. «Editar» junto a `Motivo_HPS` pide un motivo por `InputBox`, lo añade a la colección y lo deja seleccionado en el combo.
5. Al pulsar «Grabar», el usuario nuevo se registra en `TbUsuarios`, se crean hasta dos observaciones (`General` y `Curso`) si los campos contienen texto, y se regenera la caché.
6. Tras el alta, el formulario dispara `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` y luego se cierra.
7. Un DNI introducido con longitud distinta de 9 o con formato inválido bloquea la salida del control con `MsgBox`; un DNI que ya existe en `TbUsuarios` o `TbUsuariosHistoricos` borra el campo y eleva error.
8. Un correo con formato inválido se rechaza al `Exit`; un correo vacío no se valida.
9. «Salir» cierra el formulario con `acSaveYes` sin preguntar nada.

## Pendientes

- Sentencia `SQL` exacta de `Constructor.getExpedientesBusqueda` (`Constructor.bas:3340-3396`) y de `DNIEnActivo`/`DNIEnHistorico` (`FG.bas:2737-2802`): no se han reproducido aquí para evitar duplicar la fuente.
- Firma exacta de la publicación `Seleccionado(m_Exp As Expediente)` en `Form_FormExpedientesBusqueda`: vive fuera de `src/` y no se ha verificado.
- Confirmar si `cache_usuario_regenerar` (`cacheUsuario.bas:65-228`) cierra las conexiones que abre o las deja vivas.
- Definición de `m_ObjEntorno.ColMotivos` en `classes/Entorno.cls` (no en `src/`): de dónde sale la lista de motivos que se muestra al usuario y `MeterUnMotivoEnListaMOtivos` (`FG.bas:4234`) enriquece.
- Comportamiento de `copiarEstiloBoton` con la firma de un argumento vs dos: verificar que la versión compilada resuelve a la misma sub en ambos casos.
- Sentido exacto de `acSaveYes` en `CmdSalir_Click` con un formulario sin `RecordSource`: si Access ignora el flag o guarda la «modificación» del estado `ParaAlta`.
- Diferencias operativas entre `IDEmpresaUsuario` e `IDEmpresaHPS` en `TbDatosLocal`: si se persiste por identificador o por nombre, y si el cambio de razón social rompe la búsqueda retroactiva.
