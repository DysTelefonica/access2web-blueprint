# Form_FormLogin

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormLogin.cls` (311 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de acceso a la Lanzadera: pide correo y contraseña, valida el usuario contra `TbUsuariosAplicaciones` y abre el menú principal que corresponde a su rol. También permite solicitar una contraseña de un solo uso. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormLogin.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Genreales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalAdmin`, `Form_FormMenuPrincipalUsuario`, `Form_FormMenuPrincipalCalidad` | Cierre de sesión: `DoCmd.OpenForm "FormLogin"` | `Form_FormMenuPrincipalAdmin.cls:97`, `Form_FormMenuPrincipalUsuario.cls:127`, `Form_FormMenuPrincipalCalidad.cls:140` |
| `Form_FormUsuarioCambioPass` | Tras cambiar la contraseña, lo reabre y vacía `Pass` | `Form_FormUsuarioCambioPass.cls:51-52` |
| `Form_FormObtenerContraseña` | Tras enviar la contraseña de un solo uso | `Form_FormObtenerContraseña.cls:46` |

Abre `FormUsuarioCambioPass` (`:131, 158`) y, vía `LoginCorrecto`, `FormMenuPrincipalAdmin`, `FormMenuPrincipalCalidad` o `FormMenuPrincipalUsuario` (`VG.bas:724-736`). Cierra la aplicación con `DoCmd.Quit` en varios caminos (`:114, 171, 195, 236, 260`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CorreoUsuario` | texto | — | Correo; se precarga desde el fichero de sesión local | `:241-244`, `.form.txt:1233` |
| `Pass` | texto (`InputMask = Password`) | `KeyDown` | Con Intro lanza `ComandoIniciarSesion_Click` | `:299-311`, `.form.txt:1264-1267` |
| `ComandoIniciarSesion` («Login») | botón | `OnClick` | Valida y autentica (ver Flujos) | `:58-185`, `.form.txt:1294-1296` |
| `ComandoCambioDeContrasenia` («Obtener») | botón | `OnClick` | Genera y envía por correo una contraseña de un solo uso | `:8-53`, `.form.txt:1605-1607` |
| `ComandoSalir` | botón | `OnClick` | Confirma y cierra Access con `acQuitSaveAll` | `:189-196`, `.form.txt:1515-1517` |
| `ComandoAyudaFormulario` | botón | `OnClick` | `AbrirArchivoAyuda` busca el documento en `TbHerramientaDocAyuda` | `:264-267`, `FG.bas:538-554` |
| `lblEstado` | etiqueta oculta | — | Sin uso en el código | `.form.txt:1473-1486` |
| `CuadroSuperior`, `Cuadro67` | rectángulos | `OnDblClick` declarado | Sin handler en el `.cls` | `.form.txt:1386, 1458` |

## Flujos

### Apertura (`Form_Load`, `:200-261`)

1. `Ajustar Me` y oculta la cinta (`:211-212`).
2. `EVEMin` reinicia globales, crea el `Entorno`, vacía el usuario conectado y lee el correo del fichero de sesión local (`VG.bas:669-682`; lectura en `FG.bas:655-678`).
3. Si `AperturaSinLogin` es cierto —parámetro `PermitirDiaSinContrasenia` activo, mismo identificador de equipo y fichero del mismo día (`FG.bas:853-886`, `:779-845`)—, `EVEMin` carga el usuario por correo y ejecuta `EVE` (`VG.bas:686-705`); el formulario llama a `LoginCorrecto` **sin pedir contraseña** (`:218-228`).
4. Comprueba que el componente de cifrado responde calculando un SHA-256 de prueba; si devuelve vacío, avisa de que falta .NET Framework 3.5 y cierra Access (`:229-238`).
5. Precarga `CorreoUsuario` si el correo del fichero contiene «@» y pone el foco en `Pass` (`:241-244`).

### Inicio de sesión (`ComandoIniciarSesion_Click`, `:58-185`)

1. Rechaza si `NumeroVecesIntentado > 5` (`:71-74`), contraseña vacía (`:75-78`) o correo sin «@» (`:79-82`).
2. Ejecuta `EVE` si no se ha hecho (`:84-90`).
3. Obtiene el usuario de Windows conectado (`m_ObjEntorno.UsuarioConectadoMaquina`, que busca por `UsuarioRed`: `Entorno.cls:792`, `Constructor.bas:11`) y su condición de administrador (`:91-100`).
4. Carga el usuario por correo con `Constructor.getUsuario` (`:103`; SQL en `Constructor.bas:46-73`).
5. Si no existe, avisa «Está intentando iniciar sesión con un correo electrónico no registrado…» y cierra Access (`:108-116`).
6. Si el usuario de Windows es administrador, el estado es `Correcto` **sin comprobar la contraseña** (`:117-118`); si no, llama a `Login` (`:120`).
7. Según el estado: `caducada` o `NecesitaCambio` ⇒ aviso, cierra el formulario y abre `FormUsuarioCambioPass` (`:127-134, 154-161`); `Correcto` ⇒ `SetSSOFAArchivo` escribe el fichero de sesión local y `LoginCorrecto` abre el menú (`:135-149`); `Incorrecto` ⇒ error «Usuario/Pass no válida…» (`:150-153`); `Bloqueada` ⇒ correo de cuenta bloqueada, aviso y cierre de Access (`:162-173`).

### Obtener contraseña (`ComandoCambioDeContrasenia_Click`, `:8-53`)

Carga el usuario por correo; si no existe, «No hay registrado ningún usuario con ese correo…» (`:22-29`). `ResetearPass` (`Usuario.cls:729-783`) genera una contraseña aleatoria de 8 caracteres, guarda su hash en `Password`, marca `TieneQueCambiarLaContrasenia` y guarda también la contraseña **en claro** en `PassIncialPlana` (`Usuario.cls:740, 757-765`); después la envía por correo al usuario (`Usuario.cls:786-829`). Muestra el aviso de los 10 minutos (`:41`).

### Errores

Errores internos (1000) muestran `m_Error`; los demás se componen, se envían con `CorreoAlAdministrador` y se muestran como `MsgBox` crítico (`:45-52, 175-184`). En `Form_Load` cualquier error cierra Access (`:251-260`).

## Reglas, validaciones y cálculos

- **Función `Login`** (`FG.bas:1345-1446`): compara el SHA-256 de la contraseña con `Password` (`:1362-1368`); si falla devuelve `Incorrecto` y suma `m_ObjEntorno.NumeroLogins`; con más de 4 fallos llama a `Usuario.Bloquear` y pone el contador a 0 (`FG.bas:1369-1378`). Con contraseña correcta y `FechaBloqueo` informada devuelve `Bloqueada` o desbloquea (`FG.bas:1383-1394`). Si `TieneQueCambiarLaContrasenia`, compara la contraseña tecleada con `PassIncialPlana` en claro: igual ⇒ `NecesitaCambio` (`FG.bas:1413-1424`).
- **Bloqueo:** `Bloquear` pone `FechaBloqueo = Now()` y `Activado = false` filtrando por `UsuarioRed` (`Usuario.cls:687-692`); `DesBloquear` hace lo inverso (`Usuario.cls:712-716`).
- **Cuenta bloqueada:** `CorreoCuentaBloqueada` envía un aviso a un destinatario fijo en el código (`FG.bas:1311-1331`).
- **Menú por rol:** administrador ⇒ `FormMenuPrincipalAdmin`; calidad ⇒ `FormMenuPrincipalCalidad`; resto ⇒ `FormMenuPrincipalUsuario` (`VG.bas:724-730`).
- **Fichero de sesión:** `SetSSOFAArchivo` sustituye el fichero local por una línea con identificador de equipo, correo y fecha (`FG.bas:979-989`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario por correo o por usuario de red | `TbUsuariosAplicaciones` | SQL dinámico concatenado | `Constructor.bas:46-73` |
| Bloqueo y desbloqueo | `TbUsuariosAplicaciones` (`FechaBloqueo`, `Activado`) | `UPDATE` dinámico | `Usuario.cls:688-692, 712-716` |
| Reseteo de contraseña | `TbUsuariosAplicaciones` (`Password`, `TieneQueCambiarLaContrasenia`, `PassIncialPlana`) | DAO `Edit/Update` | `Usuario.cls:748-765` |
| Documento de ayuda | `TbHerramientaDocAyuda` | SQL dinámico | `FG.bas:538-541` |
| Fichero de sesión local | fichero de texto | `FileSystemObject` | `FG.bas:984-991` |

Sin transacciones. Envío de correo por la clase `Correo` (`FG.bas:1326-1331`; `Usuario.cls:820-825`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con el correo tecleado (`Constructor.bas:70`); literal de fecha Jet `#…#` en `Bloquear` (`Usuario.cls:688`); booleanos `true/false` de Access (`Usuario.cls:689, 713`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Si el usuario de Windows es administrador, cualquier correo registrado entra sin contraseña (`:117-118`).
- El estado `caducada` nunca se devuelve: `Login` sale en `FG.bas:1409-1411` o `:1413-1424` antes de evaluar la caducidad (`FG.bas:1425-1437`); la rama `:127-134` es inalcanzable.
- `NumeroVecesIntentado` nunca se incrementa; el bloqueo de `:71-74` no actúa.
- Se lee `m_ObjUsuarioConectadoLogin.CorreoUsuario` (`:107`) antes de comprobar `Is Nothing` (`:108`): un correo no registrado daría error 91 en vez del aviso.
- `DateDiff("m", Now(), FechaBloqueo) < 20` (`FG.bas:1385`) usa meses y orden invertido: el resultado es ≤ 0 y el bloqueo no caducaría.
- El quinto fallo bloquea pero devuelve `Incorrecto` (`FG.bas:1369-1375`); el bloqueo filtra por `UsuarioRed` (`Usuario.cls:690`), que podría estar vacío.
- La contraseña de un solo uso se guarda en claro y se compara en claro (`Usuario.cls:763`; `FG.bas:1418`).
- Cualquiera que conozca un correo registrado puede resetear su contraseña sin confirmación (`:22-31`).
- `LoginCorrecto` compara `Err.Number <> ""` en su manejador (`VG.bas:739`).
- Restos: variables sin uso (`:59, 65`) y función `EstablecerInicio` sin llamadas (`:268-297`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: autenticación corporativa (SSO) en vez de contraseña propia; bloqueo por usuario con caducidad en minutos; enlace de restablecimiento con caducidad en lugar de contraseña por correo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Correo sin «@» o contraseña vacía muestran su mensaje y no consultan la base.
2. Un correo no registrado muestra el aviso de cuenta no registrada.
3. Con contraseña correcta, el usuario llega al menú de su rol (administrador, calidad o usuario).
4. Con contraseña de un solo uso, el usuario debe cambiarla antes de entrar.
5. Tras el número de fallos acordado, la cuenta queda bloqueada y se avisa al administrador.
6. «Obtener» envía una contraseña de un solo uso al correo registrado.

## Pendientes

- Formulario de inicio de la base (propiedad de arranque o `AutoExec`): no está en `src`.
- Contenido de `getSSID` y formato del identificador de equipo (`FG.bas:722`).
- Contenido de `EVE` por cada propiedad del `Entorno` (`VG.bas:606-637`).
- Si `ResetGlobals` reinicia `NumeroLogins` (`VG.bas:593, 670`).
