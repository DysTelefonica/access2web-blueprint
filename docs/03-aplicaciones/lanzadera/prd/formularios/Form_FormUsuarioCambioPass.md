# Form_FormUsuarioCambioPass

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuarioCambioPass.cls` (111 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de cambio de contraseña del usuario que ha iniciado sesión: pide la nueva contraseña dos veces, la valida, la guarda cifrada y devuelve al usuario a la pantalla de acceso. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); trabaja sobre la global `m_ObjUsuarioConectadoLogin`. Las citas son `Form_FormUsuarioCambioPass.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Genreales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormLogin` | Contraseña de un solo uso o caducada | `Form_FormLogin.cls:131, 158` |
| `Form_FormPerfil` | Botón «Cambiar Pass» | `Form_FormPerfil.cls:201-204` |
| `Form_FormUsuarioDatosGenerales` | Botón de cambio de contraseña | `Form_FormUsuarioDatosGenerales.cls:221-224` |

Al terminar cierra `FormLogin` y los menús principales de administrador y de usuario si están abiertos, abre `FormLogin` y se cierra (`:40-53`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CorreoUsuario` | texto deshabilitado | — | Correo del usuario conectado | `:86`, `.form.txt:1507, 1521` |
| `Nombre` | texto deshabilitado | — | Nombre del usuario conectado | `:87`, `.form.txt:1071, 1085` |
| `PassNueva` | texto (`InputMask = Password`) | — | Nueva contraseña | `.form.txt:1139-1140` |
| `PassNueva1` | texto (`InputMask = Password`) | regla de validación | Repetición; `ValidationRule = [PassNueva]` con texto de error propio | `.form.txt:1193-1196` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida y guarda (ver Flujos) | `:7-63`, `.form.txt:1441-1443` |
| `ComandoSalir` | botón | `OnClick` | Cierra el formulario | `:64-67`, `.form.txt:1394-1396` |
| `ComandoAyudaFormulario` | botón | `OnClick` | `AbrirArchivoAyuda` | `:105-108`, `.form.txt:1334-1336` |

## Flujos

### Apertura (`Form_Open`, `:68-103`)

1. `Ajustar Me` y `DoCmd.Restore` (`:74, 80`).
2. Sin usuario conectado, error «No hay usuario activo» y cierre del formulario (`:82-85, 94-102`).
3. Muestra correo y nombre del usuario conectado (`:86-87`).

### Registrar (`ComandoRegistrar_Click`, `:7-63`)

1. Exige usuario conectado (`:17-20`).
2. Compara las dos contraseñas en binario; si difieren, foco en `PassNueva1` y «Las Passs no coinciden» (`:21-27`).
3. `Usuario.CambioPass` (`Usuario.cls:881-957`):
   - Formato: 8 a 16 caracteres con al menos un dígito, una minúscula y una mayúscula (`FG.bas:1177`); si no cumple, mensaje en `Usuario.cls:902`.
   - Repetición: busca el hash en `TbUsuariosHistoricoContrasenias` del usuario; si existe, «La Pass ha sido ya usada las últimas 10 veces» (`Usuario.cls:906-913, 971-981`).
   - Guarda el hash en `Password`, fija `FechaProximoCambioContrasenia` a hoy + 3 meses o `Null` si la contraseña nunca caduca, pone `TieneQueCambiarLaContrasenia = False` y vacía `PassIncialPlana` (`Usuario.cls:923-933`).
   - Registra el hash en el histórico con `HistoricoPass.Save` (`Usuario.cls:937-946`).
4. Recarga el usuario por `ID` (`:32`), avisa «Se ha cambiado la Pass satisfactoriamente» (`:39`) y vuelve al acceso con `Pass` vacía (`:40-53`).

### Errores

Errores internos (1000) muestran `m_Error`; el resto se envía con `CorreoAlAdministrador` y se muestra como `MsgBox` crítico (`:55-62`). Si falla el histórico tras guardar, el mensaje es «La Pass se guardó pero no quedó registrada en el histórico» (`Usuario.cls:943-945`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Las dos contraseñas coinciden | formulario y código | `.form.txt:1194-1195`; `:23-27` |
| Formato 8-16 con dígito, minúscula y mayúscula | `PassFormatoValido` | `FG.bas:1167-1183` |
| No repetir una contraseña del histórico | `PassRepetida` | `Usuario.cls:958-993` |
| Caducidad a 3 meses salvo `PasswordNuncaCaduca` | `CambioPass` | `Usuario.cls:925-929` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Histórico de contraseñas | `TbUsuariosHistoricoContrasenias` | SQL dinámico | `Usuario.cls:971-975` |
| Nueva contraseña | `TbUsuariosAplicaciones` | DAO `Edit/Update` | `Usuario.cls:914-933` |
| Alta en histórico | `TbUsuariosHistoricoContrasenias` | `HistoricoPass.Save` | `Usuario.cls:937-942` |
| Recarga del usuario | `TbUsuariosAplicaciones` | SQL dinámico | `:32`; `Constructor.bas:46-73` |

Sin transacción entre la actualización del usuario y el histórico (`Usuario.cls:933, 942`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: correo concatenado en los `WHERE` (`Usuario.cls:916, 973`); dos escrituras sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- No pide la contraseña actual: con la sesión abierta basta la nueva (`:17-28`).
- El mensaje habla de «las últimas 10 veces», pero `PassRepetida` busca en todo el histórico (`Usuario.cls:911, 971-974`).
- La expresión regular usa `\w*` en las comprobaciones previas (`FG.bas:1177`): una contraseña con símbolo antes del dígito o de las letras podría rechazarse.
- `Forms("FormLogin")` (`:52`) fallaría si `FormLogin` se cierra solo al abrir por apertura sin contraseña (`Form_FormLogin.cls:218-228`).
- No cierra `FormMenuPrincipalCalidad` (`:40-49`).
- Si falla el histórico, la contraseña ya está guardada (`Usuario.cls:933, 943-945`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pedir la contraseña actual salvo en el primer acceso; limitar la comprobación de repetición a las N últimas; guardar usuario e histórico en una transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Dos contraseñas distintas muestran el error de coincidencia y no cambian nada.
2. Una contraseña fuera de formato muestra el mensaje de formato.
3. Una contraseña ya usada por el usuario se rechaza.
4. Un cambio correcto fija la próxima caducidad a tres meses y obliga a volver a iniciar sesión.

## Pendientes

- Contenido de `HistoricoPass.Save` (`classes/HistoricoPass.cls`).
- Uso de `FechaProximaPassCalculado` frente a la columna guardada (`Usuario.cls:131`).
