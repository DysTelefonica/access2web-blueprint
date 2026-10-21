# Form_FormObtenerContraseña

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormObtenerContraseña.cls` (98 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario para solicitar una contraseña de un solo uso a partir del correo del usuario. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormObtenerContraseña.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Genreales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| — | No hay `OpenForm` de este formulario en `src` (ver Pendientes) | búsqueda en `src` |

Acepta el correo por `OpenArgs` y lo copia a `CorreoUsuario` (`:77-80`). Al terminar con éxito abre `FormLogin` y se cierra (`:46-47`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CorreoUsuario` | texto | — | Correo del usuario | `.form.txt:1270` |
| `ComandoObtener` («Obtener») | botón | `OnClick` | Resetea la contraseña y la envía por correo | `:7-57`, `.form.txt:1237-1239` |
| `ComandoSalir` | botón | `OnClick` | Cierra el formulario | `:58-61`, `.form.txt:1208-1210` |
| `ComandoAyudaFormulario` | botón | sin `OnClick` | No hace nada | `.form.txt:1169-1170` |

## Flujos

### Apertura (`Form_Open`, `:62-94`)

1. `DoCmd.Restore` y `Application.Echo False` (`:74-75`).
2. Si llega un correo por `OpenArgs`, lo escribe en `CorreoUsuario` (`:77-80`).

### Obtener (`ComandoObtener_Click`, `:7-57`)

1. Llama a `FormatoCorreoCorrecto` (`:20`), que exige un dominio corporativo fijo y un punto en la parte local (`FG.bas:1049-1062`); solo se comprueba `m_Error` (`:22-24`).
2. Carga el usuario por correo con `Constructor.getUsuario` (`:25`); si no existe, «No hay registrado ningún usuario con ese correo. Hable con su administrador» (`:29-32`).
3. `Usuario.ResetearPass` genera una contraseña aleatoria de 8 caracteres, guarda su hash en `Password`, marca `TieneQueCambiarLaContrasenia` y guarda la contraseña en claro en `PassIncialPlana` (`Usuario.cls:740, 757-765`); la envía al correo del usuario (`Usuario.cls:770, 801-825`).
4. Aviso «En un periodo máximo de 10 minutos recibirá un correo electrónico…» (`:44`), abre `FormLogin` y se cierra (`:46-47`).

### Errores

Errores internos (1000) muestran `m_Error`; el resto se envía con `CorreoAlAdministrador` y se muestra como `MsgBox` crítico (`:49-56`). En `Form_Open` el error cierra el formulario (`:85-93`).

## Reglas, validaciones y cálculos

- **Usuario existente:** la búsqueda es por `CorreoUsuario` exacto en `TbUsuariosAplicaciones` (`Constructor.bas:61-73`).
- **Contraseña de un solo uso:** `GenerarPassValida(8)` (`Usuario.cls:740`; `FG.bas:1070`); en el siguiente login se compara en claro con `PassIncialPlana` (`FG.bas:1418-1421`).
- Mismo comportamiento que el botón «Obtener» de [Form_FormLogin](Form_FormLogin.md) (`Form_FormLogin.cls:8-53`), que no valida el formato del correo.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario por correo | `TbUsuariosAplicaciones` | SQL dinámico concatenado | `Constructor.bas:46-73` |
| Reseteo de contraseña | `TbUsuariosAplicaciones` (`Password`, `TieneQueCambiarLaContrasenia`, `PassIncialPlana`) | DAO `Edit/Update` | `Usuario.cls:748-765` |
| Envío del correo | clase `Correo` | — | `Usuario.cls:820-825` |

Sin transacciones: si falla el envío, la contraseña ya está cambiada (`Usuario.cls:765, 770-773`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: correo concatenado en el `WHERE` (`Usuario.cls:750`; `Constructor.bas:70`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El resultado de `FormatoCorreoCorrecto` se descarta (`:20`): un correo con formato incorrecto no se rechaza.
- `Application.Echo False` (`:75`) no se restaura en este formulario.
- `ComandoAyudaFormulario` no tiene evento (`.form.txt:1169-1170`).
- La contraseña de un solo uso se guarda en claro (`Usuario.cls:763`).
- Si el envío falla, el usuario queda con una contraseña que no conoce (`Usuario.cls:765, 770`).
- Posible formulario huérfano: nadie lo abre en `src`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: enlace de restablecimiento con caducidad; misma respuesta exista o no el correo, para no revelar cuentas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un correo con formato no válido se rechaza antes de consultar la base.
2. Un correo no registrado muestra el aviso de usuario no registrado.
3. Un correo registrado recibe la contraseña de un solo uso y el usuario vuelve a la pantalla de acceso.
4. En el siguiente acceso con esa contraseña se exige cambiarla.

## Pendientes

- Quién abre el formulario: no hay `OpenForm` en `src`; puede abrirse desde el panel de navegación o estar en desuso.
- Contenido de `GenerarPassValida` (`FG.bas:1070`).
