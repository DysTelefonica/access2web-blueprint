# Form_FormMenuPrincipalCalidad

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormMenuPrincipalCalidad.cls` (375 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal del usuario de Calidad: misma cabecera que el del administrador y barra lateral con cuatro menús (Aplicaciones, Videos-Guías, Perfil, Configuración); Configuración solo contiene la gestión de Videos-Guías. Formulario **no enlazado** (sin `RecordSource` en el `.form.txt`), sin acceso propio a datos. Las citas `.cls:línea` y `.form.txt:línea` son de este formulario salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormLogin` (inicio de sesión y entrada automática) | `LoginCorrecto` lo elige si el usuario **no** es administrador y `EsUsuarioCalidadCalculado = Sí`; cierra el login y lo abre | `Form_FormLogin.cls:142, 219`; `Variables Globales.bas:724-736` |

Rol: `EsUsuarioCalidadCalculado` se toma de los permisos del usuario (`Usuario.cls:544-565`), que traducen el campo `EsUsuarioCalidad` («Sí»/«No») (`UsuarioAplicacionPermisos.cls:118-129`). Sin permisos cargados, el usuario no es de Calidad (`Usuario.cls:552-558`). El formulario no vuelve a comprobar el rol (`.cls:6-118`).

Salidas: `ImagenLogout` cierra y abre `FormLogin` (`.cls:129-140`); `ImagenSalir` cierra Access (`.cls:152-159`). `LoginCorrecto` lo cierra si ya estaba abierto (`Variables Globales.bas:732-734`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1`, `Imagen1`, `IND1`, `lblMenu1` («Aplicaciones») | rectángulo, imagen, etiquetas | `OnClick` → `MENU1_Click` | Carga `FormAplicaciones` y, si hay última aplicación abierta, activa su pestaña | `.cls:162-233`, `.form.txt:434-503` |
| `MENU2`, `Imagen2`, `IND2`, `lblMenu2` («Videos-Guías») | ídem | `OnClick` → `MENU2_Click` | Solo muestra «En construcción»; la carga de `FormVideosUsuarios` está comentada | `.cls:234-249`, `.form.txt:525-593` |
| `MENU3`, `IND3`, `lblMenu3` («Perfil»), `Imagen3` | ídem | solo `OnMouseMove` en `MENU3` | Sin handler de clic en el `.cls` ni `OnClick` en el `.form.txt` | `.cls:312-315`, `.form.txt:615-650, 1425` |
| `MENU4`, `Imagen4`, `IND4`, `lblMenu4` («Configuración») | ídem | `OnClick` → `MENU4_Click` | Despliega o pliega los submenús según `SUBMENU1.Visible` | `.cls:254-269, 285-299`, `.form.txt:1476-1540` |
| `SUBMENU1`, `lblSubMenu1` («Videos-Guías», `Tag = MENU4`) | rectángulo, etiqueta | `OnClick` → `SUBMENU1_Click` | Vacía `m_ObjUsuarioActivo` y carga `FormVideosGestion` | `.cls:271-281`, `.form.txt:1561-1586` |
| `MENU1`-`MENU4`, `SUBMENU1` | — | `OnMouseMove` | Resaltado bajo el ratón (`Menu.ActDesActMenu`, `Menu.ActivarSubMenu`) | `.cls:302-324`; `Menu.cls:75-171` |
| `PANEL1` | rectángulo | `OnMouseMove` | `RatonSobrePANEL1` | `.cls:326-330`; `Menu.cls:173-194` |
| sección `Detalle` | sección | `OnMouseMove` | Sin subformulario: `RatonSobreDetalle`; con `FormAplicaciones`: deja activo solo el menú 1 | `.cls:331-363` |
| `ImagenLogout` | imagen | `OnClick` | Confirma «¿Desea cerrar sesión?», marca `HechoLogout`, cierra y abre `FormLogin` | `.cls:129-140`, `.form.txt:354-355` |
| `ImagenSalir` | imagen | `OnClick` | Confirma y ejecuta `DoCmd.Quit acQuitSaveAll` | `.cls:152-159`, `.form.txt:320-321` |
| `lblUsuarioConectado`, `lblUsuarioUbicacion`, `lblVer` | etiquetas | — | Rellenadas en `Form_Load` | `.cls:29-39` |
| `FormDetalle` | subformulario | — | Contenido fijado por código | `.form.txt:1405-1411` |

Los rótulos de diseño de `lblUsuarioConectado` y del título del formulario contienen un nombre de persona real (`.form.txt:36, 280`); se sobrescriben en `Form_Load` (`.cls:30-33`).

## Flujos

### Apertura (`Form_Load`, `.cls:6-118`)

1. Reloj de arena y `Ajustar Me` (`.cls:17-19`; `Funciones Genreales.bas:89-116`).
2. `getMenu` construye el objeto `Menu` y oculta el submenú (`.cls:22-25`; `Menu.cls:216-265`).
3. Rotula título con `getNombreUsuarioConectado`, versión, usuario y equipo (aquí con `m_ObjUsuarioConectadoLogin.Nombre`) y ubicación según `m_EnOficina` (`.cls:26-39`).
4. Activa el menú 1, ejecuta `MENU1_Click`, fuerza la pestaña 1 de `FormAplicaciones` y carga `FormAplicacionesFueraOficina` (`.cls:40-64`). El bloque alternativo por última aplicación está comentado (`.cls:65-114`).

### Menú Aplicaciones (`MENU1_Click`, `.cls:162-215`)

Igual que en el menú de administrador: carga `FormAplicaciones` (`:167-169`) y, con última aplicación abierta, activa la pestaña «1» o «2» de `EtiquetaAplicacionesAPulsar` (`:170-212`; `Usuario.cls:650-670`).

### Errores

`Form_Load`, `MENU1_Click` e `ImagenLogout_Click` envían `CorreoAlAdministrador` salvo el error 1000 y muestran `MsgBox` crítico (`.cls:119-126, 216-223, 143-149`).

## Reglas, validaciones y cálculos

- Acceso por rol: administrador tiene prioridad; Calidad solo si no es administrador (`Variables Globales.bas:724-727`).
- Configuración pliega y despliega su único submenú y vacía `FormDetalle` en ambos casos (`.cls:254-260, 285-299`).
- Frente al menú de administrador no hay gestión de usuarios ni de aplicaciones (`.cls:254-281`).

## Lecturas y escrituras

Sin acceso propio a datos. Lectura indirecta de la última aplicación abierta (`Usuario.cls:490-503`; `Constructor.bas:1095-1100`).

## Coexistencia

Pendiente de `../coexistencia.md`. Mismo punto que el menú de administrador: SQL concatenado con el correo en `Constructor.bas:1100`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El menú «Perfil» no hace nada al pulsarlo: `MENU3`, `IND3`, `lblMenu3` e `Imagen3` no tienen `OnClick` (`.form.txt:615-650, 1425`) ni existe `MENU3_Click` en el `.cls`; en el menú de usuario sí abre `FormPerfil`.
- `Form_FormUsuarioCambioPass` cierra los menús de administrador y de usuario tras cambiar la contraseña, pero no este (`Form_FormUsuarioCambioPass.cls:44-49`): quedaría abierto junto al nuevo login.
- `Form_Load` deshace la pestaña de última aplicación que `MENU1_Click` acaba de activar (`.cls:44-64` frente a `:192-209`).
- El mensaje de error de `Form_Load` dice «Al Form_Activate» (`.cls:121`).
- `ActivarObjetoMenu` comprueba `m_Error` sin habérselo pasado a `getMenu` (`.cls:368-371`).
- `ActDesActMenu` con `p_EsActivar = No` usa `i = 0` en vez de `p_Numero` (`Menu.cls:83, 89`); `Detalle_MouseMove` lo compensa a mano solo con `FormAplicaciones` (`.cls:337-362`).
- `ImagenSalir` usa `acQuitSaveAll` (`.cls:159`), al contrario que el menú de administrador.
- `HechoLogout` se asigna (`.cls:136`) y nunca se lee.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir si Calidad tiene «Perfil» o retirarlo; menú único con secciones por rol autorizadas en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario de Calidad no administrador entra a este menú; un administrador entra al suyo aunque también sea de Calidad.
2. «Aplicaciones» abre el catálogo en la pestaña «fuera de la oficina».
3. «Configuración» muestra solo la gestión de Videos-Guías.
4. Tras cambiar la contraseña no queda abierto ningún menú principal.
5. «Cerrar sesión» y «Salir» piden confirmación.

## Pendientes

- Si «Perfil» sin acción es intencionado (no hay comentario ni código comentado que lo explique).
- Diferencia de efecto entre `acQuitSaveAll` y `acQuitSaveNone` en un formulario no enlazado.
- Cálculo de `EnOficina` (`Funciones Genreales.bas:1003`).
- Comportamiento de `FormAplicaciones` y `FormVideosGestion` (sus fichas).
