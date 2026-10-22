# Form_FormMenuPrincipalUsuario

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormMenuPrincipalUsuario.cls` (373 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal del usuario sin rol de administración ni de Calidad: cabecera con usuario, equipo, ubicación y versión; tres menús (Aplicaciones, Videos-Guías, Perfil) sin submenús; subformulario `FormDetalle`. Formulario **no enlazado** (sin `RecordSource` en el `.form.txt`), sin acceso propio a datos. Las citas `.cls:línea` y `.form.txt:línea` son de este formulario salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormLogin` (inicio de sesión y entrada automática) | `LoginCorrecto` lo abre cuando el usuario no es administrador ni de Calidad | `Form_FormLogin.cls:142, 219`; `Variables Globales.bas:724-736` |

Es el formulario por defecto: cualquier usuario que no cumpla `EsAdministradorCalculado` ni `EsUsuarioCalidadCalculado` llega aquí (`Variables Globales.bas:728-729`). El formulario no comprueba el rol (`.cls:6-105`).

Salidas: `ImagenLogout` cierra y abre `FormLogin` (`.cls:116-127`); `ImagenSalir` cierra Access (`.cls:139-146`). Lo cierran `LoginCorrecto` si ya estaba abierto (`Variables Globales.bas:732-734`) y `Form_FormUsuarioCambioPass` tras cambiar la contraseña (`Form_FormUsuarioCambioPass.cls:47-49`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1`, `Imagen1`, `IND1`, `lblMenu1` («Aplicaciones») | rectángulo, imagen, etiquetas | `OnClick` → `MENU1_Click` | Carga `FormAplicaciones`; la activación por última aplicación está comentada | `.cls:149-220`, `.form.txt:434-503` |
| `MENU2`, `Imagen2`, `IND2`, `lblMenu2` («Videos-Guías») | ídem | `OnClick` → `MENU2_Click` | Solo muestra «En construcción»; la carga de `FormVideosUsuarios` está comentada | `.cls:221-245`, `.form.txt:525-593` |
| `MENU3`, `Imagen3`, `IND3`, `lblMenu3` («Perfil») | ídem | `OnClick` → `MENU3_Click` | Carga `FormPerfil` en `FormDetalle` | `.cls:246-270`, `.form.txt:615-654, 1428-1429` |
| `MENU1`-`MENU3` | — | `OnMouseMove` | Resaltado bajo el ratón (`Menu.ActDesActMenu`) | `.cls:295-318`; `Menu.cls:75-131` |
| `PANEL1` | rectángulo | `OnMouseMove` | `RatonSobrePANEL1` | `.cls:321-328`; `Menu.cls:173-194` |
| sección `Detalle` | sección | `OnMouseMove` | Sin subformulario: `RatonSobreDetalle`; con `FormAplicaciones`: deja activo solo el menú 1 | `.cls:330-362`, `.form.txt:400-401` |
| `ImagenLogout` | imagen | `OnClick` | Confirma «¿Desea cerrar sesión?», marca `HechoLogout`, cierra y abre `FormLogin` | `.cls:116-127`, `.form.txt:354-355` |
| `ImagenSalir` | imagen | `OnClick` | Confirma y ejecuta `DoCmd.Quit acQuitSaveAll` | `.cls:139-146`, `.form.txt:320-321` |
| `lblUsuarioConectado`, `lblUsuarioUbicacion`, `lblVer` | etiquetas | — | Rellenadas en `Form_Load` | `.cls:28-38` |
| `FormDetalle` | subformulario | — | Contenido fijado por código | `.form.txt:1414` |

Los rótulos de diseño del título y de `lblUsuarioConectado` contienen un nombre de persona real y un equipo que no es el de este rol (`.form.txt:36, 280`); se sobrescriben en `Form_Load` (`.cls:29-32`).

## Flujos

### Apertura (`Form_Load`, `.cls:6-105`)

1. Reloj de arena, `Ajustar Me` y `getMenu` (`.cls:18-24`; `Funciones Genreales.bas:89-116`; `Menu.cls:216-265`).
2. Rotula título, versión, usuario y equipo, y ubicación según `m_EnOficina` (`.cls:26-38`).
3. Activa el menú 1, ejecuta `MENU1_Click`, fuerza la pestaña 1 de `FormAplicaciones` y carga `FormAplicacionesFueraOficina` (`.cls:39-59`; destino de la pestaña en `Form_FormAplicaciones.cls:39-46`).

### Navegación

Cada clic apaga el repintado (`Application.Echo False`), cambia `FormDetalle` y lo vuelve a encender (p. ej. `.cls:247-252`). Todos los handlers de menú y de ratón usan `On Error Resume Next` (`.cls:152, 204, 247, 296, 331, 366`).

### Errores

Solo `Form_Load` e `ImagenLogout_Click` envían `CorreoAlAdministrador` y muestran `MsgBox` crítico (`.cls:106-113, 130-136`). La etiqueta `errores` de `MENU1_Click` (`.cls:194-201`) es inalcanzable por el `On Error Resume Next` de `.cls:152`.

## Reglas, validaciones y cálculos

- Rol por defecto: ni administrador ni Calidad (`Variables Globales.bas:724-729`).
- No hay submenús en el `.form.txt` (ningún control `SUBMENU*`); `Desplegar` y `Plegar` existen (`.cls:272-292`) pero ningún handler los llama.
- A diferencia de los menús de administrador y Calidad, «Aplicaciones» no reabre la pestaña de la última aplicación (`.cls:157-191`, comentado).

## Lecturas y escrituras

Sin acceso propio a datos. `getNombreUsuarioConectado` devuelve el nombre del usuario del login o, en su defecto, el del usuario por máquina o «Desconocido» (`Variables Globales.bas:224-251`).

## Coexistencia

Pendiente de `../coexistencia.md`. Sin SQL propio.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `On Error Resume Next` en `ActivarObjetoMenu` (`.cls:366`) y en todos los handlers de navegación oculta cualquier fallo: si `getMenu` falla, los `m_ObjMenu.*` posteriores fallan en silencio.
- `IND3_Click` no restaura `Application.Echo True` tras `MENU3_Click` (`.cls:254-258`), al contrario que `imagen3_Click` y `lblMenu3_Click`; lo restaura `MENU3_Click` (`.cls:252`) salvo que falle antes.
- El mensaje de error de `Form_Load` dice «Al Form_Activate» (`.cls:108`).
- `ActDesActMenu` con `p_EsActivar = No` usa `i = 0` en vez de `p_Numero` (`Menu.cls:83, 89`): `RatonSobreDetalle` y `RatonSobrePANEL1` no desactivan menús.
- `HechoLogout` se asigna (`.cls:123`) y nunca se lee.
- Código muerto: `Desplegar`, `Plegar` y las variables `ctlIndicador`, `ctlMenu`, `m_NumeroEtiquetaUltima` (`.cls:272-292, 8-10`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: menú único con secciones por rol; tratar los errores de navegación de forma visible en lugar de ignorarlos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario sin rol de administrador ni de Calidad entra a este menú.
2. La cabecera muestra nombre, equipo, ubicación y versión.
3. «Aplicaciones» abre el catálogo en la pestaña «fuera de la oficina».
4. «Perfil» abre el perfil del usuario conectado.
5. «Videos-Guías» informa de que está en construcción.
6. «Cerrar sesión» y «Salir» piden confirmación.

## Pendientes

- Contenido de `FormPerfil` y `FormAplicaciones` (sus fichas).
- Cálculo de `EnOficina` (`Funciones Genreales.bas:1003`).
- Si el rótulo de equipo de diseño (`.form.txt:280`) indica que este formulario se copió del de Calidad.
