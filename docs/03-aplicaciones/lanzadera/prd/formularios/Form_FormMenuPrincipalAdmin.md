# Form_FormMenuPrincipalAdmin

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormMenuPrincipalAdmin.cls` (323 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal del administrador: cabecera con usuario, equipo, ubicación y versión; barra lateral con tres menús (Aplicaciones, Videos-Guías, Configuración) y un subformulario `FormDetalle` donde se carga cada sección. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe en la base de datos por sí mismo. Las citas `.cls:línea` y `.form.txt:línea` son de este formulario salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormLogin` (inicio de sesión con contraseña y entrada automática en `Form_Load`) | `LoginCorrecto` elige este formulario si `EsAdministradorCalculado = Sí`, cierra el login y lo abre | `Form_FormLogin.cls:142, 219`; `Variables Globales.bas:724-736` |

Rol: `EsAdministradorCalculado` traduce el campo `EsAdministrador` («Sí»/«No») del usuario (`Usuario.cls:67-77`). El formulario **no comprueba el rol** al abrirse (`.cls:15-75`): la única barrera es la elección de `LoginCorrecto`.

Salidas: `ImagenLogout` cierra el formulario y abre `FormLogin` (`.cls:86-97`); `ImagenSalir` cierra Access (`.cls:109-117`). Lo cierran también `LoginCorrecto` si ya estaba abierto (`Variables Globales.bas:732-734`) y `Form_FormUsuarioCambioPass` tras cambiar la contraseña (`Form_FormUsuarioCambioPass.cls:44-46`). `ActualizarAperturas` lo localiza por nombre para refrescar contadores en el subformulario de aplicaciones (`Funciones Genreales.bas:3450-3455`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1`, `Imagen1`, `IND1`, `lblMenu1` («Aplicaciones») | rectángulo, imagen, etiquetas | `OnClick` → `MENU1_Click` | Carga `FormAplicaciones` y, si hay última aplicación abierta, activa su pestaña | `.cls:124-186`, `.form.txt:751-816` |
| `MENU2`, `Imagen2`, `IND2`, `lblMenu2` («Videos-Guías») | ídem | `OnClick` → `MENU2_Click` | Solo muestra «En construcción»; la carga de `FormVideosUsuarios` está comentada | `.cls:187-202`, `.form.txt:837-902` |
| `MENU3`, `Imagen3`, `IND3`, `lblMenu3` («Configuración») | ídem | `OnClick` → `MENU3_Click` | Despliega o pliega los submenús según `SUBMENU1.Visible` | `.cls:204-219, 7-13, 261-267`, `.form.txt:923-988` |
| `SUBMENU1`, `lblSubMenu1` («Usuario») | rectángulo, etiqueta (`Tag = MENU3`) | `OnClick` → `SUBMENU1_Click` | Vacía `m_ObjUsuarioActivo`, carga `FormUsuarioGeneral` y llama a su `MENU1_Click` | `.cls:221-240`, `.form.txt:1010-1036` |
| `SUBMENU2`, `lblSubMenu2` («Aplicaciones») | ídem | `OnClick` → `SUBMENU2_Click` | Vacía `m_ObjUsuarioActivo`, carga `FormAplicacionGeneral` y llama a su `MENU1_Click` | `.cls:241-260`, `.form.txt:1058-1084` |
| `SUBMENU3`, `lblSubMenu3` («Videos-Guías») | ídem | `OnClick` → `SUBMENU3_Click` | Vacía `m_ObjUsuarioActivo` y carga `FormVideosGestion` | `.cls:268-275, 120-122`, `.form.txt:1106-1132` |
| `MENU1`-`MENU3`, `SUBMENU1`-`SUBMENU3` | — | `OnMouseMove` | Resaltado del menú o submenú bajo el ratón (`Menu.ActDesActMenu`, `Menu.ActivarSubMenu`) | `.cls:280-304`; `Menu.cls:75-171` |
| `PANEL1`, sección `Detalle` | rectángulo, sección | `OnMouseMove` | `RatonSobrePANEL1` / `RatonSobreDetalle` | `.cls:305-313`; `Menu.cls:173-215` |
| `ImagenLogout` | imagen | `OnClick` | Confirma «¿Desea cerrar sesión?», marca `HechoLogout`, cierra y abre `FormLogin` | `.cls:86-97`, `.form.txt:668-669` |
| `ImagenSalir` | imagen | `OnClick` | Confirma y ejecuta `DoCmd.Quit acQuitSaveNone` | `.cls:109-117`, `.form.txt:632-633` |
| `lblUsuarioConectado`, `lblUsuarioUbicacion`, `lblVer` | etiquetas | — | Rellenadas en `Form_Load` | `.cls:37-47` |
| `FormDetalle` | subformulario | — | Sin `SourceObject` de diseño; lo fija el código | `.form.txt:1886-1900` |

## Flujos

### Apertura (`Form_Load`, `.cls:15-75`)

1. Reloj de arena y `Ajustar Me` (ajuste de tamaño de ventana, `Funciones Genreales.bas:89-116`) (`.cls:25-27`).
2. `getMenu` crea el objeto `Menu`: cuenta etiquetas `lblMENU*` y `lblSUBMENU*`, toma el menú desplegable del `Tag` del primer submenú, pinta colores y oculta submenús (`.cls:29-32`; `Constructor.bas:1062-1082`; `Menu.cls:216-265`).
3. Rotula título, versión (`m_ObjEntorno.VersionAplicacion`), usuario y equipo, y ubicación «En Oficina» o «Fuera de la Oficina» según `m_EnOficina` (`.cls:34-47`). `m_EnOficina` se calcula en `EVE` (`Variables Globales.bas:598`).
4. Activa el menú 1 y ejecuta `MENU1_Click` (`.cls:48-52`).
5. Sobre `FormAplicaciones` ya cargado, fuerza la pestaña 1 (indicadores y colores) y carga `FormAplicacionesFueraOficina` en su subformulario (`.cls:53-70`).

### Menú Aplicaciones (`MENU1_Click`, `.cls:124-168`)

Carga `FormAplicaciones` (`:129-131`). Si el usuario tiene última aplicación abierta, `EtiquetaAplicacionesAPulsar` devuelve «1» (fuera de oficina) o «2» (en oficina) (`Usuario.cls:650-670`) y el handler activa esa pestaña con `Form_FormAplicaciones.MENU1_Click` o `MENU2_Click` (`:132-165`), que cargan `FormAplicacionesFueraOficina` u `FormAplicacionesOficina` (`Form_FormAplicaciones.cls:39-58`).

### Errores

`Form_Load`, `MENU1_Click` e `ImagenLogout_Click` envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`.cls:76-83, 169-176, 100-106`). `ImagenSalir_Click` usa `On Error Resume Next` (`.cls:110`).

## Reglas, validaciones y cálculos

- Acceso por rol: solo `LoginCorrecto` lo decide; administrador tiene prioridad sobre Calidad (`Variables Globales.bas:724-729`).
- Submenús de Configuración: visibles solo tras pulsar `MENU3`; pulsar de nuevo los oculta; en ambos casos se vacía `FormDetalle` (`.cls:204-210, 7-13, 261-267`).
- Antes de cargar Usuario, Aplicaciones o Videos-Guías de configuración se vacía la global `m_ObjUsuarioActivo` (`.cls:223, 243, 270`).

## Lecturas y escrituras

Sin acceso propio a datos. Lecturas indirectas: `UltimaAplicacionAbierta` consulta `TbConexionesRegistro`, `TbConexionUltimaAppAbierta` y `TbAplicaciones` con SQL concatenado (`Usuario.cls:490-503`; `Constructor.bas:1095-1100`). El correo de error lo envía `CorreoAlAdministrador` (`Funciones Genreales.bas:1447`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la consulta de última aplicación concatena el correo del usuario en el SQL (`Constructor.bas:1100`); en PostgreSQL debe parametrizarse.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Form_Load` deshace la lógica de última aplicación: tras `MENU1_Click` fuerza pestaña 1 y `FormAplicacionesFueraOficina` (`.cls:52-70`), aunque `MENU1_Click` hubiera activado la pestaña «en oficina» (`.cls:150-163`).
- `ActDesActMenu` con `p_EsActivar = No` usa la variable `i` sin inicializar (0) en vez de `p_Numero` (`Menu.cls:83, 89`): `RatonSobreDetalle` y `RatonSobrePANEL1` no desactivan ningún menú.
- `ActivarObjetoMenu` llama a `getMenu(Me)` sin pasar `m_Error` y luego comprueba `m_Error` (`.cls:317-319`): un fallo de `getMenu` no se detecta.
- `HechoLogout` se asigna (`.cls:93`) pero ningún módulo ni formulario lo lee: solo se declara (`Variables Globales.bas:127`) y se asigna en los tres menús principales.
- `lblSubMenu13_Click` (`.cls:276-278`) no corresponde a ningún control del `.form.txt`: código muerto.
- `SUBMENU2` y `SUBMENU3` vacían `m_ObjUsuarioActivo` aunque no tratan usuarios (`.cls:243, 270`).
- `ImagenSalir` sale con `acQuitSaveNone` (`.cls:117`); los menús de Calidad y Usuario usan `acQuitSaveAll`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: autorizar el menú de administración en el servidor por rol, no por el formulario elegido; recordar la última pestaña de aplicaciones de forma coherente; una sola plantilla de menú con secciones por rol en lugar de tres formularios.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario administrador entra al menú de administración; uno no administrador no puede acceder a él aunque conozca la ruta.
2. La cabecera muestra nombre, equipo, ubicación (en oficina / fuera) y versión.
3. «Aplicaciones» abre el catálogo en la pestaña «fuera de la oficina».
4. «Videos-Guías» informa de que está en construcción.
5. «Configuración» muestra Usuario, Aplicaciones y Videos-Guías, que abren la gestión correspondiente.
6. «Cerrar sesión» pide confirmación y vuelve al login; «Salir» pide confirmación y termina la sesión.

## Pendientes

- Si Access ejecuta el `On Error` de `ActivarObjetoMenu` cuando `m_ObjMenu` falla a mitad de `EstablecerDatosIniciales` (`Menu.cls:216-265` no tiene manejador propio).
- Cálculo de `EnOficina` (`Funciones Genreales.bas:1003`) y contenido de `CorreoAlAdministrador`.
- Comportamiento de `FormUsuarioGeneral`, `FormAplicacionGeneral`, `FormVideosGestion` y `FormAplicaciones` (sus fichas).
- Efecto real de `CloseButton` y `AutoCenter` del `.form.txt:8-10` en ejecución.
