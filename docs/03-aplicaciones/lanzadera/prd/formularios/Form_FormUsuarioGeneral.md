# Form_FormUsuarioGeneral

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuarioGeneral.cls` (181 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de la gestión de usuarios: cabecera con tres pestañas dibujadas a mano (rectángulos `MENU1`-`MENU3` con etiquetas e indicadores `IND1`-`IND3`), la etiqueta del usuario activo y el subformulario `FormDetalle` que aloja cada pestaña. No lee ni escribe datos: solo gestiona la navegación y la global `m_ObjUsuarioActivo` (`Variables Globales.bas:100`). Las citas son `Form_FormUsuarioGeneral.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalAdmin` | `SUBMENU1_Click` («Usuario») vacía el usuario activo, carga `FormUsuarioGeneral` en su `FormDetalle`, marca la pestaña 1 y llama a `Form_FormUsuarioGeneral.MENU1_Click` | `Form_FormMenuPrincipalAdmin.cls:221-237`, `Form_FormMenuPrincipalAdmin.form.txt:1032-1033` |

El menú de administración se abre tras el login solo si el usuario conectado es administrador (`Variables Globales.bas:724-725, 736`). Abre (como subformulario): `FormUsuariosGestion` (`:68`), `FormUsuarioDatosGenerales` (`:79`) y `FormUsuarioPerfilAplicaciones` (`:92`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1` / `lblMenu1` («Búsqueda») | rectángulo / etiqueta | `OnClick` | Carga `FormUsuariosGestion`, vacía el usuario activo y engancha sus eventos | `:66-76`, `.form.txt:244-286` |
| `MENU2` / `lblMenu2` («Datos Generales») | rectángulo / etiqueta | `OnClick` | Carga `FormUsuarioDatosGenerales`; no exige usuario activo | `:77-85`, `.form.txt:307-353` |
| `MENU3` / `lblMenu3` («Perfil Aplicaciones») | rectángulo / etiqueta | `OnClick` | Sin usuario activo: «No se ha establecido un usuario activo»; con él, carga `FormUsuarioPerfilAplicaciones` | `:86-98`, `.form.txt:461-507` |
| `MENU1`-`MENU3`, `lblMenu1`-`lblMenu3` | — | `OnMouseMove` | Resaltan la pestaña bajo el ratón (indicador visible y color de `m_ObjEntorno`) | `:100-179` |
| `EncabezadoDelFormulario` | sección | `OnMouseMove` | Oculta indicadores y apaga los colores | `:7-26`, `.form.txt:226-227` |
| `lblUsuarioActivo` | etiqueta | — | «Usuario Activo: » más el nombre recibido por evento | `:45, 59, 63`, `.form.txt:571-572` |
| `FormDetalle` | subformulario | — | Aloja la pestaña activa; vacío al abrir | `:43`, `.form.txt:599` |

## Flujos

### Apertura (`Form_Load`, `:28-56`)

1. Muestra `IND1`, oculta `IND2` y pone colores de pestaña desde `m_ObjEntorno` (`:34-40`).
2. Vacía `FormDetalle`, el usuario activo y la etiqueta (`:43-45`). El contenido lo carga después `MENU1_Click`, llamado desde el menú (`Form_FormMenuPrincipalAdmin.cls:236`).

### Navegación

1. `MENU1_Click` solo cambia `SourceObject` si es distinto (`:67-69`), vacía siempre el usuario activo (`:70`) y asigna `m_Form` a la instancia cargada (`:71`).
2. `m_Form_UsuarioSeleccionado` escribe el nombre en la etiqueta (`:62-64`); `m_Form_QuitarUsuarioSeleccionado` la deja en «Usuario Activo: » (`:58-60`).
3. `MENU2_Click` carga la ficha sin condición (`:77-81`): sin usuario activo, la ficha entra en modo alta (`Form_FormUsuarioDatosGenerales.cls:101-110`).
4. `MENU3_Click` exige usuario activo (`:87-90`).

### Errores

Solo `Form_Load` tiene manejador: `CorreoAlAdministrador` y `MsgBox` crítico (`:47-55`). Los handlers de menú no lo tienen.

## Reglas, validaciones y cálculos

- Perfil de aplicaciones requiere usuario activo; datos generales, no (`:77-94`).
- Volver a «Búsqueda» descarta el usuario activo (`:70`).
- Permisos: el formulario no comprueba rol; lo protege el menú que lo aloja (`Variables Globales.bas:724-725`).

## Lecturas y escrituras

Ninguna propia. Los colores proceden de `m_ObjEntorno` (`:17-25, 38-40`).

## Coexistencia

Sin acceso a datos: no requiere adaptación de SQL. Debe preservarse el estado de sesión «usuario activo» compartido entre pestañas (ver `../coexistencia.md`, tanda posterior).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `MENU1_Click` vacía `m_ObjUsuarioActivo` (`:70`) pero no la etiqueta `lblUsuarioActivo`: tras volver a Búsqueda la cabecera puede mostrar un usuario que ya no está activo, y `MENU3` responde «No se ha establecido un usuario activo» (`:87-89`).
- `m_Form` solo se asigna en `MENU1_Click` (`:71`); si la llamada desde el menú alcanza otra instancia (llamada por nombre de clase, `Form_FormMenuPrincipalAdmin.cls:236`), los eventos de selección no llegarían a esta etiqueta.
- `Form_Load` no inicializa `IND3` (`:34-36`).
- Entrar en «Datos Generales» sin usuario activo abre un alta silenciosa y deja en la global un usuario nuevo sin guardar (`Form_FormUsuarioDatosGenerales.cls:101-110`), con el que `MENU3` deja pasar a Perfil (`:87`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pestañas como rutas con el ID del usuario en la URL; acción explícita «Nuevo usuario» separada de «Datos generales».

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar, se muestra la pestaña Búsqueda y la cabecera dice «Usuario Activo: » sin nombre.
2. Seleccionar un usuario en Búsqueda muestra su nombre en la cabecera.
3. Perfil Aplicaciones sin usuario activo muestra «No se ha establecido un usuario activo» y no cambia de pestaña.
4. Volver a Búsqueda descarta el usuario activo.

## Pendientes

- Semántica de `Form_FormUsuarioGeneral.MENU1_Click` cuando el formulario está alojado como subformulario (`Form_FormMenuPrincipalAdmin.cls:236`).
- Valores de `ColorMenuActivado`/`ColorMenuDesactivado` en `Entorno` (no trazados).
