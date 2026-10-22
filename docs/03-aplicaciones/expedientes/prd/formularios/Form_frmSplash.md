# Form_frmSplash

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_frmSplash.cls` (63 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de arranque de la aplicación: oculta la cinta, carga el entorno y el usuario conectado con `EVE` y abre el menú principal según el rol. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no escribe datos por sí mismo. Las citas son `Form_frmSplash.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Arranque de Access (presumible formulario de inicio) | Ningún `OpenForm "frmSplash"` en `src` | ver Pendientes |

| Abre | Condición | Evidencia |
|---|---|---|
| `Form0BDOpciones` | `EsTecnico <> Sí`, o cualquier error en el arranque | `:44-46, 61` |
| `Form0BDOpcionesTecnicos` | `EsTecnico = Sí` | `:47-49` |

Se cierra a sí mismo tras abrir el menú (`:53, 62`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Oculta la cinta salvo con `/dev` en la línea de comandos, ajusta tamaño, reinicia el contador de progreso | `:9-24`, `.form.txt:42` |
| (formulario) | — | `OnTimer` → `Form_Timer` (`TimerInterval = 50`) | Ejecuta la carga una sola vez y abre el menú | `:26-63`, `.form.txt:29, 377` |
| `lblEstado` | etiqueta | — | Texto de estado; lo actualizan el formulario y `Avance` | `:35, 37`; `FUNCIONES UTILES.bas:2735-2736`; `.form.txt:1016` |
| `lblProgresoFondo`, `lblProgresoBarra` | etiquetas | — | Barra de progreso: ancho de la barra proporcional al fondo | `:19`; `FUNCIONES UTILES.bas:2740-2764`; `.form.txt:1256, 1274` |
| `imgLogo`, `imgLogoDepartamento`, `lblTituloPrincipal`, `lblSubtitulo` | imagen y etiquetas | — | Decorativos | `.form.txt:773, 1033, 1210, 1233` |

El formulario no tiene botón de cierre ni cuadro de control (`.form.txt:10, 13`).

## Flujos

### Apertura (`Form_Open`, `:9-24`)

1. Si `Command()` no contiene `/dev`, `GestionarRibbon(False)` oculta la cinta (`:13-15`; `FUNCIONES UTILES.bas:6843-6852`).
2. `Ajustar Me` ajusta el alto y ancho interiores a las secciones (`:17`; `FUNCIONES UTILES.bas:2703-2727`).
3. `ResetAvanceCounter` pone a cero el contador de pasos y la barra a ancho 0 (`:18-19`; `FUNCIONES UTILES.bas:6838-6842`).

### Carga (`Form_Timer`, `:26-55`)

1. Desactiva el temporizador (`:31`) y muestra los textos de estado (`:34-37`).
2. `EVE` vacía las `TempVars`, lee la configuración local, identifica al usuario (por `Command` si viene informado o por el usuario de red), crea `Entorno`, calcula el rol y recorre las propiedades del entorno llamando a `Avance` en cada una (`:38`; `Variables Globales.bas:307-325, 348-375, 381-439`).
3. Si `EVE` devuelve error, lo eleva con número 1000 (`:39`).
4. Abre el menú según `EsTecnico` (`:44-50`), que `EVE` fija a `Sí` solo si el usuario no es administrador ni de calidad (`Variables Globales.bas:383-406`).

### Errores

Muestra «Ha habido un error en el inicio:» con la descripción y `m_Error` (`:58-59`), abre igualmente `Form0BDOpciones` y cierra el splash (`:61-62`). En `Form_Open` solo muestra `Err.Description` (`:22-23`).

## Reglas, validaciones y cálculos

- **Progreso:** `Avance` incrementa un contador por paso y fija el ancho de la barra como `pasos / ColItems.Count × ancho del fondo`, con tope en el ancho del fondo (`FUNCIONES UTILES.bas:2745-2763`).
- **Rol:** administrador y calidad excluyen técnico; técnico es el valor residual (`Variables Globales.bas:383-406`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario conectado | vía `constructor.getUsuario` | no trazado aquí | `Variables Globales.bas:355, 364` |
| Propiedades del entorno | vía `Entorno.getPropiedad` | no trazado aquí | `Variables Globales.bas:414, 426` |

Sin accesos propios: todo pasa por `EVE`.

## Coexistencia

Lo relevante está en `EVE` y `Entorno`, no en este formulario. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Ante cualquier error de `EVE`, incluido no identificar al usuario (`Variables Globales.bas:372-375`), abre igualmente `Form0BDOpciones` (`:61`) con el entorno a medio cargar.
- `/dev` en la línea de comandos solo mantiene la cinta (`:13`), pero `EVE` usa el mismo `Command` como identificador de usuario si no está vacío (`Variables Globales.bas:348-358`): el arranque con `/dev` podría fallar al buscar el usuario.
- El mensaje de error repite el texto: `Err.Raise 1000, , m_Error` pone `m_Error` como descripción y luego se concatenan ambos (`:39, 58-59`).
- `ProgresoActual` es público y no se usa en `src` (`:7`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir el splash por la carga de sesión y rol en el servidor, con página de error explícita si el usuario no se identifica.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario técnico entra al menú de técnicos; administradores y calidad, al menú principal.
2. Si el usuario no se puede identificar, la aplicación no muestra el menú y explica el motivo.

## Pendientes

- Confirmar que `frmSplash` es el formulario de inicio (propiedades de la base o macro `AutoExec`, no exportadas en `src`).
- Contenido de `LeeConfiguracionLocal`, `constructor.getUsuario` y `Entorno.getPropiedad`.
