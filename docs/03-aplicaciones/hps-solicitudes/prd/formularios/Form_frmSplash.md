# Form_frmSplash

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_frmSplash.cls` (63 líneas) y `.form.txt` (15 controles `Begin`) · **PRD:** [README](../README.md)

Pantalla de arranque: oculta la cinta, inicializa el entorno con `EVE` y el cálculo de tareas pendientes del `Entorno`, y abre `Form0BDOpciones`. Es **no enlazada** (sin `RecordSource` efectivo) y dispone de una barra de progreso pintada por código. Las citas son `Form_frmSplash.cls:línea` salvo indicación; `VG.bas` = `modules/Variables Globales.bas`, `FU.bas` = `modules/FUNCIONES UTILES.bas`, `classes/Entorno.cls` y demás clases figuran como `classes/Nombre.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| (propiedad de arranque de la base) | `Form_frmSplash` figura como arranque; ningún `DoCmd.OpenForm "Form_frmSplash"` se observa en `src`. | (no en `src`) |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpciones` | `DoCmd.OpenForm "Form0BDOpciones"` tras el cálculo de pendientes; también se abre en el bloque `errores:` para no dejar la aplicación cerrada. | `:48, 56` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | Oculta la cinta salvo con `/dev` (`:11-17`), aplica `Ajustar Me` (`:21`), reinicia el contador de avance (`:24`) y vacía la barra (`:30`). | `:9-36`, `.form.txt:41` |
| Formulario | — | `OnTimer` (`TimerInterval = 50`, `.form.txt:36`) → `Form_Timer` | Arranque completo (ver §Flujos); detiene su propio temporizador (`:41`). | `:38-58`, `.form.txt:36` |
| `lblEstado` | etiqueta | — | Cambia su `Caption` a «Cargando configuración...» y luego «Analizando y calculando tareas pendientes...». | `:50, 60`, `.form.txt:1014-1024` |
| `lblProgresoFondo`, `lblProgresoBarra` | etiquetas | — | Tamaño y avance de la barra; la lógica de progreso vive en `FU.bas` (`Avance`, `ResetAvanceCounter`). | `:30`, `.form.txt:1281-1310` |
| `imgLogoDepartamento`, `imgLogo` | imágenes | — | Cabecera y logotipo del centro; estáticos. | `.form.txt:765-993, 1026-1224` |
| `lblTituloPrincipal`, `lblSubtitulo` | etiquetas | — | Rótulos fijos «HPS SOLICITUDES» y subtítulo. | `.form.txt:1235-1280` |

## Flujos

### Normal

1. `Form_Open`: si la línea de comandos no contiene `/dev`, `GestionarRibbon(False)` oculta la cinta (`:17`); `Ajustar Me` (`:21`); `ResetAvanceCounter` (`:24`); ancho de la barra a 0 (`:30`).
2. `Form_Timer`: detiene el temporizador (`:41`), fija `Me.Caption = "Iniciando HPS SOLICITUDES ..."` y `lblEstado.Caption = "Cargando configuración..."` (`:50-53`), `DoEvents` y llama a `EVE m_Error` (`:55`).
3. Si `EVE` devuelve error, `Err.Raise 1000, , m_Error` fuerza la rama de error (`:56-57`).
4. Cambia el rótulo a «Analizando y calculando tareas pendientes...» (`:60-63`) y, si `m_ObjEntorno` está vivo, invoca `m_ObjEntorno.CalcularTotalesPendientes m_ObjUsuarioConectado, EsTecnico` (`:67-69`). Esta llamada es la carga pesada tras `EVE`.
5. Abre `Form0BDOpciones` y se cierra (`:71-72`).

### Errores

- En `errores:` muestra «Ha habido un error en el inicio:» con `Err.Description` y `m_Error` (`:74-76`); como red de seguridad, **abre también `Form0BDOpciones` y se cierra**, sin distinguir la causa (`:78-79`). Esto difiere del patrón de Lanzadera (que aborta la sesión) y podría llevar a un usuario a trabajar con sesión incompleta.
- `Form_Open` tiene un handler `errores:` propio que muestra `Err.Description` (`:32-34`); aborta sin continuar.

## Reglas, validaciones y cálculos

- El arranque solo se distingue por la presencia de `/dev` en la línea de comandos (`Command()`); este conmutador también se pasa a `EVE` (que lo lee como pseudo-correo, ver `Form_frmSplash.cls:17` y la observación en §Posibles defectos).
- `Entorno.CalcularTotalesPendientes` (en `classes/Entorno.cls`) y la barra de progreso son los consumidores de tiempo del arranque: la invocación está sincronizada con `DoEvents` (`:55, 63`) para mantener la barra viva y no bloquear el repintado.
- El método de inicio de la base (propiedad «Form/Report de inicio») no aparece en `src`; para reproducir el arranque hay que abrir manualmente `Form_frmSplash` o el `.accdb` resuelve el formulario de inicio a otro nombre.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Identidad de usuario y sesión | `TbUsuariosAplicaciones` (Lanzadera) y permisos `TbUsuariosAplicacionesPermisos` (lectura indirecta desde `EVE`) | `QueryDef` parametrizada / lectura DAO | `VG.bas:103-118`, `classes/Usuario.cls:1-50`, `classes/UsuarioAplicacionPermisos.cls:1-30` |
| Configuración local y remota del entorno | Tabla `TbConfiguracion` (la del backend `Solicitudes_HPS_datos.accdb`), clave de backend y propiedades de `Entorno` | `QueryDef` / API de perfil (`Configuracion` y `ConfiguracionOperaciones`) | `VG.bas:128-158`, `classes/Configuracion.cls:1-100` |
| Cálculo de tareas pendientes | Variables del `Entorno` y `UsuarioConectado` (sin tabla fija, suma por rol y tipo) | DAO sobre `TbSolicitudes` (lectura) y `TbResponsables` desde `Entorno.CalcularTotalesPendientes` | `classes/Entorno.cls:1-3065` (función no acotada en `src`; a acotar en la matriz) |

Sin escrituras en arranque; el splash no crea filas.

## Coexistencia

`EVE` se apoya en `getdb()` con backend Access, abre `TbConfiguracion` y `TbConfiguracionBackends` (`VG.bas`); la conversión a PostgreSQL no se aborda aquí (queda para `../coexistencia.md`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Caption heredado de otro aplicativo.** La propiedad `Caption` del formulario queda fijada en tiempo de diseño a `"CONDOR - Iniciando..."` (`.form.txt:40`) y solo se sobreescribe en `Form_Timer` a «Iniciando HPS SOLICITUDES ...» (`:50`). Entre el `OnOpen` y el `OnTimer` la ventana muestra el rótulo de Condor.
- **Typo en el subtítulo.** `Caption ="Control y Gestión de las Solcitudes de HPS"` (`.form.txt:1259`): «Solcitudes» debería ser «Solicitudes».
- **`/dev` también llega a `EVE`.** La comprobación `InStr(...) = 0` se hace antes de pasar `Command()` a `EVE` (`:17`; receptor en `VG.bas:103-108`): si el desarrollador usa `/dev` real, `EVE` lo trata como correo y el usuario queda en «no registrado» (a verificar; el comportamiento no se ha observado en runtime).
- **Rama de errores que abre el menú principal.** En el `errores:` de `Form_Timer` se abre `Form0BDOpciones` para «no dejar la aplicación cerrada» (`:78-79`) aunque la sesión no esté validada: un fallo de `EVE` permite entrar al menú con la identidad local de Windows, no la de la aplicación.
- **`Caption` sobreescrito en `OnTimer` compite con el de diseño.** Si el temporizador se desactiva antes por otra ruta, el rótulo de Condor persiste (a verificar: ningún path salta `Me.TimerInterval = 0` antes de `Form_Timer`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: ruta de arranque gestionada por la sesión autenticada del frontend web; traducción del rótulo desde un solo origen (i18n); sustituir el splash por una redirección a `/solicitudes`.

## Criterios de aceptación

Propuestos, no aprobados:

1. La aplicación arranca con el rótulo «HPS SOLICITUDES» desde el primer repintado.
2. Un usuario autenticado abre `Form0BDOpciones` en menos del umbral definido por `EVE` + `CalcularTotalesPendientes`.
3. Una sesión sin permisos válidos no llega al menú principal.
4. Un error de `EVE` muestra el aviso y aborta el arranque sin abrir `Form0BDOpciones`.

## Pendientes

- Nombre real del formulario de inicio declarado en la base (no figura en `src`).
- Forma exacta de `Entorno.CalcularTotalesPendientes` (lectura de `TbSolicitudes` o cálculo completo en `classes/Entorno.cls`): no acotado en este fichero.
- Valor de `EnDesarrollo` y otros `TempVar` fijados por `EVE` (lectura transversal en `VG.bas`).
- Citas en `.form.txt` (líneas 26447, 26465 en Lanzadera son equivalentes al `.form.txt` local; aquí la barra `lblProgresoFondo` y `lblProgresoBarra` están en `1281-1310`).
