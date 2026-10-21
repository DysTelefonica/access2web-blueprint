# Form_Form0BDOpciones

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_Form0BDOpciones.cls` (407 líneas) y `.form.txt` (33 controles `Begin`) · **PRD:** [README](../README.md)

Menú principal de la aplicación tras `Form_frmSplash`: muestra indicadores de tareas pendientes del usuario y globales, abre los formularios de gestión (solicitudes, entidades, configuración, tareas, correos) y reacciona al cierre de los subformularios (`Configuracion`, `TareasTramitadorPendientes`) vía eventos `WithEvents`. Es **no enlazado** (sin `RecordSource` efectivo). Las citas son `Form_Form0BDOpciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_frmSplash` | `DoCmd.OpenForm "Form0BDOpciones"` | `Form_frmSplash.cls:71, 78` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormConfiguracion` | `DoCmd.OpenForm "FormConfiguracion"` desde `ComandoConfiguracion_Click` | `:135` |
| `FormTareasTramitadorPendientes` | `DoCmd.OpenForm "FormTareasTramitadorPendientes"` desde `ComandoTareasPendientes_Click` | `:262` |
| `FormEntidades` | `DoCmd.OpenForm "FormEntidades"` desde `ComandoEntidades_Click` | `:170` |
| `FormCorreosGestion` | `DoCmd.OpenForm "FormCorreosGestion"` desde `ComandoGestionCorreos_Click` | `:204` |
| `FormSolicitudesGestion` | `DoCmd.OpenForm "FormSolicitudesGestion"` desde `ComandoGestionSolicitudes_Click` | `:240` |

Eventos hacia los subformularios: `m_FormConf_Modificado` (`:374-407`) y `m_FormTareas_SolicitudModificada` (`:409-411`) atajan cambios en `Configuracion` y `TareasTramitadorPendientes`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | `Ajustar Me`, llama a `EstablecerDatos m_Error`, muestra/oculta `lblEstado`, gestiona error con `DoCmd.Quit` | `:317-352` |
| Formulario | — | (no `OnTimer`) | — | — |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Recalcula `Entorno.CalcularTotalesPendientes` y repinta con `EstablecerDatos` | `:41-71` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Llama a `CorreoAlAdministrador "Prueba de Error"` como marcador y luego `AbrirAyuda m_Error` | `:73-103` |
| `ComandoConfiguracion` | botón | `OnClick` → `ComandoConfiguracion_Click` | Cierra `FormConfiguracion` si está abierto y lo reabre; captura `m_FormConf = Forms("FormConfiguracion")` para los eventos | `:106-140` |
| `ComandoEntidades` | botón | `OnClick` → `ComandoEntidades_Click` | Abre `FormEntidades` si no está abierto | `:142-175` |
| `ComandoGestionCorreos` | botón | `OnClick` → `ComandoGestionCorreos_Click` | Abre `FormCorreosGestion` si no está abierto | `:177-209` |
| `ComandoGestionSolicitudes` | botón | `OnClick` → `ComandoGestionSolicitudes_Click` | Abre `FormSolicitudesGestion` si no está abierto | `:211-244` |
| `ComandoTareasPendientes` | botón | `OnClick` → `ComandoTareasPendientes_Click` | Cierra y reabre `FormTareasTramitadorPendientes`; captura `m_FormTareas` | `:246-275` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Pregunta «¿Desea realmente salir de la aplicación?» y `DoCmd.Quit acQuitSaveAll` | `:11-26` |
| `lblVersion` | etiqueta | — | «Versión: » + `m_ObjEntorno.VersionAplicacion` (pintada por `EstablecerDatos`) | `:223` |
| `lblEstado` | etiqueta | — | Estado de carga, oculta por defecto tras `AvanceCerrar` | `:343` |
| `lblParametros` | etiqueta | — | «Parámetros INCOMPLETOS» (rojo `683236`) o «Parámetros COMPLETOS» (verde `16737792`) | `:238-243` |
| `lblConfCorreosAutomaticos` | etiqueta | — | Texto de configuración de correos automáticos (invisible por defecto) | `:230` |
| `lblTareasPendientes` | etiqueta | — | Contador global y por usuario según `EstablecerIndicadores` | `:393-407` |

## Flujos

### Arranque del menú (`Form_Open`, `:317-352`)

1. `On Error GoTo errores`; `DoCmd.Hourglass True`; `Ajustar Me` (`:330-333`).
2. `EstablecerDatos m_Error` (`:336`). Si devuelve error, `Err.Raise 1000` (`:337-339`) salta a la rama `errores`.
3. `AvanceCerrar` y `Me.lblEstado.Visible = False` (`:342-345`).
4. Si hay error, sale con `DoCmd.Quit acQuitSaveAll` salvo en `EnDesarrollo = "Sí"` (`:349-351`).

### Pintado de datos (`EstablecerDatos`, `:218-269`)

1. `Me.Caption = m_TituloFormulario`; `lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion`; pinta configuración de correos automáticos (`:222-225`).
2. `EstablecerIndicadores m_Error` calcula y pinta `lblTareasPendientes` (`:227-229`).
3. Activa los comandos `ComandoTareasPendientes`, `ComandoConfiguracion`, `ComandoGestionSolicitudes` (`:236-238`).
4. Actualiza `lblParametros` con el color correspondiente a `m_FaltaAlgunParametroEnConfiguracion` (`:240-243`).
5. Hay un bloque de `If Application.TempVars("ActivadoCorreoAutomatico") = "No"` comentado en `:251-258`; en su lugar siempre muestra `lblConfCorreosAutomaticos.Visible = False` y deja `ComandoGestionCorreos.Enabled = True` (`:259-262`).

### Indicadores (`EstablecerIndicadores`, `:380-407`)

1. Lee `m_ObjEntorno.TotalPendientesGlobal` y `m_ObjEntorno.TotalPendientesUsuario` (calculados en `Splash` o en `ComandoActualizar_Click`).
2. Si `EsTecnico = EnumSiNo.Sí`, pinta `Pendientes (0/<global>)` (sin contador del usuario) (`:394-397`).
3. En caso contrario pinta `Pendientes (<usuario>/<global>)` (`:400-405`); pinta rojo si hay pendientes, verde si no.

### Apertura de subformularios

Los seis manejadores siguen el mismo patrón de cinco pasos:

1. `DoCmd.Hourglass True`; `Ajustar Me` (cuando aplica); `Me.AllowEdits = False`.
2. Cierra el formulario destino si está abierto (`ComandoConfiguracion_Click` y `ComandoTareasPendientes_Click` lo cierran antes de abrirlo, líneas `:127-131, 254-259`; los demás solo abren si no está abierto, líneas `:170-173, 204-207, 240-243`).
3. `DoCmd.OpenForm "<destino>"`.
4. Reengancha los eventos si los hay: `Set m_FormConf = Forms("FormConfiguracion")` (`:135-137`), `Set m_FormTareas = Forms("FormTareasTramitadorPendientes")` (`:266-268`).
5. `DoCmd.Hourglass False`; `Me.AllowEdits = True`. En `errores:`, `CorreoAlAdministrador m_Error` y `MsgBox` con el detalle.

### Eventos de subformularios

- `m_FormConf_Modificado` (`:374-407`): repinta `lblParametros` (estados `INCOMPLETOS`/`COMPLETOS` con sus `ForeColor`), vuelve a leer `TextoConfiguracionCorreosAutomaticos`. Mismo patrón de error que los comandos.
- `m_FormTareas_SolicitudModificada` (`:409-411`): re-llama a `EstablecerIndicadores` para refrescar contadores tras un cambio de tarea.

### Salida (`cmdSalir_Click`, `:11-26`)

Confirma con `MsgBox` y `DoCmd.Quit acQuitSaveAll`. La rama `errores:` muestra `MsgBox` con `Err.Description`.

## Reglas, validaciones y cálculos

- Todos los manejadores siguen el patrón `m_Error` (`On Error GoTo errores` → `If m_Error <> "" Then Err.Raise 1000`) que distingue los errores lanzados internamente de los del runtime; este patrón es **propio de la aplicación** y consistente con Lanzadera y Condor.
- `FormularioAbierto("...")` decide si abrir o no abrir el subformulario (no se ha verificado la función; se encuentra en `modules/FU.bas`).
- `EsTecnico` y `EnumSiNo.Sí` son globales del módulo; el resto de cálculo de indicadores vive en `classes/Entorno.cls`.
- `m_NombreTramitadorAlInicio` se declara en `:5` pero no se usa; candidata a retirar.
- El conmutador «Activar correos automáticos» está comentado pero su `Else` deja `ComandoGestionCorreos.Enabled = True` siempre activo (`:251-262`): la rampa de migración de correos automáticos en producción no está reflejada en el código.
- La variable `pregunta` se reusa en seis manejadores sin `Dim` local pese a `Option Explicit`: indica `Dim` omitido en algún procedimiento o se trata de una variable implícita (a verificar con debug).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Versión y textos del entorno | `TbConfiguracion` (versión, indicadores, correos automáticos) | lectura vía `classes/Entorno.cls` y `classes/ConfiguracionOperaciones.cls` | `:223, 230` |
| Totales de tareas pendientes | `TbSolicitudes` / `TbResponsables` (agregación por usuario y global) | DAO agregado en `m_ObjEntorno.CalcularTotalesPendientes` | `:46, 385-388` |
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` (Lanzadera) | lectura implícita en `EVE` y `comprobarPermisosAplicacion` | transversal (`VG.bas`) |

Sin escrituras en este formulario: el menú solo dispara acciones.

## Coexistencia

La apertura de subformularios no comparte datos persistentes con este menú (los subformularios tienen su propio `RecordSource` cuando enlazan). La coexistencia se centra en cómo cada subformulario se conecta al backend; queda para `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`ComandoAyuda_Click` envía un correo real al administrador con la cadena «Prueba de Error»** (`:83-95`): el botón «Ayuda» no abre una ayuda, dispara un email de marcador. El método `AbrirAyuda` se llama después (`:91`); el contenido del correo depende de `m_Error = ""` por lo que llega el literal. Debe revisarse si quedaban marcadores en producción o si la rampa se apagó y nunca se cerró la llamada.
- **Indicadores incoherentes para perfil técnico** (`:393-407`): cuando el usuario es técnico la etiqueta muestra `Pendientes (0/<total>)` con el contador `0` literal, no `0` como valor nulo ni un texto distinto. Un técnico no ve su propio contador, solo el global.
- **`lblConfCorreosAutomaticos` siempre oculto** (`:259`): la migración a correos automáticos según `TempVars("ActivadoCorreoAutomatico")` está comentada. Si en producción los correos están desactivados, este menú los deja igual de activados.
- **Botones «Tareas» y «Configuración» se cierran y reabren siempre** (`:127-131, 254-259`) mientras los demás usan `If Not FormularioAbierto`. Comportamiento asimétrico (a verificar si es intencional para forzar recarga).
- **`Form_Open` aborta con `DoCmd.Quit` salvo en `EnDesarrollo = "Sí"`** (`:348-352`): una sesión de desarrollo deja el usuario atrapado en el menú tras un error grave.
- **`m_NombreTramitadorAlInicio` declarada y nunca usada** (`:5`): candidata a retirar; posibles referencias dinámicas que no se ven desde `src`.
- **`ComandoTareasPendientes_Click` re-engancha `m_FormTareas` cada vez** (`:265-269`) sin `Set m_FormTareas = Nothing` previo: los eventos `WithEvents` podrían mantener referencias colgantes tras un cierre de Access no controlado (a verificar).
- **Lectura repetida de `m_ObjEntorno.TotalPendientesGlobal/Usuario`** (`:385-388`) sin sincronización explícita entre el click «Actualizar» y la apertura de subformularios: si la base cambia mientras el menú está abierto, los indicadores quedan obsoletos hasta pulsar «Actualizar».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: reemplazar la matriz de botones por un panel de dashboard asíncrono; sustituir `WithEvents` por un store compartido o notificaciones en backend; validar la rampa de correos automáticos con un único origen de verdad (`Configuracion.CorreosAutomaticos`).

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario abre cada subformulario desde su botón; los contadores cambian tras editar en `TareasTramitadorPendientes` o en `Configuracion`.
2. Un fallo en `EstablecerDatos` cierra la aplicación salvo en modo desarrollo.
3. El botón «Ayuda» abre la ayuda real y no envía correo al administrador.
4. Los indicadores se actualizan al pulsar «Actualizar» y al volver de un subformulario.
5. El menú distingue «Pendientes del usuario» vs «Pendientes global» para todos los perfiles (técnico, tramitador, administrador).

## Pendientes

- Orígenes reales de `m_TituloFormulario`, `m_ObjEntorno`, `EsTecnico`, `EnumSiNo.Sí`: dónde se declaran y cómo se inicializan (probable `modules/VG.bas` y `classes/Entorno.cls`).
- Si `m_NombreTramitadorAlInicio` se usa mediante enlace dinámico no resuelto por CodeGraph.
- Comportamiento real del comando «Ayuda» en producción y si la rampa a `AbrirAyuda` quedó completa.
- Si los seis subformularios permiten doble apertura o si se impide con `FormularioAbierto` desde `Form_Form0BDOpciones.cls` y desde otros menús.
