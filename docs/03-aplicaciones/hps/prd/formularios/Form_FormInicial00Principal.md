# Form_FormInicial00Principal

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial00Principal.cls` (149 líneas) y `.form.txt` (1262 líneas) · **PRD:** [README](../README.md)

Contenedor raíz de la aplicación HPS. Aloja dos subformularios (`SubFormLateral` y `SubFormCentral`) y sirve de **ancla obligatoria** para los nueve formularios de consulta del bloque 03: `EstablecerFiltroUsuarios` exige que `FormularioAbierto("FormInicial00Principal")` sea cierto y, si lo es, reescribe el `RecordSource` del subformulario de resultados que cuelga de `SubFormCentral → SubFormResultados` (`modules/Filtro.bas:21-26, 107`). El formulario no tiene lógica propia de filtrado: la búsqueda global se delega en `FormInicial06BusquedaGlobal` y la sincronización de caches en `ComandoActualizar`. Las citas son `Form_FormInicial00Principal.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Arranque de la aplicación | Se abre como primer formulario (no hay `DoCmd.OpenForm` en `src/`) | `VariablesEntorno.bas:380` carga `m_ObjEntorno` antes de cualquier apertura |
| `SubFormLateral` (carga) | `Form_Load` fija `SubFormLateral.SourceObject = "FormInicial01Lateral"` | `:122`, `.form.txt:1013` |
| `EstablecerFiltroUsuarios` (los 9 filtros del bloque 03) | `Forms("FormInicial00Principal").Controls("SubFormCentral").Form.Controls("SubFormResultados").Form` reasignando `RecordSource` | `Filtro.bas:24-26, 107` |
| `Form_FormUsuario07SICAEdicion` | Si `SubFormCentral` no es `FormInicial02ConsultasPrincipal`, lo monta y desciende | `Form_FormUsuario07SICAEdicion.cls:186, 195-196` |
| `FormInicial06BusquedaGlobal` | `ComandoBuscar` lo abre / reutiliza y le pasa la palabra clave | `:55-62` |

## Controles y eventos

`DefaultView =0` (`.form.txt:11`); `OnLoad ="[Event Procedure]"` (`.form.txt:374`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Ajusta tamaño, monta `SubFormLateral`, aplica paleta y compone título | `:82-135`, `.form.txt:374` |
| `CmdSalir` | botón | `Click` → `CmdSalir_Click` | Cierra sin guardar | `:10-12`, `.form.txt:950` |
| `ComandoActualizar` | botón | `Click` → `ComandoActualizar_Click` | Resincroniza caches locales y refresca contadores | `:15-41`, `.form.txt:903` |
| `ComandoBuscar` | botón | `Click` → `ComandoBuscar_Click` | Lanza la búsqueda global con `PalabraClave` | `:43-79`, `.form.txt:859` |
| `PalabraClave` | caja de texto | `OnKeyDown` → `PalabraClave_KeyDown` | Si se pulsa Intro, llama a `ComandoBuscar_Click` | `:138-149`, `.form.txt:832` |
| `SubFormLateral` | subformulario | — | Aloja `FormInicial01Lateral` (lateral con todas las entradas del menú) | `:122`, `.form.txt:1013` |
| `SubFormCentral` | subformulario | — | Aloja el formulario activo del contexto (consultas, edición, histórico, indicadores…) | `.form.txt:1045` |
| `EncabezadoDelFormulario` | sección | — | `BackColor = m_ObjEntorno.CabeceraBackColor` | `:110`, `.form.txt:812` |
| `PieDelFormulario` | sección | — | Mismo `BackColor` que la cabecera | `:111`, `.form.txt:1076` |
| `lblAvance` | etiqueta | — | Receptor del avance de sincronización | `:21`, `.form.txt:1092` |
| `Etiqueta1`, `Etiqueta3` | etiquetas (`Secundario0:`, `Secundario2:`) | — | Marcadores de los subformularios; sin lógica | `.form.txt:1027, 1060` |

## Flujos

### Apertura (`Form_Load`, `:82-135`)

1. `Application.Echo False` y `DoCmd.Hourglass True` con `VBA.DoEvents` interpuesto (`:88-91`): bloquea redibujado mientras monta.
2. `AjustarTamaño Me` (`:93`; `FG.bas:901`).
3. Fija `Application.TempVars("EnDesarrollo") = EnumSiNo.No` (`:94-95`).
4. `Application.Echo True` (`:97`).
5. Si `m_ObjEntorno Is Nothing`, llama a `EVE m_Error` (`:98-104`); si vuelve con error, `Err.Raise 1000`. `EVE` carga `m_ObjEntorno` desde `TbConfiguracionHPS` (`VariablesEntorno.bas:380-432`).
6. Pinta cabecera y pie con `m_ObjEntorno.CabeceraBackColor` (`:110-111`).
7. `CopiarEstiloBoton` sobre `ComandoBuscar` (`:113`) y `ComandoActualizar` (`:117`) con comprobación de `m_Error` tras cada llamada (`:114, 118`; `FG.bas:2435`).
8. `Me.Caption = "Usuario conectado: " & m_ObjUsuarioConectado.Nombre` (`:121`; `VariablesEntorno.bas:267`).
9. `Me.SubFormLateral.SourceObject = "FormInicial01Lateral"` (`:122`) y fin del hourglass (`:124-125`).
10. En error: `CorreoAlAdministrador` solo si `Err.Number <> 1000` (`:129-131`), `DoCmd.Quit acQuitSaveNone` (`:134`) — **un fallo al abrir tira la sesión, no la continúa**.

### Salir (`CmdSalir_Click`, `:10-12`)

`DoCmd.Close acForm, Me.Name, acSaveNo` directo, sin `On Error`.

### Actualizar caches (`ComandoActualizar_Click`, `:15-41`)

1. `VBA.DoEvents` + `DoCmd.Hourglass True` (`:17-19`), `m_Error = ""` (`:20`).
2. `Set lbl = Me.lblAvance` (`:21`); `lbl` no se declara en este `.cls`, se resuelve contra un `Label` del módulo `Form_FormInicial`.
3. `SincronizarCachesLocalesFrontend p_Error:=m_Error` (`:22`; `Mod_StartupCacheInitialization.bas:62-115`); si vuelve con error, `Err.Raise 1000` (`:24`).
4. `EstablecerContadores` (`:26`; `FG.bas:1440`) y `AvanceCerrar` (`:27`; `FG.bas:4415`).
5. `DoCmd.Hourglass False` y `Exit Sub` (`:29, 32`).
6. En error: `CorreoAlAdministrador` salvo 1000 (`:35-36`) y `MsgBox` crítico (`:40`).

### Buscar (`ComandoBuscar_Click`, `:43-79`)

1. `Me.AllowEdits = False` (`:52`); `Palabra = Nz(Me.PalabraClave, "")` (`:53`).
2. Si `FormInicial06BusquedaGlobal` no está abierto, `DoCmd.OpenForm` (`:55-56`).
3. Si está abierto tras la apertura, fija `frm.Controls("PalabraClave") = Palabra` (`:61`) y dispara `Form_FormInicial06BusquedaGlobal.PalabraClave_Exit False` (`:62`).
4. `DoCmd.Hourglass False` y `Me.AllowEdits = True` en ambas ramas (`:67-69, 76-77`).
5. En error: `CorreoAlAdministrador` salvo 1000 (`:73-74`), `Me.AllowEdits = True` y `MsgBox` (`:76, 78`).

### Atajo de teclado (`PalabraClave_KeyDown`, `:138-149`)

Si `KeyCode = vbKeyReturn` y `ComandoBuscar.Enabled = True`, desplaza foco al botón y ejecuta `ComandoBuscar_Click`; `KeyCode = 0` cancela el pitido y la acción por defecto de Enter (`:139-148`).

## Reglas, validaciones y cálculos

- **El contenedor es el requisito de todo filtrado.** `EstablecerFiltroUsuarios` (`modules/Filtro.bas:21-23`) sale sin tocar nada si `FormularioAbierto("FormInicial00Principal") = False`; los nueve formularios del bloque 03 componen un `WHERE` y lo aplican al `RecordSource` del subformulario de resultados (`Filtro.bas:107`). **Sin este formulario abierto, ningún filtro surte efecto.**
- **Atadura a dos niveles.** La cadena `FormInicial00Principal → SubFormCentral → SubFormResultados` (`Filtro.bas:24-26`) es el contrato que ata toda la pantalla: cualquier cambio de `SourceObject` en `SubFormCentral` debe respetar que el subformulario de resultados vive en el segundo nivel de anidamiento. `Form_FormUsuario07SICAEdicion.cls:186, 195-196` reifica ese contrato.
- **Identidad del usuario conectado.** El título de la ventana se compone con `m_ObjUsuarioConectado.Nombre` (`:121`); ese objeto se hidrata en `VariablesEntorno.bas:392-401`.
- **Salir cierra el contenedor; si la apertura falla, aborta la sesión.** `CmdSalir_Click` no captura errores (`:11`), pero `Form_Load` sí: el raise 1000 conduce a `DoCmd.Quit acQuitSaveNone` (`:134`).
- **Variable temporal de marca.** `Application.TempVars("EnDesarrollo") = EnumSiNo.No` se fija aquí (`:94`) y desde otros formularios se lee.
- **`SubFormLateral` expone su nombre por contrato.** El `.form.txt:1013` fija `Name ="SubFormLateral"`, y `Form_Load` lo reescribe en cada apertura (`:122`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Hidratación del entorno al abrir | `TbConfiguracionHPS` (`EVE` → `m_ObjEntorno.CargarConfiguracion`) | `SELECT` + cache en memoria | `:99`, `VariablesEntorno.bas:380-432` |
| Datos del usuario conectado (título) | `TbUsuarios` (vía `Constructor.getUsuario`) | `SELECT` indirecto | `:121`, `VariablesEntorno.bas:392` |
| Sincronización de caches locales | múltiples `Tb…Local` | `DELETE + INSERT` masivos | `:22`, `Mod_StartupCacheInitialization.bas:62-115` |
| Apertura del formulario global de búsqueda | ninguna directa | `DoCmd.OpenForm` | `:56` |
| Aplicación del filtro | `TbDatosLocal` (vía `EstablecerFiltroUsuarios`) | `RecordSource` reasignado al subform de resultados | `Filtro.bas:107` |

No hay `INSERT`/`UPDATE`/`DELETE` directos en este `.cls` (`:1-149`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la **cadena hardcodeada** `Forms("FormInicial00Principal").Controls("SubFormCentral").Form.Controls("SubFormResultados").Form` (`Filtro.bas:24-26`) convierte a este formulario en un **ancla física**: cualquier renombrado rompe los nueve filtros del bloque 03 sin error de compilación; `Application.TempVars` (`VariablesEntorno.bas:330-332`) y los **`m_Obj…`** globales (`VariablesEntorno.bas:221-227, 256, 267-268`) son estado compartido entre formularios sin equivalente directo web; `Set lbl = Me.lblAvance` (`:21`) sigue siendo una variable de **módulo no declarada** en este `.cls` (a verificar contra el binario); el redibujado se gobierna con `Application.Echo False/True` y `DoCmd.Hourglass True/False` (`:88-91, 17-19, 29, 67-68`); y `DoCmd.Quit acQuitSaveNone` (`:134`) aborta la sesión completa en caso de error de apertura.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La cadena de acceso `Forms("FormInicial00Principal").Controls("SubFormCentral").Form.Controls("SubFormResultados").Form` (`Filtro.bas:24-26`) está **hardcodeada**; cualquier cambio de nombre en cualquier eslabón rompe los nueve filtros sin error de compilación.
- `m_ObjUsuarioConectado.Nombre` (`:121`) puede ser `Nothing` si este `.cls` se monta antes de que termine `EVE`; el `.cls:99-104` lo evita con el `Err.Raise 1000`, pero ese caso aborta la sesión con `DoCmd.Quit` y deja al usuario sin saber por qué.
- `Set lbl = Me.lblAvance` (`:21`) escribe una variable implícitamente cualificada de módulo. VBA lo admite, pero rompe cualquier rename del control y vuelve opaca la dependencia con el formulario padre.
- `Application.TempVars("EnDesarrollo")` se reescribe a `No` en cada apertura (`:94`), pisando el valor que otro flujo hubiera podido fijar.
- `ComandoBuscar` abre un formulario de búsqueda global sin comprobar que el usuario realmente quiere navegar (`:55-62`); pulsar Intro con `PalabraClave` vacío abre y cierra el destino sin aviso.
- `Form_Load` hace `DoCmd.Quit acQuitSaveNone` en cualquier error (`:134`); una falla transitoria de `CopiarEstiloBoton` (`FG.bas:2435`) tira la sesión de Access completa.
- El `.form.txt` mantiene el `Caption` antiguo (`.form.txt:46`) hasta que `Form_Load` lo sobrescribe (`:121`); entre el inicio del `Load` y la asignación, la barra muestra el último usuario persistido en el binario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: extraer un **servicio de sesión** que reciba los nueve formularios del bloque 03 (la cadena se convierte en un endpoint único de filtrado con contexto inyectado); mover el avance de sincronización a un **`SignalR` / WebSocket**; sustituir `DoCmd.Quit` por una página de error que permita reintentar la carga del contenedor sin perder la sesión.

## Criterios de aceptación

Propuestos, no aprobados:

1. `FormularioAbierto("FormInicial00Principal")` debe ser cierto para que cualquier filtro del bloque 03 tenga efecto.
2. La cadena `Forms("FormInicial00Principal").Controls("SubFormCentral").Form.Controls("SubFormResultados").Form` resuelve siempre.
3. El título de la ventana muestra el nombre del usuario conectado.
4. `ComandoActualizar` resincroniza caches y refresca contadores sin perder el hourglass mientras dura.
5. `ComandoBuscar` con `PalabraClave` vacío abre y cierra el destino sin error; con valor, propaga la palabra antes de mostrar resultados.
6. Un error en `Form_Load` (no 1000) cierra la sesión con `DoCmd.Quit`, correo al administrador y un `MsgBox` crítico previo.

## Pendientes

- Si `SubFormCentral` aloja otros formularios además de `FormInicial02ConsultasPrincipal` (a verificar contra `.form.txt:1045`).
- Nombre real del control `lbl` referenciado implícitamente en `:21` — comprobar que pertenece a `Form_FormInicial`.
- Si hay un `AutoExec` o macro del binario que abre `FormInicial00Principal` (no aparece `DoCmd.OpenForm` en `src/`).
