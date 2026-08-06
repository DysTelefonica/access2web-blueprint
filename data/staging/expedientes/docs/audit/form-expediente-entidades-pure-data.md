# Audit — Form_FormExpedienteEntidades (Phase 3.3 / PR-9)

> **Change**: `forms-thin-coverage` (Phase 3.3 — PR-9, batch of 6 forms)
> **Form**: `Form_FormExpedienteEntidades` (pestaña de entidades del expediente: Comerciales / CPVs / Lugares / PECAL / RACs / Responsables / Anualidades)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** |
| Helper extraction | **PENDING** (~25 helpers planned) |
| Binario sync | PENDING |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 1802 |
| Handlers totales | 28+ (click + WithEvents subscribers + Form_Load) |
| Public methods | 11 (`RellenarListaPECAL`, `RellenarListaLugares`, `RellenarListaComerciales`, `RellenarListaCPVs`, `RellenarListaRACs`, `RellenarListaResponsables`, `RellenarListaAnualidades`, `RellenarListas`, `EstablecerDatos`, `CambiarLineaListaResponsables`) — **MANTENER firmas** |
| Public events | 6 (`m_FormComercial_Seleccionar`, `m_FormCPV_Seleccionar`, `m_FormLugar_Seleccionar`, `m_FormPECAL_Seleccionar`, `m_FormRAC_Seleccionar`, `m_FormUsuarios_Seleccionar`) |
| WithEvents | 6 (`m_FormPECAL`, `m_FormLugar`, `m_FormComercial`, `m_FormCPV`, `m_FormRAC`, `m_FormUsuarios`) — **MANTENER** (subscribers a Seleccionar de los forms hijo) |
| Private Sub helpers | 5 (`AltaComercial`, `AltaCPV`, `AltaLugarEjecucion`, `AltaPECAL`, `AltaRAC`) — lógica a extraer |
| MsgBox call sites | ~24 (casi todos los handlers tienen uno en el bloque `errores:`) |
| DoCmd.OpenForm en .cls | 5 (ComandoElegirUsuario, ComandoElegirPECAL, ComandoElegirRAC, ComandoElegirComercial, ComandoElegirCPV, ComandoElegirLugar) |
| DoCmd.Close en .cls | 5 (idem, para cerrar form previo si está abierto) |
| Forms abiertas | `FormUsuariosGestion`, `FormPECALESGestion`, `FormRACSGestion`, `FormComercialesGestion`, `FormCPVsGestion`, `FormLugarEjecucionGestion` |

## 3. CONTROLES REFERENCIADOS (Me.X)

| Control | Tipo | Usado en |
|---|---|---|
| `Responsable` | TextBox | ComandoAltaReponsable_Click, ComandoLimpiarUsuario_Click, m_FormUsuarios_Seleccionar |
| `EsJefeProyecto` | CheckBox | ComandoAltaReponsable_Click, ComandoLimpiarUsuario_Click |
| `CorreoSiempre` | CheckBox | ComandoAltaReponsable_Click, ComandoLimpiarUsuario_Click |
| `AÑO` | TextBox | ComandoAltaAnualidad_Click, ComandoLimpiarAnualidad_Click |
| `BIIVA`, `BIIPSI`, `BIIGIC`, `BIEXENTA` | TextBox | ComandoAltaAnualidad_Click, ComandoLimpiarAnualidad_Click |
| `ListaAnualidad` | ListBox | ComandoAltaAnualidad_Click, RellenarListaAnualidades, RellenarListas |
| `ListaComerciales` | ListBox | AltaComercial, ComandoEliminarComercial_Click, RellenarListaComerciales, RellenarListas |
| `ListaCPVs` | ListBox | AltaCPV, ComandoCopiarCPV_Click, ComandoEliminarCPV_Click, RellenarListaCPVs, RellenarListas |
| `ListaLugaresEjecucion` | ListBox | AltaLugarEjecucion, ComandoCopiarLugar_Click, ComandoEliminarLugarEjecucion_Click, RellenarListaLugares, RellenarListas |
| `ListaPECALES` | ListBox | AltaPECAL, ComandoEliminarPECAL_Click, RellenarListaPECAL, RellenarListas |
| `ListaRACs` | ListBox | AltaRAC, ComandoEliminarRACs_Click, RellenarListaRACs, RellenarListas |
| `ListaResponsables` | ListBox | ComandoCambiarEnvios_Click, ComandoCambiarJP_Click, ComandoEliminarReponsable_Click, ListaResponsables_Click, RellenarListaResponsables, RellenarListas, CambiarLineaListaResponsables |
| `ComandoCambiarJP`, `ComandoCambiarEnvios` | CommandButton | ListaResponsables_Click |
| `AllowEdits` | Form property | Form_Load, ComandoCambiarEnvios_Click, ComandoCambiarJP_Click, ComandoCopiarCPV_Click, ComandoCopiarLugar_Click |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `ExpedienteOperaciones` | `Expediente`, `RegistrarAnualidad(p_Anualidad, p_Error)`, `RegistrarComercial(p_IDComercial, p_Error)`, `RegistrarCPV(p_IDCPV, p_Error)`, `RegistrarLugarEjecucion(p_IDLugar, p_Error)`, `RegistrarPECAL(p_IDPEcal, p_Error)`, `RegistrarRAC(p_IDRAC, p_Error)`, `RegistraResponsable(p_Responsable, p_Error)`, `EditarEnvioCorreoResponsable(p_IDUsuario, p_EnvioFinal, p_Error)`, `EditarJPResponsable(p_IDUsuario, p_JPFinal, p_Error)`, `EliminarAnualidad(p_Año, p_Error)`, `EliminarComercial(p_IDComercial, p_Error)`, `EliminarCPV(p_IDCPV, p_Error)`, `EliminarLugarEjecucion(p_IDLugar, p_Error)`, `EliminarPECAL(p_IDPEcal, p_Error)`, `EliminarRAC(p_IDRAC, p_Error)`, `EliminarResponsable(p_IDUsuario, p_Error)` |
| `ExpedienteDTO` (global `m_ObjExpedienteDTOActivo`) | `Expediente`, `ColComerciales`, `ColCPVs`, `ColLugaresEjecucion`, `ColPECALES`, `ColRACs`, `ColResponsables`, `ColAnualidades` |
| `Expediente` | `IDExpediente`, `Comerciales`, `CPVs`, `LugaresEjecucion`, `PECALES`, `RACs`, `Responsables`, `Anualidades` |
| `Comercial` | `IDComercial`, `Comercial` |
| `CPV` | `IDCPV`, `CPV` |
| `LugarEjecucion` | `IDLugarEjecucion`, `LugarEjecucion` |
| `PECAL` | `IDPECAL`, `PECAL` |
| `RAC` | `IDRAC`, `RAC` |
| `ExpedienteResponsable` | `IdUsuario`, `EsJefeProyecto`, `CorreoSiempre`, `USUARIO.Nombre` |
| `USUARIO` | `ID`, `Nombre` |
| `ExpedienteAnualidad` | `AÑO`, `Presupuesto` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `FormularioAbierto` | Module call (verifica si form ya está abierto antes de `DoCmd.OpenForm`) |
| `Forms("FormUsuariosGestion")` etc. | Asignación a `m_FormX` (WithEvents) — **MANTENER** (necesario para recibir `_Seleccionar`) |
| `constructor.getComercial / getCPV / getLugarEjecucion / getPecal / getRAC / getUsuario` | Module call (lookup por ID) |
| `constructor.getUsuario` | Module call (también) |
| `getExpedientePECALES / LugaresEjecucion / Comerciales / CPVS / RACS / Responsables / Anualidades` | Module call (refresh collections) |
| `CopiarAlPortapapeles` | Module call (clipboard) |
| `EsAdministrador`, `AbiertoParaEditar` | Module call (auth/state flags) |
| `MostrarPopupProgreso`, `CerrarPopupProgreso` | Module call (UI guard) |
| `CorreoAlAdministrador` | Module call (error notification) |
| `pregunta = MsgBox(...)` | **inline en handlers (anti-pattern)** — la constante `pregunta` parece no estar declarada |

## 6. ANTI-PATTERN DETECTADO

⚠️ **BLOQUE COMÚN**: todos los handlers tienen el mismo patrón de UI guard:
```vb
VBA.DoEvents
DoCmd.Hourglass True
VBA.DoEvents
' ... lógica ...
VBA.DoEvents
DoCmd.Hourglass False
VBA.DoEvents
```
y el mismo bloque de error:
```vb
On Error GoTo errores
' ...
errores:
If Err.Number <> 1000 Then
    m_Error = "Al X se ha producido el error n: " & ...
    CorreoAlAdministrador m_Error
End If
DoCmd.Hourglass False
pregunta = MsgBox(m_Error, vbCritical, "Error")
```
Esto es ruido mecánico — el helper NO debe repetirlo. La UI guard se queda en el form (regla #1), pero el helper retorna JSON con el error.

⚠️ **`pregunta = MsgBox(...)`** — la variable `pregunta` no está declarada en el módulo. Esto es un bug pre-existente (no compila en `Option Explicit`). El patrón correcto sería `MsgBox ...` (sin asignación) o `Dim pregunta As VbMsgBoxResult`.

⚠️ **`On Error GoTo errores + Err.Raise 1000`** — patrón de "control flow por error". El helper debería usar `p_Error` (regla Telefónica D&S), no `Err.Raise 1000`.

⚠️ **`m_ObjExpedienteDTOActivo` global** — el helper recibe un Dictionary wrapper, no la global. El form hace el wrap antes de llamar (regla #9 — helpers sin globals).

## 7. RECOMENDACIÓN DE PATRÓN

**Pattern**: SPECIAL — entidades tab con 7 sub-entidades + 2 mutaciones de responsables. Cada entidad tiene su propio par Alta/Baja; el patrón es repetitivo (mismo template, distinto método DAO). El helper NO recibe Form ref — recibe un Dictionary wrapper con `{Expediente, ColComerciales, ColCPVs, ColLugaresEjecucion, ColPECALES, ColRACs, ColResponsables, ColAnualidades}`.

### 7.1 Helpers planeados (~25)

**Alta / Eliminar (10)**:
| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteEntidades_AltaComercial` | `(p_DTO, p_IDComercial, p_Error)` | `constructor.getComercial` + `ExpedienteOperaciones.RegistrarComercial` + refresh `ColComerciales` |
| `ExpedienteEntidades_EliminarComercial` | `(p_DTO, p_IDComercial, p_Error)` | `constructor.getComercial` + `ExpedienteOperaciones.EliminarComercial` + refresh |
| `ExpedienteEntidades_AltaCPV` | `(p_DTO, p_IDCPV, p_Error)` | mismo patrón |
| `ExpedienteEntidades_EliminarCPV` | `(p_DTO, p_IDCPV, p_Error)` | mismo patrón |
| `ExpedienteEntidades_AltaLugar` | `(p_DTO, p_IDLugar, p_Error)` | mismo patrón |
| `ExpedienteEntidades_EliminarLugar` | `(p_DTO, p_IDLugar, p_Error)` | mismo patrón |
| `ExpedienteEntidades_AltaPECAL` | `(p_DTO, p_IDPECAL, p_Error)` | mismo patrón |
| `ExpedienteEntidades_EliminarPECAL` | `(p_DTO, p_IDPECAL, p_Error)` | mismo patrón |
| `ExpedienteEntidades_AltaRAC` | `(p_DTO, p_IDRAC, p_Error)` | mismo patrón |
| `ExpedienteEntidades_EliminarRAC` | `(p_DTO, p_IDRAC, p_Error)` | mismo patrón |

**Anualidades (2)**:
| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteEntidades_AltaAnualidad` | `(p_DTO, p_Año, p_BIIVA, p_BIIPSI, p_BIIGIC, p_BIEXENTA, p_Error)` | Validación IsNumeric + `RegistrarAnualidad` + refresh |
| `ExpedienteEntidades_EliminarAnualidad` | `(p_DTO, p_Año, p_EsAdmin, p_Error)` | Auth check + `EliminarAnualidad` + refresh |

**Responsables (4)**:
| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteEntidades_AltaResponsable` | `(p_DTO, p_ResponsableNombre, p_EsJefeProyecto, p_CorreoSiempre, p_Error)` | `constructor.getUsuario` + `RegistraResponsable` + refresh |
| `ExpedienteEntidades_EliminarResponsable` | `(p_DTO, p_IDUsuario, p_Error)` | `constructor.getUsuario` + `EliminarResponsable` + refresh |
| `ExpedienteEntidades_CambiarJP` | `(p_DTO, p_IDUsuario, p_JPFinal, p_LineaActual, p_Error)` | `EditarJPResponsable` + retorna `{lineaFinal}` |
| `ExpedienteEntidades_CambiarEnvio` | `(p_DTO, p_IDUsuario, p_EnvioFinal, p_LineaActual, p_Error)` | `EditarEnvioCorreoResponsable` + retorna `{lineaFinal}` |

**Form-level (8)**:
| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteEntidades_Form_Load` | `(p_DTO, p_EsAdmin, p_Error)` | Orquesta inicialización |
| `ExpedienteEntidades_EstablecerDatos` | `(p_DTO, p_EsAdmin, p_Error)` | Decide enable de controles Tag="EJECUTIVO" |
| `ExpedienteEntidades_RellenarListas` | `(p_DTO, p_Error)` | Compone y retorna JSON con los 7 rowSources |
| `ExpedienteEntidades_RellenarListaComerciales` | `(p_DTO, p_Error)` | Walk ColComerciales → rowSource |
| `ExpedienteEntidades_RellenarListaCPVs` | `(p_DTO, p_Error)` | walk ColCPVs → rowSource |
| `ExpedienteEntidades_RellenarListaLugares` | `(p_DTO, p_Error)` | walk ColLugaresEjecucion → rowSource |
| `ExpedienteEntidades_RellenarListaPECALES` | `(p_DTO, p_Error)` | walk ColPECALES → rowSource |
| `ExpedienteEntidades_RellenarListaRACs` | `(p_DTO, p_Error)` | walk ColRACs → rowSource |
| `ExpedienteEntidades_RellenarListaResponsables` | `(p_DTO, p_Error)` | walk ColResponsables → rowSource (4-col) |
| `ExpedienteEntidades_RellenarListaAnualidades` | `(p_DTO, p_Error)` | walk ColAnualidades → rowSource |
| `ExpedienteEntidades_CambiarLineaListaResponsables` | `(p_DTO, p_LineaActual, p_LineaFinal, p_Error)` | Walk RowSource, replace línea |

**UI misc (1, en form)**:
- `CopiarAlPortapapeles` — ya existe en constructor.bas, queda en form como llamada UI.

### 7.2 UI orchestration que se queda en el form (regla #1)

- `ComandoElegirX_Click` (5) — `DoCmd.Close` + `DoCmd.OpenForm` + `Set m_FormX = Forms("...")` — UI puro.
- `m_FormX_Seleccionar(...)` (6 WithEvents subscribers) — recibe la entidad seleccionada y delega al helper de Alta. El form NO llama al helper directamente: primero llama al `AltaX` helper que hace todo (registro + refresh), luego el form limpia la selección visual.
- `ComandoAltaX_Click` (7) — llama al helper `AltaX`, refresca visual.
- `ComandoEliminarX_Click` (7) — llama al helper `EliminarX`, refresca visual.
- `ComandoCopiarX_Click` (2) — `CopiarAlPortapapeles` + `MsgBox` confirmación.
- `ComandoLimpiarX_Click` (2) — `Me.X = Null` puro UI.
- `ListaResponsables_Click` — enable/disable de `ComandoCambiarJP`/`ComandoCambiarEnvios`.
- `Form_Load` — orquesta inicialización.

### 7.3 Public surface preservada (per e2e-methodology §6)

Los 10 Public methods deben MANTENER firma:
- `RellenarListaPECAL`, `RellenarListaLugares`, `RellenarListaComerciales`, `RellenarListaCPVs`, `RellenarListaRACs`, `RellenarListaResponsables`, `RellenarListaAnualidades`, `RellenarListas`, `EstablecerDatos`, `CambiarLineaListaResponsables`.
- El cuerpo se reescribe: en vez de `Me.ListaX.AddItem` (UI), llama al helper y aplica el resultado al listbox (JSON → RowSource / AddItem).

## 8. NOTAS Y RIESGOS

1. **`pregunta = MsgBox(...)` no compila con Option Explicit** — bug pre-existente. La forma thinned debe corregirlo (quitar la asignación o declarar `Dim pregunta As VbMsgBoxResult`).
2. **Constructor lookup** — el helper llama a `constructor.getComercial/getCPV/etc.`. Esos métodos tienen su propio `p_Error ByRef`. El helper propaga el error.
3. **DAO `RegistrarX` retorna estado "1" o "0"** — el helper trata "1" como éxito y refresca la colección. Estado ≠ "1" se considera no-registrado.
4. **`Expediente.ColX`** — `ColComerciales`, `ColCPVs`, etc. son propiedades de Expediente que retornan Dictionary. El helper las llama después de cada alta/baja para refrescar.
5. **`m_ObjExpedienteDTOActivo`** — global del proyecto. El form wrappea en `{Expediente, ColComerciales, ...}` antes de pasar al helper. Tests inyectan Dictionary stub directamente.
6. **`WithEvents` subscribers** — no se pueden extraer a helpers (es UI event handling). Quedan en form y delegan al helper.
7. **`CambiarLineaListaResponsables`** — la lógica pura (string replace) se extrae a un helper que recibe el RowSource actual + línea a buscar + línea nueva, y retorna el RowSource final. El form hace la asignación al listbox.

## 9. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `DoCmd.*` / `Forms(...)` / `Application.Echo` / `Me.X` | OK esperado |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` | OK esperado |
| Helpers: ZERO `On Error GoTo` + `Err.Raise 1000` pattern (use `p_Error`) | OK esperado |
| Helpers: ZERO MsgBox | OK esperado |
| `pregunta = MsgBox(...)` bug fixed in thinned form | OK esperado |
| Public method signatures preserved (10 methods) | OK esperado |
| WithEvents declarations preserved (6) | OK esperado |
| Atoms estimated | ~50-80 (10 helpers × 5-8 atoms avg) |

## 10. REWORK (2026-06-26, PR-9 / Phase 3.3)

> Status: **PENDING** — helper + test modules to be written, then imported via dysflow, then form .cls thinned.

### 10.1 Files to add

| File | Purpose |
|---|---|
| `src/modules/modExpedienteEntidadesHelper.bas` | Pure-data helpers (~25 Public Functions + private utilities) |
| `src/modules/Test_ExpedienteEntidadesHelper.bas` | TDD atoms (~50-80 Public atoms + RunAll) |

### 10.2 Form .cls changes

| File | Before | After (est.) | Delta |
|---|---|---|---|
| `src/forms/Form_FormExpedienteEntidades.cls` | 1802 | ~600-800 | -50% approx |

**Public surface preserved** (per e2e-methodology §6):
- 10 Public Functions — signature unchanged, body now wraps DTO → Dictionary → helper → applies JSON to controls.
- 6 WithEvents declarations — unchanged.
- 6 WithEvents subscribers — body now delegates to helper.

**Removed**:
- `On Error GoTo errores + Err.Raise 1000` pattern in handlers (replaced by helper error returns).
- `VBA.DoEvents / DoCmd.Hourglass True/False` blocks in handlers (kept at handler level for UI orchestration, removed from helper).
- `pregunta = MsgBox(...)` bug fixed.

**Added**:
- Dictionary wrapper construction in handlers.
- Helper delegation pattern: handler → wrap DTO → helper → apply result.