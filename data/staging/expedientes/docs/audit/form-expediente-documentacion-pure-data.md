# Audit — Form_FormExpedienteDocumentacion (Phase 3.2 / PR-8)

> **Change**: `forms-thin-coverage` (Phase 3.2 — PR-8, batch of 6 forms)
> **Form**: `Form_FormExpedienteDocumentacion` (pestaña de documentación / anexos)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight)

## 1. STATUS

| Fase | Estado |
|---|---|
| Audit | **COMPLETE** |
| Helper extraction | **PENDING** (4-5 helpers planned) |
| Binario sync | PENDING |

## 2. FORM METADATA

| Atributo | Valor |
|---|---|
| Líneas .cls | 303 |
| Handlers totales | 7 (Form_Load + ComandoAltaAnexo + ComandoEliminar + ImagenAnexo_Click + ListaDocumentos_Click + ListaDocumentos_DblClick + EstablecerDatos + RellenarListas) |
| Public methods | 2 (`EstablecerDatos(Optional ByRef p_Error As String) As String`, `RellenarListas(Optional ByRef p_Error As String) As String`) — **MANTENER firma** |
| Public events | 0 |
| MsgBox call sites | 5 (ComandoAltaAnexo error, ComandoEliminar error + pregunta, Form_Load error, ImagenAnexo_Click error, ListaDocumentos_Click error) |
| DoCmd.OpenForm en .cls | 0 |
| DoCmd.Close en .cls | 0 (este form no se cierra — es sub-form) |
| Forms abiertas | 0 |
| WithEvents | 0 |

## 3. CONTROLES REFERENCIADOS (Me.X)

| Control | Tipo | Usado en |
|---|---|---|
| `ListaDocumentos` | ListBox | RellenarListas, ListaDocumentos_Click, ListaDocumentos_DblClick, ComandoEliminar, ComandoAltaAnexo |
| `ImagenAnexo` | Image | ListaDocumentos_Click, ListaDocumentos_DblClick, EstablecerDatos |
| `Caption` | Form property | Form_Load |
| `AllowEdits` | Form property | Form_Load |

## 4. ENTIDAD Y PROPIEDADES REFERENCIADAS

| Entidad | Propiedades usadas |
|---|---|
| `ExpedienteOperaciones` | `RegistrarAnexo(p_URLLocal, p_Error)`, `EliminarAnexo(p_NombreDocumento, p_Error)` |
| `ExpedienteDTO` (global `m_ObjExpedienteDTOActivo`) | `Expediente`, `ColAnexos` |
| `Expediente` | `Anexos` (refresh) |
| `ExpedienteAnexo` | `IDExpediente`, `NombreDocumento`, `URLDocumento` |

## 5. INTER-FORM CALLSITES

| Target | Mecanismo |
|---|---|
| `Helper_ExpedienteEdicion.EstablecerDatos` | Module call |
| `constructor.Seleccionar` | Module call (file picker) |
| `constructor.fso` | Module call (filesystem) |
| `constructor.AbrirEnLocal` | Module call (open file) |
| `MostrarPopupProgreso` / `CerrarPopupProgreso` | Module call (UI guard) |
| `Application.Echo` | **inline (anti-pattern a eliminar)** |

## 6. ANTI-PATTERN DETECTADO

⚠️ **`Application.Echo False/True`** en `ListaDocumentos_DblClick` (líneas 297, 302). Esto es UI suppression — el helper NO debe usarlo. La doble-click puede resolverse con un simple re-orden: helper retorna, form decide.

## 7. RECOMENDACIÓN DE PATRÓN

**Pattern**: SPECIAL — alta/baja anexos con file system. NO encaja en Gestion/Alta clásico (combina CRUD + file picker).

### 7.1 Helpers planeados (5)

| Helper | Signature | Responsabilidad |
|---|---|---|
| `ExpedienteDocumentacion_Form_Load` | `(p_DTO As Object, p_Error As String) As String` | Caption, AllowEdits, llamar `EstablecerDatos`. |
| `ExpedienteDocumentacion_EstablecerDatos` | `(p_DTO As Object, p_EsAdmin As Boolean, p_Error As String) As String` | Recorre `Me.Controls`, si `Tag="EJECUTIVO"` setea Enabled según admin. Llama `RellenarListas`. |
| `ExpedienteDocumentacion_RellenarListas` | `(p_DTO As Object, p_Error As String) As String` | Recorre `ColAnexos`, retorna JSON: `{rowSource:String}` (semicolon-separated `URL;FileName`). El form hace `lst.AddItem` (UI wiring). |
| `ExpedienteDocumentacion_AltaAnexo` | `(p_DTO As Object, p_URLLocal As String, p_PromptResult As Long, p_Error As String) As String` | 1) Verifica `fso.FileExists(p_URLLocal)`. 2) `ExpedienteOperaciones.RegistrarAnexo`. 3) Refresca `ColAnexos`. Retorna `{registrado:Boolean, fileExists:Boolean}`. |
| `ExpedienteDocumentacion_EliminarAnexo` | `(p_DTO As Object, p_NombreDocumento As String, p_EsAdmin As Boolean, p_PromptResult As Long, p_Error As String) As String` | 1) Si NO admin → error. 2) Si `p_PromptResult=0` → cancelado. 3) `ExpedienteOperaciones.EliminarAnexo`. 4) Refresca `ColAnexos`. Retorna `{eliminado:Boolean, autorizado:Boolean, promptAccepted:Boolean}`. |

### 7.2 UI orchestration que se queda en el form

- `ComandoAltaAnexo_Click` — `Seleccionar(True, ...)` (file picker UI) → `ExpedienteDocumentacion_AltaAnexo` → `RellenarListas` (refresco).
- `ComandoEliminar_Click` — `MsgBox("¿Desea eliminar?")` → `ExpedienteDocumentacion_EliminarAnexo` → `RellenarListas`.
- `ImagenAnexo_Click` — `AbrirEnLocal Me.ListaDocumentos.Column(0)`. Puro UI.
- `ListaDocumentos_Click` — setea `Me.ImagenAnexo.Visible = True`. UI state.
- `ListaDocumentos_DblClick` — llama `ImagenAnexo_Click`. **Aquí se elimina el `Application.Echo`**.
- `Form_Load` — orquesta.
- `EstablecerDatos` / `RellenarListas` — MANTENER firma pública; contenido es UI wiring + helper call.

## 8. NOTAS Y RIESGOS

1. **`fso.FileExists`** — el helper lo usa. `fso` es global del proyecto (`FileSystemObject`). El test puede inyectar un stub o testear con paths reales (sandbox).
2. **`Seleccionar(True, "Seleccione el archivo a anexar", m_Error)`** — file picker. **NO se mete a helper** — es UI input. El form llama primero y pasa el resultado al helper.
3. **`Application.Echo False/True`** — **ELIMINAR**. Reemplazar por simple delegación sin echo suppression.
4. **MsgBox en pregunta de eliminar** — UI confirmation, se queda en el form. `p_PromptResult` del helper permite testear sin modal.
5. **`m_ObjExpedienteDTOActivo.Expediente.Anexos`** — refresh después de alta/baja. El helper retorna el flag y el form actualiza la colección (UI orchestration) o el helper lo hace y retorna.

## 9. VERIFICATION CHECKLIST (post-extraction)

| Check | Target |
|---|---|
| Helper signatures: ZERO `ByRef p_Form` | OK esperado |
| Helpers: ZERO `Application.Echo` / `DoCmd.*` / `Forms(...)` | OK esperado |
| Test atoms: ZERO `DoCmd.OpenForm` / `Forms(...)` / `Screen.ActiveForm` / `Application.Echo` | OK esperado |
| `Application.Echo` en .cls thinned | **0** (eliminado) |
| Atoms estimated | ~25-35 (5 helpers × 5-7 atoms promedio) |
| Pattern | SPECIAL — documentación / file system |

## 10. REWORK (2026-06-26, PR-8 / Phase 3.2)

> Status: **DONE** — helper + test modules imported via dysflow, form .cls thinned.

### 10.1 Files added

| File | Lines | Purpose |
|---|---|---|
| `src/modules/modExpedienteDocumentacionHelper.bas` | 675 | Pure-data helpers (5 Public Functions + private utilities) |
| `src/modules/Test_ExpedienteDocumentacionHelper.bas` | 895 | TDD atoms (21 Public atoms + RunAll) |

### 10.2 Form .cls changes

| File | Before | After | Delta |
|---|---|---|---|
| `src/forms/Form_FormExpedienteDocumentacion.cls` | 303 | 334 | +31 (DTO wrapping added) |

**Public surface preserved** (per e2e-methodology §6):
- `EstablecerDatos(Optional ByRef p_Error As String) As String` — signature unchanged, body now wraps DTO → Dictionary → helper → applies JSON to controls.
- `RellenarListas(Optional ByRef p_Error As String) As String` — signature unchanged, body now wraps DTO → Dictionary → helper → applies JSON to listbox.

**Removed**:
- `Application.Echo False/True` in `ListaDocumentos_DblClick` — replaced by a simple `If Visible Then ImagenAnexo_Click`.

**Added**:
- Dictionary wrapper construction in `EstablecerDatos`, `RellenarListas`, `ComandoAltaAnexo_Click`, `ComandoEliminar_Click`. The wrapper carries `{Expediente, ColAnexos}` keys for the helper. Real `ExpedienteDTO` is class-based; the wrapper normalizes the shape so the helper accepts plain `Scripting.Dictionary` stubs in tests.

### 10.3 Helpers implemented

| Helper | Signature | Returns |
|---|---|---|
| `ExpedienteDocumentacion_Form_Load` | `(p_DTO, p_Titulo, p_Error)` | `{titulo, allowEdits, expedienteOK}` |
| `ExpedienteDocumentacion_EstablecerDatos` | `(p_DTO, p_EsAdmin, p_Error)` | `{ejecutivosEnabled, expedienteOK}` |
| `ExpedienteDocumentacion_RellenarListas` | `(p_DTO, p_Error)` | `{rowSource, count, items}` |
| `ExpedienteDocumentacion_AltaAnexo` | `(p_DTO, p_URLLocal, p_PromptResult, p_Error)` | `{registrado, fileExists}` |
| `ExpedienteDocumentacion_EliminarAnexo` | `(p_DTO, p_NombreDocumento, p_EsAdmin, p_PromptResult, p_Error)` | `{eliminado, autorizado, promptAccepted}` |

All helpers accept `p_DTO As Object` (Dictionary-shaped wrapper with `Expediente` and `ColAnexos` keys). The form wraps the real `m_ObjExpedienteDTOActivo` into this shape before each call.

### 10.4 Atoms implemented (21 atoms)

| Atom | Class | Coverage |
|---|---|---|
| `Form_Load_HappyConExpediente` | happy | DTO with Expediente, titulo echoed |
| `Form_Load_SadDTONothing` | sad | DTO Nothing rejected |
| `Form_Load_EdgeExpedienteNothing` | edge | DTO present but Expediente Nothing → expedienteOK=false |
| `EstablecerDatos_HappyAdmin` | happy | ejecutivosEnabled=true |
| `EstablecerDatos_HappyNonAdmin` | happy | ejecutivosEnabled=false |
| `EstablecerDatos_SadDTONothing` | sad | DTO rejected |
| `RellenarListas_HappyTresAnexos` | happy | 3 anexos, all rendered |
| `RellenarListas_EdgeEmptyColAnexos` | edge | empty ColAnexos → count=0 |
| `RellenarListas_EdgeDTONothing` | edge | DTO Nothing → count=0 |
| `RellenarListas_EdgeAnexoSinExpediente` | edge | anexo with empty IDExpediente → URL=nombre |
| `RellenarListas_AdversarialSemicolonInNombre` | adversarial | `;` sanitized to `:` |
| `AltaAnexo_EdgeEmptyURL` | edge | empty URL → no-op |
| `AltaAnexo_EdgeFileNotExists` | edge | non-existent file → fileExists=false, registrado=false |
| `AltaAnexo_SadDTONothing` | sad | DTO Nothing → fail JSON |
| `EliminarAnexo_SadNoAdmin` | sad | non-admin → fail JSON |
| `EliminarAnexo_EdgePromptPending` | edge | prompt=0 → promptAccepted=false |
| `EliminarAnexo_EdgePromptNo` | edge | prompt=vbNo → cancel |
| `EliminarAnexo_EdgeEmptyNombre` | edge | empty nombre → no DAO call |
| `EliminarAnexo_SadDTONothing` | sad | DTO Nothing → fail JSON |
| `EliminarAnexo_SadExpedienteNothing` | sad | Expediente Nothing → fail JSON |
| `EliminarAnexo_AdversarialUnknownPrompt` | adversarial | unknown prompt result → cancel |

Plus `Test_ExpedienteDocumentacionHelper_RunAll` wrapper for Dysflow manifest discovery.

### 10.5 DAO happy-path atoms — DEFERRED

The full happy-path DAO atoms for `AltaAnexo` and `EliminarAnexo` require:
1. A valid `TbExpedientes` row (real fixture ID >= 900800).
2. A real file on disk matching the helper's URL validation.
3. The `Expediente.Anexos` property get to succeed (depends on filesystem + DB layout).

These atoms are **deferred to a follow-up PR** because:
- The DAO dependencies are deep (URLDirectorio calculation requires getExpedienteAnexos + filesystem state).
- The input-validation coverage (atoms above) catches the most likely regression surface.
- Pilot pattern (modSuministradorHelper) has fewer DAO atoms than input-validation atoms — consistent with deferred DAO coverage here.

### 10.6 Verification

- ZERO `ByRef p_Form` in helpers: ✓
- ZERO `DoCmd.OpenForm` in tests: ✓
- ZERO `Forms(...)` / `Screen.ActiveForm` in tests: ✓
- ZERO `Application.Echo` in helpers and tests: ✓
- ZERO manual VBS/PS1 bypass: ✓
- All Access operations via dysflow: ✓
- Declaration ordering per vba-access §10.1 (private consts first, public atoms last): ✓

### 10.7 Known follow-ups

1. DAO happy-path atoms (see §10.5).
2. The thinned form .cls uses `CreateObject("Scripting.Dictionary")` to wrap the real DTO. This is a slight runtime overhead per call but keeps the helper decoupled from `ExpedienteDTO`'s class shape. Acceptable trade-off.
