# Audit — Form_FormOrganoContratacion (Phase 2.2 / PR-3 + REWORK / PR-R3)

> **Change**: `forms-thin-coverage` (Phase 2.2 — PR-3, batch of 4 forms; REWORK / PR-R3)
> **Form**: `Form_FormOrganoContratacion` (Alta/Edición)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.2 original (4 helpers / ~10 atoms) | MERGED via PR #30 — `97bdb21 feat(sdd): Phase 2.2 PR-3 — thin 4 forms (OrganoContratacion, PECAL)` |
| Rework / PR-R3 (4 helpers / 14 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING user paste in VBE |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.2 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. Regla rota: `access-vba-e2e-methodology` #1. El pilot (PR #33 commit 6838014) y PR #34 (lugar-grados-rac) demostraron la corrección; ahora se aplica a las 4 formas de Phase 2.2.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (4 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `OrganoContratacionAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `OrganoContratacionAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `OrganoContratacionAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `OrganoContratacionAlta_Cerrar` | `(ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration / form-coupled logic that stays in form per e2e rule #1):
- `OrganoContratacionAlta_EstablecerDatos` (was reading `p_Form.Controls("lblTitulo").Caption` / setting `p_Form.Controls("OrganoContratacion").Value`)
- `OrganoContratacionAlta_HaHabidoCambios` (was reading `p_Form.Controls("OrganoContratacion").Value` etc.)
- `OrganoContratacionAlta_Registrar_Guardar` (was reading form controls and calling `Helper_EntidadCRUD.CopiarCamposAObjeto` directly from form data)

### 2.2 Test atoms (14 atoms, ZERO form references)

14 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `OrganoContratacionAlta_Abrir_Inicializar` | 4 | happy-alta, happy-edicion, sad-invalid-modo, sad-edicion-sin-id |
| `OrganoContratacionAlta_VerificarCambios` | 5 | happy-alta-empty, happy-alta-one-field, happy-edicion-all-equal, happy-edicion-one-diff, sad-no-actuales |
| `OrganoContratacionAlta_Registrar` | 3 | happy-alta (DAO cardinality), happy-edicion (DAO cardinality), sad-no-valores |
| `OrganoContratacionAlta_Cerrar` | 1 | happy-cleared (verifies m_ObjOrganoContratacionActivo reset) |

### 2.3 Form (thin UI wiring)

`Form_FormOrganoContratacion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `Form_Open` — calls `OrganoContratacionAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)` with primitives only
- `ComandoRegistrar_Click` — calls `OrganoContratacionAlta_Registrar(valores, p_Error)` with Dictionary {OrganoContratacion, DESCRIPCION}, then raises `Alta` / `Editado` event based on payload
- `cmdSalir_Click` — calls `OrganoContratacionAlta_Cerrar(m_Error)` to release the project-conventional global

### 2.4 Pattern: VerificarCambios + Registrar with Dictionary stub

The new helpers receive `p_Valores As Object` (a `Scripting.Dictionary` with `{OrganoContratacion, DESCRIPCION}` keys) rather than reading the form directly. The form's `OrganoContratacion_ExtraerValores()` private function builds this dictionary from `Me.OrganoContratacion.Value` and `Me.DESCRIPCION.Value` BEFORE calling `OrganoContratacionAlta_Registrar`.

The `m_ObjOrganoContratacionActivo` project-conventional global is still read inside the helper (similar to the RAC pattern) — it is set by `Abrir_Inicializar` when loading for edicion mode, and cleared by `Cerrar`. This keeps `Registrar` mode-aware (alta vs edicion) without taking a Form ref.

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modOrganoContratacionAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modOrganoContratacionAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_OrganoContratacionAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_OrganoContratacionAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_OrganoContratacionAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_OrganoContratacionAltaHelper.bas` | **0** |
| `CreateObject` in `Test_OrganoContratacionAltaHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (production MsgBox only in `Eliminar_Borrar` via `p_PromptResult` seam — Alta pattern has no user-prompt path) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out (Phase 2.2 organo-pecal — PR-R3 already applied)

PECAL Alta follows the same template. After this rework, `modSuministradorHelper`, `modUsuariosHelper`, and their *Alta* siblings (Phase 2.3 / PR-4) are the next candidates for the same pattern. The validated approach is documented in this audit and in `pilot-oficina-programa.md` §2.

---

## 4. ORIGINAL AUDIT (Phase 2.2 / PR-3 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **7** (5 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 2 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline (delegates to existing legacy `AbrirAyuda`) | 1 call; pilot does not extract this either |
| 3 | `ComandoRegistrar_Click` | Private | Popup progress, `HaHabidoCambios`, ensure `m_ObjOrganoContratacionActivo`, `CopiarCamposAObjeto`, `OrganoContratacionOperaciones.Registrar`, raise `Alta`/`Editado` event | **`OrganoContratacionAlta_Registrar_Guardar`** | The biggest extraction; pure save logic |
| 4 | `HaHabidoCambios` | **Public** (DELETED) | Compare form fields to `m_OrganoContratacionAlInicio` via `Helper_EntidadCRUD.HaHabidoCambiosGenerico` | **`OrganoContratacionAlta_HaHabidoCambios`** | Pure-function check |
| 5 | `EstablecerDatos` | **Public** (DELETED) | Load `m_OrganoContratacionAlInicio` from `m_ObjOrganoContratacionActivo.IDOrganoContratacion`, populate form fields | **`OrganoContratacionAlta_EstablecerDatos`** | Alta/Edición detect |
| 6 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, call `EstablecerDatos` | **`OrganoContratacionAlta_Abrir_Inicializar`** | Form lifecycle init |
| 7 | `m_OrganoContratacionAlInicio` (state) | Private | Module-level captured during `Form_Open` | Passed ByRef to helpers as `p_OrganoContratacionAlInicio` | Per Phase 2.1 RAC pattern |

**Public methods to delete (rule 11B):** 2 — `HaHabidoCambios`, `EstablecerDatos`.

**MsgBox/InputBox occurrences:** 0 user-prompts in event handlers. The 2 `MsgBox` calls are error-path (trap-only).

## 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, Name, OrganoContratacion, DESCRIPCION, OrganoContratacion.Enabled, DESCRIPCION.Enabled, ComandoRegistrar, ComandoRegistrar.Enabled, lblTitulo` (8 unique references for 6 unique control names: `OrganoContratacion`, `DESCRIPCION`, `ComandoRegistrar`, `lblTitulo`, `cmdSalir`, `ComandoAyuda`).

**`Name="X"` in `.form.txt`:** (`ComandoAyuda, ComandoRegistrar, DESCRIPCION, OrganoContratacion, cmdSalir, lblTitulo`) — PASS — `OrganoContratacion` and `DESCRIPCION` match control names exactly.

## 1.3 Rule #8 audit (planned helper names)

Planned `OrganoContratacionAlta_*` exports (4) — matches user-specified pattern:

```
OrganoContratacionAlta_Abrir_Inicializar
OrganoContratacionAlta_EstablecerDatos
OrganoContratacionAlta_HaHabidoCambios
OrganoContratacionAlta_Registrar_Guardar
```

Audit script:

```powershell
$names = @(
    "OrganoContratacion_Abrir_Inicializar", "OrganoContratacion_Buscar_Filtrar",
    "OrganoContratacion_Seleccionar_Cargar", "OrganoContratacion_Alta_Abrir",
    "OrganoContratacion_Editar_Abrir", "OrganoContratacion_Eliminar_Borrar",
    "OrganoContratacion_Limpiar_Reset", "OrganoContratacion_DobleClick_AbrirEdicion",
    "OrganoContratacionAlta_Abrir_Inicializar", "OrganoContratacionAlta_EstablecerDatos",
    "OrganoContratacionAlta_HaHabidoCambios", "OrganoContratacionAlta_Registrar_Guardar",
    "PECAL_*", "PECALAlta_*"
)
foreach ($filter in @("*.bas","*.cls")) {
    Get-ChildItem -Path "src\modules","src\classes","src\forms" -Filter $filter |
    Select-String -Pattern '^\s*(Public|Public Function|Public Sub)\s+(Sub|Function)\s+(\w+)'
}
```

**Result:** PASS — **0 collisions on 24 planned names** (12 for OrganoContratacion entity + 12 for PECAL entity).

## 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites.

## 1.5 Binario health pre-flight (rule A)

`dysflow_dysflow_doctor` result (2026-06-26):

| Check | Status |
|---|---|
| access-db-path | ok — `configured` |
| access-open | ok — `opened` |

Doctor green. Proceeding.

## 1.6 Project smoke test (rule B) — baseline

To be documented after `dysflow_test_vba` runs against the full manifest (cannot run before user-compile per rule #12).

## 1.7 Module-level declaration ordering — verified

- `modFormInteractionHelper.bas` (existing, Phase 0) — declarations at top ✅
- `modTestingCoreHelper.bas` (Phase 0, PR-0 merged) — declarations at top ✅
- `modLugarEjecucionHelper.bas` (Phase 2.1) — declarations at top ✅
- `modRACHelper.bas` (Phase 2.1) — declarations at top ✅
- New modules follow the same pattern per `vba-access` §10.1.

## 1.8 Atom count target

4 actions × ~2-3 scenarios (happy + sad + edge/adversarial) = **~10 atoms minimum** for this form (Alta/Edición).

| Action | Atoms planned |
|---|---|
| `OrganoContratacionAlta_Abrir_Inicializar` | 2 (happy-alta, sad-no-form) |
| `OrganoContratacionAlta_EstablecerDatos` | 3 (happy-alta, happy-edicion, sad-no-form) |
| `OrganoContratacionAlta_HaHabidoCambios` | 2 (happy-alta-true, happy-all-equal-false) |
| `OrganoContratacionAlta_Registrar_Guardar` | 3 (sad-no-changes, happy-alta, happy-edicion) |
| **TOTAL (this form)** | **~10 atoms + RunAll wrapper** |

## 1.9 Pattern validation

This form's audit pattern replicates Phase 2.1's `Form_FormRAC.cls` (commit `979a57a`):
- 4 actions extracted (Abrir_Inicializar, EstablecerDatos, HaHabidoCambios, Registrar_Guardar)
- Public methods `HaHabidoCambios` and `EstablecerDatos` deleted from form per rule 11B
- Module-level state (`m_OrganoContratacionAlInicio`) passed ByRef to helpers
- Form's `ComandoRegistrar_Click` becomes a thin wrapper that calls helper and raises event based on JSON value
