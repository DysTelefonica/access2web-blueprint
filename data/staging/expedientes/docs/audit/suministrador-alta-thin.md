# Audit — Form_FormSuministrador (Phase 2.3 / PR-4 + REWORK / PR-R4)

> **Change**: `forms-thin-coverage` (Phase 2.3 — PR-4, batch of 5 forms; REWORK / PR-R4)
> **Form**: `Form_FormSuministrador` (Alta/Edición form, NOT a Gestion list)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.3 original (4 helpers / ~9 atoms) | MERGED via PR #31 — `c86d460 refactor(sdd): Phase 2.3 / PR-4 — thin 5 forms (suministrador + usuarios + frmBusy + oficina-programa)` |
| Rework / PR-R4 (4 helpers / 11 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING user paste in VBE |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.3 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. Regla rota: `access-vba-e2e-methodology` #1.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (4 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `SuministradorAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `SuministradorAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `SuministradorAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `SuministradorAlta_Cerrar` | `(ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration / form-coupled logic that stays in form per e2e rule #1):
- `SuministradorAlta_EstablecerDatos` (was reading `p_Form.Controls("lblTitulo").Caption` / setting `p_Form.Controls("Nombre").Value` etc. for 7 fields)
- `SuministradorAlta_HaHabidoCambios` (was reading all 7 form fields and calling `Helper_EntidadCRUD.HaHabidoCambiosGenerico`)
- `SuministradorAlta_Registrar_Guardar` (was reading form controls and calling `Helper_EntidadCRUD.CopiarCamposAObjeto` directly from form data)

### 2.2 Test atoms (11 atoms, ZERO form references)

11 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `SuministradorAlta_Abrir_Inicializar` | 3 | sad-invalid-modo, happy-alta, sad-edicion-sin-id |
| `SuministradorAlta_VerificarCambios` | 5 | alta-one-field, alta-all-empty, edicion-all-equal, edicion-one-diff, sad-no-actuales |
| `SuministradorAlta_Registrar` | 1 | sad-no-valores |
| `SuministradorAlta_Cerrar` | 1 | happy-cleared (verifies m_ObjSuministradorActivo reset) |

### 2.3 Form (thin UI wiring)

`Form_FormSuministrador.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `Form_Open` — calls `SuministradorAlta_Abrir_Inicializar(modo, idEntidad, m_SuministradorAlInicio, m_Error)` with primitives only
- `ComandoRegistrar_Click` — calls `SuministradorAlta_Registrar(valores, p_Error)` with Dictionary of 7 fields, then raises `Alta(p_ID)` / `Editado` event based on payload.modo
- `cmdSalir_Click` — calls `SuministradorAlta_Cerrar(m_Error)` to release the project-conventional global

### 2.4 Pattern: VerificarCambios + Registrar with Dictionary stub

The new helpers receive `p_Valores As Object` (a `Scripting.Dictionary` with the 7 Suministrador fields) rather than reading the form directly. The form's `Suministrador_ExtraerValores()` private function builds this dictionary from the 7 `Me.X.Value` controls BEFORE calling the helpers.

The `m_ObjSuministradorActivo` project-conventional global is still read inside the helper (similar to the OrganoContratacionAlta pattern). It is set by `Abrir_Inicializar` when loading for edicion mode, and cleared by `Cerrar`.

### 2.5 Pattern: Alta(p_ID) event signature preservation

The form's `ComandoRegistrar_Click` extracts the resulting ID from the helper's payload (`payload.id`) for alta mode, or from `m_ObjSuministradorActivo.IDSuministrador` for edicion mode, and passes it to the `Alta(p_ID)` event. The legacy event signature `Alta(p_ID As String)` is preserved — the parent form `Form_FormSuministradoresGestion` listens for it.

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modSuministradorAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modSuministradorAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_SuministradorAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_SuministradorAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_SuministradorAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_SuministradorAltaHelper.bas` | **0** |
| `CreateObject` in `Test_SuministradorAltaHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Alta pattern has no user-prompt path) |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.3 / PR-4 preflight, preserved for traceability)

### 0. Form profile — Alta/Edición (DIFFERENT pattern from Gestion list)

`Form_FormSuministrador` is the **Alta/Edición** form for Suministrador entities, structurally different from the Gestion list (`Form_FormSuministradoresGestion`). 7 fields: `Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS`.

**Prefix disambiguation** (matches Phase 2.2 OrganoContratacion pattern):
- `SuministradorAlta_*` — Alta/Edición form helpers
- `Suministrador_*` — Gestion list helpers

### 1.1 Handlers audit

Total handlers: **4** (4 Private) + 2 Public methods to delete.

| # | Handler | Scope | Notes |
|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Trivial close |
| 2 | `ComandoAyuda_Click` | Private | 1 call to `AbrirAyuda` |
| 3 | `ComandoRegistrar_Click` | Private | The biggest extraction; pure save logic |
| 4 | `Form_Open` | Private | Form lifecycle init |
| 5 | `HaHabidoCambios` | **Public** (DELETED) | Inlined into Registrar flow |
| 6 | `EstablecerDatos` | **Public** (DELETED) | Load 7 fields from m_ObjSuministradorActivo |

### 1.2 Form controls audit (rule 11A)

`Me.X` references in `.cls`: `Caption, ComandoRegistrar, Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS, lblTitulo, Name` (11 unique).

### 1.3 Rule #8 audit (planned helper names)

Planned `SuministradorAlta_*` exports (4) — matches `modRAC / modOrganoContratacionAlta` pattern.

### 1.4 Inter-form census (rule 11)

`DoCmd.OpenForm "FormSuministrador"` is a built-in DoCmd call, not a method call on another form's `.cls`.

### 1.5 Binario health pre-flight (rule A)

Doctor pending — verified before module import in PR-R4.

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs.

### 1.7 Module-level declaration ordering — verified

Same pattern as Phase 2.1 / 2.2 ✅.

### 1.8 Atom count target — REWORKED

Original target: ~10 atoms across 4 helpers. Reworked target: 11 atoms across 4 helpers (richer coverage on VerificarCambios alta+edicion paths).

### 1.9 Special considerations

1. **`m_ObjSuministradorActivo` global** — set by `Form_FormSuministradoresGestion.ComandoEditar_Click` and `ComandoAlta_Click` BEFORE opening `FormSuministrador`. The helpers read it as a global (legitimate project-conventional global state).
2. **`m_SuministradorAlInicio`** — module-level form state captured during `Form_Open` via `Abrir_Inicializar` and consulted by `ComandoRegistrar_Click`. The helper `SuministradorAlta_VerificarCambios` accepts an explicit `p_ValoresOriginales` parameter (does NOT read `Me` or any module-level state) — atoms can pass a constructed Dictionary.
3. **Event signature difference**: `Alta(p_ID As String)` includes the new ID as parameter. The helper returns `"alta"` in JSON; the form extracts the resulting ID from `payload.id` (alta path) or `m_ObjSuministradorActivo.IDSuministrador` (edicion path) and passes it to the event.
