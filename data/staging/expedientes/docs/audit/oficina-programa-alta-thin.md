# Audit — Form_FormOficinaPrograma (Phase 2.3 / PR-4 + REWORK / PR-R4)

> **Change**: `forms-thin-coverage` (Phase 2.3 — PR-4, batch of 5 forms; REWORK / PR-R4)
> **Form**: `Form_FormOficinaPrograma` (Alta/Edición form, NOT the Gestion list from pilot)
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
| `OficinaProgramaAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `OficinaProgramaAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `OficinaProgramaAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `OficinaProgramaAlta_Cerrar` | `(ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration / form-coupled logic that stays in form per e2e rule #1):
- `OficinaProgramaAlta_EstablecerDatos` (was reading `p_Form.Controls("lblTitulo").Caption` / setting `p_Form.Controls("OficinaPrograma").Value` / `p_Form.Controls("DESCRIPCION").Value`)
- `OficinaProgramaAlta_HaHabidoCambios` (was reading both form fields and calling `Helper_EntidadCRUD.HaHabidoCambiosGenerico`)
- `OficinaProgramaAlta_Registrar_Guardar` (was reading form controls and calling `Helper_EntidadCRUD.CopiarCamposAObjeto` directly from form data)

### 2.2 Test atoms (11 atoms, ZERO form references)

11 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `OficinaProgramaAlta_Abrir_Inicializar` | 3 | sad-invalid-modo, happy-alta, sad-edicion-sin-id |
| `OficinaProgramaAlta_VerificarCambios` | 5 | alta-one-field, alta-all-empty, edicion-all-equal, edicion-one-diff, sad-no-actuales |
| `OficinaProgramaAlta_Registrar` | 1 | sad-no-valores |
| `OficinaProgramaAlta_Cerrar` | 1 | happy-cleared (verifies m_ObjOficinaProgramaActiva reset) |

### 2.3 Form (thin UI wiring)

`Form_FormOficinaPrograma.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `Form_Open` — calls `OficinaProgramaAlta_Abrir_Inicializar(modo, idEntidad, m_OficinaProgramaAlInicio, m_Error)` with primitives only
- `ComandoRegistrar_Click` — calls `OficinaProgramaAlta_Registrar(valores, p_Error)` with Dictionary `{OficinaPrograma, DESCRIPCION}`, then raises `Alta` / `Editado` event based on payload.modo
- `cmdSalir_Click` — calls `OficinaProgramaAlta_Cerrar(m_Error)` to release the project-conventional global

### 2.4 Pattern: VerificarCambios + Registrar with Dictionary stub

The new helpers receive `p_Valores As Object` (a `Scripting.Dictionary` with `{OficinaPrograma, DESCRIPCION}` keys) rather than reading the form directly. The form's `OficinaPrograma_ExtraerValores()` private function builds this dictionary from `Me.OficinaPrograma.Value` and `Me.DESCRIPCION.Value` BEFORE calling the helper.

The `m_ObjOficinaProgramaActiva` project-conventional global is still read inside the helper (similar to the OrganoContratacionAlta pattern). It is set by `Abrir_Inicializar` when loading for edicion mode, and cleared by `Cerrar`.

### 2.5 Important: DIFFERENT from pilot's Gestion list

**This is `Form_FormOficinaPrograma` (Alta/Edición form)**, NOT `Form_FormOficinasProgramaGestion` (Gestion list, covered in PR #33 pilot commit 6838014 with `modOficinaProgramaHelper` + prefix `OficinaPrograma_*`). The pilot already extracted the Gestion list helper. This PR covers the Alta/Edición form using the `modOrganoContratacionAltaHelper` pattern (4 helpers with prefix `OficinaProgramaAlta_*`).

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modOficinaProgramaAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modOficinaProgramaAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_OficinaProgramaAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_OficinaProgramaAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_OficinaProgramaAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_OficinaProgramaAltaHelper.bas` | **0** |
| `CreateObject` in `Test_OficinaProgramaAltaHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.3 / PR-4 preflight, preserved for traceability)

### 0. Form profile — Alta/Edición (DIFFERENT form from the pilot's Gestion list)

Form controls: `OficinaPrograma`, `DESCRIPCION`, `lblTitulo`, `ComandoRegistrar`.

### 1.1 Handlers audit

Total handlers: **4** (4 Private) + 2 Public methods to delete.

| # | Handler | Scope | Notes |
|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Trivial close |
| 2 | `ComandoAyuda_Click` | Private | 1 call to `AbrirAyuda` |
| 3 | `ComandoRegistrar_Click` | Private | The biggest extraction; pure save logic |
| 4 | `Form_Open` | Private | Form lifecycle init |
| 5 | `HaHabidoCambios` | **Public** (DELETED) | Inlined into Registrar flow |
| 6 | `EstablecerDatos` | **Public** (DELETED) | Load fields from m_ObjOficinaProgramaActiva |

### 1.2 Form controls audit (rule 11A)

`Me.X` references in `.cls`: `Caption, ComandoRegistrar, OficinaPrograma, DESCRIPCION, lblTitulo, Name` (6 unique).

### 1.3 Rule #8 audit (planned helper names)

4 `OficinaProgramaAlta_*` exports — matches `modRAC / modOrganoContratacionAlta` pattern.

**Prefix disambiguation**:
- `OficinaProgramaAlta_*` — Alta/Edición form helpers (this module)
- `OficinaPrograma_*` — Gestion list helpers (modOficinaProgramaHelper, pilot PR-1)

### 1.4 Inter-form census (rule 11)

`DoCmd.OpenForm "FormOficinaPrograma"` is a built-in DoCmd call, not a method call on another form's `.cls`.

### 1.5 Binario health pre-flight (rule A)

Doctor pending — verified before module import in PR-R4.

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs.

### 1.7 Module-level declaration ordering — verified

Same pattern as Phase 2.1 / 2.2 ✅.

### 1.8 Atom count target — REWORKED

Original target: ~10 atoms across 4 helpers. Reworked target: 11 atoms.

### 1.9 Special considerations

1. **`m_ObjOficinaProgramaActiva` global** — set by `Form_FormOficinasProgramaGestion.ComandoEditar_Click` and `ComandoAlta_Click` BEFORE opening `FormOficinaPrograma`. The helpers read it as a global.
2. **`OficinaProgramaOperaciones.cls`** — operations class with `Registrar` and `Eliminar` methods. Already exists.
3. **`m_OficinaProgramaAlInicio`** — module-level form state captured during `Form_Open` via `Abrir_Inicializar`. The helper `OficinaProgramaAlta_VerificarCambios` accepts an explicit `p_ValoresOriginales` parameter (does NOT read `Me` or any module-level state).
4. **Pilot validation**: the pilot's `modOficinaProgramaHelper` already used `OficinaProgramaOperaciones` (commit 6838014), confirming the class name is correct.
