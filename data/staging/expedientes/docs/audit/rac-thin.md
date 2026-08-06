# Audit — Form_FormRAC (Phase 2.1 / PR-2 + REWORK / PR-R2)

> **Change**: `forms-thin-coverage` (Phase 2.1 — PR-2, batch of 4 forms; REWORK / PR-R2)
> **Form**: `Form_FormRAC` (Alta/Edición form, NOT a Gestion list)
> **Date original**: 2026-06-25
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

---

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.1 original (4 helpers / 11 atoms) | MERGED via PR #29 — `979a57a refactor(sdd): Phase 2.1 PR-2 — thin 4 forms (LugarEjecucion, GradosClasificacion, RAC, RACS)` |
| Rework / PR-R2 (4 helpers / 13 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING write-gate authorization |

---

## 2. REWORK — anti-pattern removed

### 2.0 Form profile — Alta/Edición (DIFFERENT pattern from Gestion list)

`Form_FormRAC` is the **Alta/Edición** form for RAC entities, structurally different from the three Gestion list forms in this PR. It does not have a `ListaFiltrados`; instead it has individual field controls (`RAC`, `CORREO`, `DESCRIPCION`) and a `ComandoRegistrar` button. The lifecycle is:

- `Form_Open` → `Abrir_Inicializar` (populates controls if editing) → ready state
- User edits fields
- `ComandoRegistrar_Click` → `Registrar` (validates + persists via RACOperaciones) → `Alta` or `Editado` event

### 2.1 Helper signatures (4 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `RAC_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `RAC_Verificar_Cambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `RAC_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `RAC_Cerrar` | `(ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `RAC_EstablecerDatos` (was populating form fields)
- `RAC_HaHabidoCambios` (was inline in ComandoRegistrar; now composed via `RAC_Verificar_Cambios` + `RAC_Registrar` flow in form)
- The `ByRef p_RACAlInicio As Object` parameter — replaced by `ByRef p_Entidad As Object` (in `Abrir_Inicializar`) which the helper populates with precargados.

**Helper signatures differ from pilot pattern** because this is an Alta/Edición form (not a list). The 4 helpers map to the 4 form events: Form_Open, helper-driven change detection, ComandoRegistrar_Click, Form_Close. The pilot's 5-helper pattern (Abrir/Buscar/Seleccionar/Eliminar/DobleClick) only applies to Gestion lists.

### 2.2 Test atoms (13 atoms, ZERO form references)

13 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` value stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `RAC_Abrir_Inicializar` | 4 | happy-alta, happy-edicion, sad-invalid-modo, sad-edicion-sin-id |
| `RAC_Verificar_Cambios` | 5 | happy-alta-empty, happy-alta-one-field, happy-edicion-all-equal, happy-edicion-one-diff, sad-no-actuales |
| `RAC_Registrar` | 3 | happy-alta (insert), happy-edicion (update), sad-no-valores |
| `RAC_Cerrar` | 1 | happy-cleared |

### 2.3 Form (thin UI wiring)

`Form_FormRAC.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoRegistrar_Click` — calls `RAC_ExtraerValores()` (form-local helper), then helper `RAC_Registrar`, then raises `Alta` or `Editado` event based on JSON `modo` field.
- `cmdSalir_Click` — calls `RAC_Cerrar` helper to release `m_ObjRACActivo`, then closes form.

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modRACHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modRACHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_RACHelper.bas` | **0** |
| `Forms(...)` in `Test_RACHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_RACHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_RACHelper.bas` | **0** |
| `CreateObject` in `Test_RACHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (no user-prompt seam needed — single-record form) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out

The Alta/Edición pattern (4 helpers: Abrir_Inicializar, Verificar_Cambios, Registrar, Cerrar) differs from the Gestion list pattern (5 helpers: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar, Eliminar_Borrar, DobleClick_AbrirEdicion). Phase 2.2 organo-pecal Alta/Edición forms should follow this same 4-helper pattern.

---

## 4. ORIGINAL AUDIT (Phase 2.1 / PR-2 preflight, preserved for traceability)

### 0. Form profile — Alta/Edición (DIFFERENT pattern from Gestion list)

`Form_FormRAC` is the **Alta/Edición** form for RAC entities, structurally different from the three Gestion list forms in this PR. It does not have a `ListaFiltrados`; instead it has individual field controls (`RAC`, `CORREO`, `DESCRIPCION`) and a `ComandoRegistrar` button. The lifecycle is:

- `Form_Open` → `EstablecerDatos` (populates controls if editing) → ready state
- User edits fields
- `ComandoRegistrar_Click` → `HaHabidoCambios` check → `CopiarCamposAObjeto` → `m_RACOp.Registrar` → `Alta` or `Editado` event

The `Public Function`s to remove per rule 11B are: `HaHabidoCambios`, `EstablecerDatos`. The business logic to extract is the **Registrar flow** (HaHabidoCambios + CopiarCamposAObjeto + Registrar) which currently lives inline in `ComandoRegistrar_Click`.

### 1.1 Handlers audit

Total handlers: **6** (4 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 2 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline — no helper extracted | 1 call |
| 3 | `ComandoRegistrar_Click` | Private | `HaHabidoCambios` check, `CopiarCamposAObjeto`, `m_RACOp.Registrar`, raise `Alta` or `Editado` event | **`RAC_Registrar_Guardar`** | Single extraction point for the full register flow |
| 4 | `Form_Open` | Private | `Ajustar Me`, set Caption, admin check, call `EstablecerDatos` | **`RAC_Abrir_Inicializar`** | Form lifecycle init |
| 5 | `HaHabidoCambios` | **Public** (DELETED) | Builds Dictionary from form fields, calls `Helper_EntidadCRUD.HaHabidoCambiosGenerico` | **Inlined into `RAC_Registrar_Guardar`** (logic is part of register flow; not standalone) | Per rule #11B: no Public methods on form |
| 6 | `EstablecerDatos` | **Public** (DELETED) | Loads `m_RACAlInicio` and populates form fields | **`RAC_EstablecerDatos`** | Reused as standalone helper for tests |

**Public methods to delete (rule 11B):** 2 — `HaHabidoCambios`, `EstablecerDatos`.

**MsgBox/InputBox occurrences:** 0 user-prompts. Error-path `MsgBox` only.

### 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, ComandoRegistrar, CORREO, DESCRIPCION, lblTitulo, Name, RAC` (7 unique).

**`Name ="X"` in `.form.txt`:** `ComandoAyuda, ComandoRegistrar, CORREO, DESCRIPCION, RAC, cmdSalir, lblTitulo` (7 controls). **PASS — all 7 match.**

### 1.3 Rule #8 audit (planned helper names)

Planned `RAC_*` exports (4) — note: only 4 because this is an Alta/Edición form, not a Gestion list (the Gestion list's helpers use `RACS_*` prefix to avoid collision):

```
RAC_Abrir_Inicializar
RAC_Registrar_Guardar
RAC_HaHabidoCambios
RAC_EstablecerDatos
```

**Result:** PASS — **0 collisions on 39 planned names** (PR-2 batch audit).

### 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites. Form_FormRAC is opened by Form_FormRACSGestion via `DoCmd.OpenForm "FormRAC"` — that is a built-in DoCmd call, not a method call on another form's `.cls`.

### 1.5 Binario health pre-flight (rule A)

Doctor green (passed in PR-2 batch audit).

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs against the full manifest in step 1.8.

### 1.7 Module-level declaration ordering — verified

`RACOperaciones` (existing) — declarations at top, public after private ✅. New module follows same pattern.

### 1.8 Atom count target (pre-rework)

4 actions × ~3 scenarios average = **~10 atoms minimum + RunAll wrapper** for this form (smaller than Gestion list forms because the register flow has fewer discrete actions).

| Action | Atoms planned |
|---|---|
| `RAC_Abrir_Inicializar` | 3 (alta-no-admin, edit-existing, sad-no-form) |
| `RAC_Registrar_Guardar` | 3 (happy-alta, happy-edicion, sad-no-changes) |
| `RAC_HaHabidoCambios` | 2 (no-inicial-true, all-equal-false) |
| `RAC_EstablecerDatos` | 2 (alta-no-populate, edit-populate-controls) |
| **TOTAL (this form)** | **~10 atoms + RunAll wrapper** |

### 1.9 Special considerations (preserved)

1. **`m_ObjRACActivo` global** — set by `Form_FormRACSGestion.ComandoEditar_Click` and `ComandoAlta_Click` BEFORE opening `FormRAC`. The `RAC_Abrir_Inicializar` and `RAC_Registrar` helpers read it as a global (legitimate project-conventional global state per AGENTS.md §Global Variables).
2. **`MostrarPopupProgreso` and `CerrarPopupProgreso`** — UI helpers used inside `ComandoRegistrar_Click`. The helper `RAC_Registrar` does NOT need these (it just returns JSON); they are called by the form's handler before/after the helper call.
3. **`m_RACAlInicio`** — was used by `RAC_EstablecerDatos` and `RAC_HaHabidoCambios` to detect changes vs original state. The new `RAC_Verificar_Cambios` accepts an explicit `p_ValoresOriginales` Dictionary parameter (does NOT read `Me` or any module-level state) — atoms can pass a constructed Dictionary.
