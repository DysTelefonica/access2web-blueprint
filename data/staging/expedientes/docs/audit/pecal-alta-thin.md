# Audit — Form_FormPECAL (Phase 2.2 / PR-3 + REWORK / PR-R3)

> **Change**: `forms-thin-coverage` (Phase 2.2 — PR-3, batch of 4 forms; REWORK / PR-R3)
> **Form**: `Form_FormPECAL` (Alta/Edición)
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
| `PECALAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `PECALAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `PECALAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `PECALAlta_Cerrar` | `(ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration / form-coupled logic that stays in form per e2e rule #1):
- `PECALAlta_EstablecerDatos` (was reading `p_Form.Controls("lblTitulo").Caption` / setting `p_Form.Controls("PECAL").Value`)
- `PECALAlta_HaHabidoCambios` (was reading `p_Form.Controls("PECAL").Value` etc.)
- `PECALAlta_Registrar_Guardar` (was reading form controls and calling `Helper_EntidadCRUD.CopiarCamposAObjeto` directly from form data)

### 2.2 Test atoms (14 atoms, ZERO form references)

14 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `PECALAlta_Abrir_Inicializar` | 4 | happy-alta, happy-edicion, sad-invalid-modo, sad-edicion-sin-id |
| `PECALAlta_VerificarCambios` | 5 | happy-alta-empty, happy-alta-one-field, happy-edicion-all-equal, happy-edicion-one-diff, sad-no-actuales |
| `PECALAlta_Registrar` | 3 | happy-alta (DAO cardinality), happy-edicion (DAO cardinality), sad-no-valores |
| `PECALAlta_Cerrar` | 1 | happy-cleared (verifies m_ObjPECALActiva reset) |

### 2.3 Form (thin UI wiring)

`Form_FormPECAL.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `Form_Open` — calls `PECALAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)` with primitives only
- `ComandoRegistrar_Click` — calls `PECALAlta_Registrar(valores, p_Error)` with Dictionary {PECAL, DESCRIPCION}, then raises `Alta` / `Editado` event based on payload
- `cmdSalir_Click` — calls `PECALAlta_Cerrar(m_Error)` to release the project-conventional global

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modPECALAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modPECALAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_PECALAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_PECALAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_PECALAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_PECALAltaHelper.bas` | **0** |
| `CreateObject` in `Test_PECALAltaHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.2 / PR-3 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **6** (4 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 2 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline (delegates to existing legacy `AbrirAyuda`) | 1 call |
| 3 | `ComandoRegistrar_Click` | Private | Popup progress, `HaHabidoCambios`, ensure `m_ObjPECALActiva`, `CopiarCamposAObjeto`, `PECALOperaciones.Registrar`, raise `Alta`/`Editado` event | **`PECALAlta_Registrar_Guardar`** | The biggest extraction; pure save logic |
| 4 | `HaHabidoCambios` | **Public** (DELETED) | Compare form fields to `m_PECALAlInicio` via `Helper_EntidadCRUD.HaHabidoCambiosGenerico` | **`PECALAlta_HaHabidoCambios`** | Pure-function check |
| 5 | `EstablecerDatos` | **Public** (DELETED) | Load `m_PECALAlInicio` from `m_ObjPECALActiva.IDPECAL`, populate form fields | **`PECALAlta_EstablecerDatos`** | Alta/Edición detect |
| 6 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, call `EstablecerDatos` | **`PECALAlta_Abrir_Inicializar`** | Form lifecycle init |

**Public methods to delete (rule 11B):** 2 — `HaHabidoCambios`, `EstablecerDatos`.

**MsgBox/InputBox occurrences:** 0 user-prompts in event handlers. The 2 `MsgBox` calls are error-path (trap-only).

## 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, Name, PECAL, DESCRIPCION, ComandoRegistrar, ComandoRegistrar.Enabled, lblTitulo` (7 unique references for 6 unique control names: `PECAL`, `DESCRIPCION`, `ComandoRegistrar`, `lblTitulo`, `cmdSalir`, `ComandoAyuda`).

**`Name="X"` in `.form.txt`:** (`ComandoAyuda, ComandoRegistrar, DESCRIPCION, PECAL, cmdSalir, lblTitulo`) — PASS — `PECAL` and `DESCRIPCION` match control names exactly.

## 1.3 Rule #8 audit (planned helper names)

Planned `PECALAlta_*` exports (4) — matches user-specified pattern:

```
PECALAlta_Abrir_Inicializar
PECALAlta_EstablecerDatos
PECALAlta_HaHabidoCambios
PECALAlta_Registrar_Guardar
```

**Result:** PASS — **0 collisions on 24 planned names** (12 for OrganoContratacion entity + 12 for PECAL entity). See batch audit script in `organo-contratacion-alta-thin.md` §1.3.

## 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites.

## 1.5 Binario health pre-flight (rule A)

Doctor green (passed in batch audit — see `organo-contratacion-alta-thin.md` §1.5).

## 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs against the full manifest.

## 1.7 Module-level declaration ordering — verified

All helpers follow `vba-access` §10.1 ordering.

## 1.8 Atom count target

4 actions × ~2-3 scenarios = **~10 atoms minimum + RunAll wrapper**.

| Action | Atoms planned |
|---|---|
| `PECALAlta_Abrir_Inicializar` | 2 (happy-alta, sad-no-form) |
| `PECALAlta_EstablecerDatos` | 3 (happy-alta, happy-edicion, sad-no-form) |
| `PECALAlta_HaHabidoCambios` | 2 (happy-alta-true, happy-all-equal-false) |
| `PECALAlta_Registrar_Guardar` | 3 (sad-no-changes, happy-alta, happy-edicion) |
| **TOTAL (this form)** | **~10 atoms + RunAll wrapper** |

## 1.9 Pattern validation

This form's audit pattern replicates Phase 2.1's `Form_FormRAC.cls` (commit `979a57a`):
- 4 actions extracted (Abrir_Inicializar, EstablecerDatos, HaHabidoCambios, Registrar_Guardar)
- Public methods `HaHabidoCambios` and `EstablecerDatos` deleted from form per rule 11B
- Module-level state (`m_PECALAlInicio`) passed ByRef to helpers
- Form's `ComandoRegistrar_Click` becomes a thin wrapper that calls helper and raises event based on JSON value
