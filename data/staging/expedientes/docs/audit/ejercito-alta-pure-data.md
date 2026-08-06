# Audit — Form_FormEjercito (Alta/Edición, Phase 3.1 / PR-7)

> **Change**: `forms-thin-coverage` (Phase 3.1 — PR-7)
> **Form**: `Form_FormEjercito` (Alta/Edición)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 3.1 (4 helpers / 14 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING |

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (4 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `EjercitoAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `EjercitoAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `EjercitoAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `EjercitoAlta_Cerrar` | `(ByRef p_Error As String) As String` |

### 2.2 Test atoms (14 atoms, ZERO form references)

| Action | Atoms | Coverage |
|---|---|---|
| `EjercitoAlta_Abrir_Inicializar` | 4 | happy-alta, happy-edicion, sad-invalid-modo, sad-edicion-sin-id |
| `EjercitoAlta_VerificarCambios` | 5 | happy-alta-all-empty, happy-alta-one-field, happy-edicion-all-equal, happy-edicion-one-diff, sad-no-actuales |
| `EjercitoAlta_Registrar` | 3 | happy-alta, happy-edicion, sad-no-valores |
| `EjercitoAlta_Cerrar` | 1 | happy-cleared |
| `RunAll` | 1 | wrapper |

### 2.3 Form (thin UI wiring)

`Form_FormEjercito.cls` (post-rework) reads `m_ObjEjercitoActivo` to derive modo/idEntidad, calls `EjercitoAlta_Abrir_Inicializar`, renders titulo + values, then registers via `EjercitoAlta_Registrar` on submit.

### 2.4 Project-conventional global

Uses `m_ObjEjercitoActivo As Ejercito` (declared in `Variables Globales.bas:239`).

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modEjercitoAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modEjercitoAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_EjercitoAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_EjercitoAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_EjercitoAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_EjercitoAltaHelper.bas` | **0** |
| `MsgBox` in helpers | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

## 4. ORIGINAL AUDIT (Phase 3.1 / PR-7 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **5** (3 Private + 2 Public — both Public get deleted per rule 11B)

| # | Handler | Scope | Inline logic | Proposed helper |
|---|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Close form | KEEP inline |
| 2 | `ComandoAyuda_Click` | Private | `AbrirAyuda m_Error` | KEEP inline |
| 3 | `ComandoRegistrar_Click` | Private | `HaHabidoCambios`, copy, `Registrar`, raise event | **`EjercitoAlta_Registrar`** |
| 4 | `HaHabidoCambios` | **Public** (DELETED) | Build Dict, `Helper_EntidadCRUD.HaHabidoCambiosGenerico` | **`EjercitoAlta_VerificarCambios`** |
| 5 | `EstablecerDatos` | **Public** (DELETED) | Set titulo, get `m_EjercitoAlInicio`, populate controls | **`EjercitoAlta_Abrir_Inicializar`** |

**Public methods to delete (rule 11B):** 2 — `HaHabidoCambios`, `EstablecerDatos`.

### 1.2 Pattern validation

Pattern replicates `Form_FormComercial.cls` and `Form_FormCPV.cls` (same PR). Field names: `Ejercito`, `DESCRIPCION`.
