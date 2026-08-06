# Audit — Form_FormComercial (Alta/Edición, Phase 3.1 / PR-7)

> **Change**: `forms-thin-coverage` (Phase 3.1 — PR-7)
> **Form**: `Form_FormComercial` (Alta/Edición)
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
| `ComercialAlta_Abrir_Inicializar` | `(ByVal p_Modo As String, ByVal p_IDEntidad As String, ByRef p_Entidad As Object, ByRef p_Error As String) As String` |
| `ComercialAlta_VerificarCambios` | `(ByVal p_ValoresActuales As Object, ByVal p_ValoresOriginales As Object, ByRef p_Error As String) As String` |
| `ComercialAlta_Registrar` | `(ByVal p_Valores As Object, ByRef p_Error As String) As String` |
| `ComercialAlta_Cerrar` | `(ByRef p_Error As String) As String` |

### 2.2 Test atoms (14 atoms, ZERO form references)

| Action | Atoms | Coverage |
|---|---|---|
| `ComercialAlta_Abrir_Inicializar` | 4 | happy-alta, happy-edicion, sad-invalid-modo, sad-edicion-sin-id |
| `ComercialAlta_VerificarCambios` | 5 | happy-alta-all-empty, happy-alta-one-field, happy-edicion-all-equal, happy-edicion-one-diff, sad-no-actuales |
| `ComercialAlta_Registrar` | 3 | happy-alta (insert), happy-edicion (update), sad-no-valores |
| `ComercialAlta_Cerrar` | 1 | happy-cleared |
| `RunAll` | 1 | wrapper |

### 2.3 Form (thin UI wiring)

`Form_FormComercial.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `Form_Open` → `ComercialAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)` with modo/idEntidad read from `m_ObjComercialActivo` global
- `ComandoRegistrar_Click` → `ComercialAlta_Registrar(valores, p_Error)`
- `cmdSalir_Click` → `ComercialAlta_Cerrar` then close

### 2.4 Project-conventional global

Uses `m_ObjComercialActivo As Comercial` (declared in `Variables Globales.bas:237`). Parent Gestion form sets it before opening this form in edicion mode.

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modComercialAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modComercialAltaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_ComercialAltaHelper.bas` | **0** |
| `Forms(...)` in `Test_ComercialAltaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_ComercialAltaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_ComercialAltaHelper.bas` | **0** |
| `MsgBox` in helpers | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

## 4. ORIGINAL AUDIT (Phase 3.1 / PR-7 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **5** (3 Private + 2 Public — both Public get deleted per rule 11B)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdSalir_Click` | Private | Close form | KEEP inline (closes + calls `Cerrar`) | UI lifecycle |
| 2 | `ComandoAyuda_Click` | Private | `AbrirAyuda m_Error` | KEEP inline | 1 call |
| 3 | `ComandoRegistrar_Click` | Private | `HaHabidoCambios`, copy form→obj, `Registrar`, raise event | **`ComercialAlta_Registrar`** | The big extraction |
| 4 | `HaHabidoCambios` | **Public** (DELETED) | Build Dict of form values, `Helper_EntidadCRUD.HaHabidoCambiosGenerico` | **`ComercialAlta_VerificarCambios`** | Inline-merged into Registrar |
| 5 | `EstablecerDatos` | **Public** (DELETED) | Set `lblTitulo.Caption`, get `m_ComercialAlInicio`, populate controls | **`ComercialAlta_Abrir_Inicializar`** | Replaced by Abrir_Inicializar |

**Public methods to delete (rule 11B):** 2 — `HaHabidoCambios`, `EstablecerDatos`.

### 1.2 Form_Open flow (post-rework)

```text
Form_Open
  ├─ Comercial_Modo()        → "alta" or "edicion" (based on m_ObjComercialActivo)
  ├─ Comercial_IDEntidad()   → ID from m_ObjComercialActivo (empty in alta mode)
  ├─ ComercialAlta_Abrir_Inicializar(modo, idEntidad, ByRef p_Entidad, m_Error)
  └─ RenderResultadoInicializar → Me.lblTitulo, Me.Comercial, Me.DESCRIPCION
```

### 1.3 Pattern validation

Pattern replicates `Form_FormOrganoContratacion.cls` (Phase 2.2 PR-3) and `Form_FormSuministrador.cls` (Phase 2.4 PR-4). Fields per form: 2 (`Comercial`, `DESCRIPCION`). The form does NOT have a `HaHabidoCambios` public method anymore — change detection is delegated entirely to the helper.
