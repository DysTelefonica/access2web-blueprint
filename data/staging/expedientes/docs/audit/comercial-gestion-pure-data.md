# Audit — Form_FormComercialesGestion (Phase 3.1 / PR-7)

> **Change**: `forms-thin-coverage` (Phase 3.1 — PR-7, batch of 6 forms)
> **Form**: `Form_FormComercialesGestion` (CRUD list for Comercial)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 3.1 (5 helpers / 18 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING |

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `Comercial_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `Comercial_Buscar_Listar` | `(ByVal p_Comerciales As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `Comercial_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_Comerciales As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `Comercial_Eliminar_Borrar` | `(ByVal p_Comercial As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `Comercial_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `Comercial_Alta_Abrir` (was opening `FormComercial`)
- `Comercial_Editar_Abrir` (was opening `FormComercial`)
- `Comercial_Limpiar_Reset` (was resetting form state directly)

### 2.2 Test atoms (18 atoms, ZERO form references)

18 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

| Action | Atoms | Coverage |
|---|---|---|
| `Comercial_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `Comercial_Buscar_Listar` | 4 | happy-all-rows, edge-empty-collection, happy-filtered-rows, adversarial-filter-with-semicolons |
| `Comercial_Seleccionar_Cargar` | 3 | happy-admin, happy-non-admin, sad-empty-selection |
| `Comercial_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `Comercial_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |
| `RunAll` | 1 | wrapper |

### 2.3 Form (thin UI wiring)

`Form_FormComercialesGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormComercial"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormComercial"` after setting `m_ObjComercialActivo`
- `ComandoLimpiar_Click` — `Me.Comercial.Value = Null`, reset selection

### 2.4 Filter field note

Per `Form_FormComercialesGestion.cls` legacy audit, the filter is on `Me.Comercial` (column 0 of the listbox) — same shape as `OrganoContratacion` / `LugarEjecucion`. The listbox itself is 2-col (`IDComercial;Comercial`) per the legacy `Filtrar` RowSource; DESCRIPCION is not shown in the listbox.

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modComercialHelper.bas` | **0** (only comments documenting removal) |
| `DoCmd.OpenForm` in `modComercialHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_ComercialHelper.bas` | **0** |
| `Forms(...)` in `Test_ComercialHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_ComercialHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_ComercialHelper.bas` | **0** |
| `CreateObject` in `Test_ComercialHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

## 4. ORIGINAL AUDIT (Phase 3.1 / PR-7 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline (UI lifecycle adapter) | Pure orchestration |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, disable buttons, set RowSource, fetch `m_ObjEntorno.Comerciales`, filter by `Me.Comercial`, populate ListBox, fire `ListaFiltrados_Click` | **`Comercial_Buscar_Listar`** | The biggest extraction |
| 4 | `ComandoAlta_Click` | Private | Clear `m_ObjComercialActivo`, close `FormComercial` if open, open it, set `m_FormComercial` | **`Comercial_Alta_Abrir`** (KEEP inline) | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline — no helper extracted | 1 call |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error`, `AvanceCerrar` | **`Comercial_Buscar_Listar`** (wrapper logic stays in caller) | The Public wrapper around `Filtrar`; thin |
| 7 | `ComandoEliminar_Click` | Private | Get selection from ListBox if Nothing, prompt user via `MsgBox`, call `Helper_EntidadCRUD.EliminarEntidadGenerico`, fire `Eliminar` event, refresh | **`Comercial_Eliminar_Borrar`** | `p_PromptResult` injected |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.Comercial = Null`, call `ComandoBuscar_Click` | KEEP inline (trivial wrapper) | Pure UI action |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjComercialActivo`, open `FormComercial` | **`Comercial_Editar_Abrir`** (KEEP inline) | Selection-aware open |
| 10 | `Comercial_KeyDown` | Private | Detect Enter key | KEEP inline — no helper | UI event |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, `cmdElegir.Visible` toggle, call `Filtrar m_Error` | **`Comercial_Abrir_Inicializar`** | Form lifecycle init |
| 12 | `ListaFiltrados_Click` | Private | Get selection from ListBox, enable `ComandoEditar`/`ComandoEliminar` | **`Comercial_Seleccionar_Cargar`** | Selection state sync |
| 13 | `ListaFiltrados_DblClick` | Private | If `cmdElegir` visible → `cmdElegir_Click`. Else if `ComandoEditar.Enabled` → `ComandoEditar_Click` | **`Comercial_DobleClick_AbrirEdicion`** | Dispatcher |
| 14 | `m_FormComercial_Alta` | Private (WithEvents) | Reset `m_ObjEntorno.Comerciales`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |
| 15 | `m_FormComercial_Editado` | Private (WithEvents) | Reset `m_ObjEntorno.Comerciales`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

## 1.2 Atom count target

5 actions × ~3 scenarios average = **~17 atoms minimum + RunAll wrapper**. Actual: 18 atoms + RunAll.

## 1.3 Pattern validation

Pattern replicates Phase 2.4's `Form_FormSuministradoresGestion.cls` (commit `4ae305c`):
- 5 actions extracted (Gestion pattern)
- Public methods `Filtrar` and `ComandoBuscar_Click` deleted from form per rule 11B
- Module-level state (`m_ComercialSeleccionado`, `m_FormComercial`) — selection held as Object to allow class swap
- Form's `ComandoEliminar_Click` checks JSON value to dispatch `RaiseEvent Eliminar` only on actual deletion
