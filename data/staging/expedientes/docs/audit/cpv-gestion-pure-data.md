# Audit — Form_FormCPVsGestion (Phase 3.1 / PR-7)

> **Change**: `forms-thin-coverage` (Phase 3.1 — PR-7, batch of 6 forms)
> **Form**: `Form_FormCPVsGestion` (CRUD list for CPV)
> **Date**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 3.1 (5 helpers / 15 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING |

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `CPV_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `CPV_Buscar_Listar` | `(ByVal p_CPVs As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `CPV_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_CPVs As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `CPV_Eliminar_Borrar` | `(ByVal p_CPV As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `CPV_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

### 2.2 Test atoms (15 atoms, ZERO form references)

15 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs.

| Action | Atoms | Coverage |
|---|---|---|
| `CPV_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `CPV_Buscar_Listar` | 3 | happy-all-rows, edge-empty-collection, happy-filtered-rows |
| `CPV_Seleccionar_Cargar` | 2 | happy-admin, sad-empty-selection |
| `CPV_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `CPV_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |
| `RunAll` | 1 | wrapper |

### 2.3 Form (thin UI wiring)

`Form_FormCPVsGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormCPV"`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormCPV"` after setting `m_ObjCPVActivo`
- `ComandoLimpiar_Click` — `Me.CPV.Value = Null`

### 2.4 Filter field note

Filter is on `Me.CPV` (the name field), matches legacy behavior. Listbox is 2-col (`IDCPV;CPV`).

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modCPVHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modCPVHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_CPVHelper.bas` | **0** |
| `Forms(...)` in `Test_CPVHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_CPVHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_CPVHelper.bas` | **0** |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

## 4. ORIGINAL AUDIT (Phase 3.1 / PR-7 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public) — identical structure to `Form_FormComercialesGestion.cls`.

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline | UI lifecycle adapter |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset, disable buttons, fetch `m_ObjEntorno.CPVs`, filter by `Me.CPV`, populate ListBox, fire `ListaFiltrados_Click` | **`CPV_Buscar_Listar`** | Extraction |
| 4 | `ComandoAlta_Click` | Private | Open FormCPV | KEEP inline (UI orchestration) | |
| 5 | `ComandoAyuda_Click` | Private | `AbrirAyuda m_Error` | KEEP inline | |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error` | wrapper | |
| 7 | `ComandoEliminar_Click` | Private | Prompt + `Helper_EntidadCRUD.EliminarEntidadGenerico` | **`CPV_Eliminar_Borrar`** | |
| 8 | `ComandoLimpiar_Click` | Private | `Me.CPV = Null` + buscar | KEEP inline | |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjCPVActivo`, open `FormCPV` | KEEP inline | |
| 10 | `CPV_KeyDown` | Private | Detect Enter key | KEEP inline | |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure, `Filtrar` | **`CPV_Abrir_Inicializar`** | |
| 12 | `ListaFiltrados_Click` | Private | Get selection, enable Editar/Eliminar | **`CPV_Seleccionar_Cargar`** | |
| 13 | `ListaFiltrados_DblClick` | Private | cmdElegir / ComandoEditar dispatch | **`CPV_DobleClick_AbrirEdicion`** | |
| 14 | `m_FormCPV_Alta` | Private (WithEvents) | Reset cache, buscar | KEEP inline | |
| 15 | `m_FormCPV_Editado` | Private (WithEvents) | Reset cache, buscar | KEEP inline | |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

## 1.2 Pattern validation

Pattern replicates `Form_FormComercialesGestion.cls` and Phase 2.x's `Form_FormOrganoContratacionGestion.cls`. Only difference: filter field name is `CPV` instead of `Comercial`.
