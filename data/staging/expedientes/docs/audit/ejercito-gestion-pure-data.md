# Audit — Form_FormEjercitosGestion (Phase 3.1 / PR-7)

> **Change**: `forms-thin-coverage` (Phase 3.1 — PR-7, batch of 6 forms)
> **Form**: `Form_FormEjercitosGestion` (CRUD list for Ejercito)
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
| `Ejercito_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `Ejercito_Buscar_Listar` | `(ByVal p_Ejercitos As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `Ejercito_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_Ejercitos As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `Ejercito_Eliminar_Borrar` | `(ByVal p_Ejercito As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `Ejercito_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

### 2.2 Test atoms (15 atoms, ZERO form references)

15 atoms + 1 RunAll wrapper.

| Action | Atoms | Coverage |
|---|---|---|
| `Ejercito_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `Ejercito_Buscar_Listar` | 3 | happy-all-rows, edge-empty-collection, happy-filtered-rows |
| `Ejercito_Seleccionar_Cargar` | 2 | happy-admin, sad-empty-selection |
| `Ejercito_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `Ejercito_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |
| `RunAll` | 1 | wrapper |

### 2.3 Form (thin UI wiring)

`Form_FormEjercitosGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormEjercito"`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormEjercito"` after setting `m_ObjEjercitoActivo`
- `ComandoLimpiar_Click` — `Me.Ejercito.Value = Null`

### 2.4 Filter field + listbox note

Filter is on `Me.Ejercito` (the name field). Listbox is 3-col (`IDEjercito;Ejercito;DESCRIPCIÓN`) — **unique among the 3 entities in PR-7**. The other two (Comercial, CPV) are 2-col.

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modEjercitoHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modEjercitoHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_EjercitoHelper.bas` | **0** |
| `Forms(...)` in `Test_EjercitoHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_EjercitoHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_EjercitoHelper.bas` | **0** |
| `MsgBox` in helpers | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

## 4. ORIGINAL AUDIT (Phase 3.1 / PR-7 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public) — identical structure to `Form_FormComercialesGestion.cls`.

| # | Handler | Scope | Inline logic | Proposed helper |
|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline |
| 3 | `Filtrar` | **Public** (DELETED) | Reset, fetch, filter, populate, fire | **`Ejercito_Buscar_Listar`** |
| 4 | `ComandoAlta_Click` | Private | Open FormEjercito | KEEP inline |
| 5 | `ComandoAyuda_Click` | Private | `AbrirAyuda` | KEEP inline |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar` | wrapper |
| 7 | `ComandoEliminar_Click` | Private | Prompt + `Helper_EntidadCRUD.EliminarEntidadGenerico` | **`Ejercito_Eliminar_Borrar`** |
| 8 | `ComandoLimpiar_Click` | Private | `Me.Ejercito = Null` + buscar | KEEP inline |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjEjercitoActivo`, open `FormEjercito` | KEEP inline |
| 10 | `Ejercito_KeyDown` | Private | Detect Enter key | KEEP inline |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure, `Filtrar` | **`Ejercito_Abrir_Inicializar`** |
| 12 | `ListaFiltrados_Click` | Private | Get selection, enable Editar/Eliminar | **`Ejercito_Seleccionar_Cargar`** |
| 13 | `ListaFiltrados_DblClick` | Private | cmdElegir / ComandoEditar dispatch | **`Ejercito_DobleClick_AbrirEdicion`** |
| 14 | `m_FormEjercito_Alta` | Private (WithEvents) | Reset cache, buscar | KEEP inline |
| 15 | `m_FormEjercito_Editado` | Private (WithEvents) | Reset cache, buscar | KEEP inline |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

## 1.2 Pattern validation

Pattern replicates the rest of Phase 3.1's `Form_FormComercialesGestion.cls` and `Form_FormCPVsGestion.cls`. The 3-col listbox is the only shape difference — `Ejercito_BuildRowLine` produces 3 fields, vs 2 for the others.
