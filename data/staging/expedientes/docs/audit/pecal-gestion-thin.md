# Audit — Form_FormPECALESGestion (Phase 2.2 / PR-3 + REWORK / PR-R3)

> **Change**: `forms-thin-coverage` (Phase 2.2 — PR-3, batch of 4 forms; REWORK / PR-R3)
> **Form**: `Form_FormPECALESGestion` (CRUD list)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.2 original (8 helpers / ~19 atoms) | MERGED via PR #30 — `97bdb21 feat(sdd): Phase 2.2 PR-3 — thin 4 forms (OrganoContratacion, PECAL)` |
| Rework / PR-R3 (5 helpers / 21 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING user paste in VBE |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.2 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. Regla rota: `access-vba-e2e-methodology` #1. El pilot (PR #33 commit 6838014) y PR #34 (lugar-grados-rac) demostraron la corrección; ahora se aplica a las 4 formas de Phase 2.2.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `PECAL_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `PECAL_Buscar_Listar` | `(ByVal p_PECALs As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `PECAL_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_PECALs As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `PECAL_Eliminar_Borrar` | `(ByVal p_PECAL As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `PECAL_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `PECAL_Alta_Abrir` (was opening `FormPECAL`)
- `PECAL_Editar_Abrir` (was opening `FormPECAL`)
- `PECAL_Limpiar_Reset` (was resetting form state directly)

### 2.2 Test atoms (21 atoms, ZERO form references)

21 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `PECAL_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `PECAL_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filtered-rows, edge-long-filter, adversarial-filter-with-semicolons, adversarial-semicolon-in-name |
| `PECAL_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `PECAL_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `PECAL_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormPECALESGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormPECAL"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormPECAL"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoLimpiar_Click` — `Me.PECAL.Value = Null`, reset selection (pure UI action, no helper)

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modPECALHelper.bas` | **0** (only comments documenting removal) |
| `DoCmd.OpenForm` in `modPECALHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_PECALHelper.bas` | **0** |
| `Forms(...)` in `Test_PECALHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_PECALHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_PECALHelper.bas` | **0** |
| `CreateObject` in `Test_PECALHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.2 / PR-3 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event | KEEP inline (UI lifecycle adapter) | Pure orchestration |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, disable buttons, set RowSource, fetch `m_ObjEntorno.PECALES`, filter by `Me.PECAL`, populate ListBox, fire `ListaFiltrados_Click` | **`PECAL_Buscar_Filtrar`** | The biggest extraction |
| 4 | `ComandoAlta_Click` | Private | Clear `m_ObjPECALActiva`, close `FormPECAL` if open, open it, set `m_FormPECAL` | **`PECAL_Alta_Abrir`** | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline — no helper extracted | 1 call |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error`, `AvanceCerrar` | **`PECAL_Buscar_Filtrar`** (wrapper logic stays in caller) | The Public wrapper around `Filtrar`; thin |
| 7 | `ComandoEliminar_Click` | Private | Get selection from ListBox if Nothing, prompt user via `MsgBox`, call `Helper_EntidadCRUD.EliminarEntidadGenerico`, fire `Eliminar` event, refresh | **`PECAL_Eliminar_Borrar`** | `p_PromptResult` injected |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.PECAL = Null`, call `ComandoBuscar_Click` | **`PECAL_Limpiar_Reset`** | Trivial wrapper |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjPECALActiva`, open `FormPECAL` | **`PECAL_Editar_Abrir`** | Selection-aware open |
| 10 | `PECAL_Exit` | Private | Detect exit event, call `ComandoBuscar_Click` | KEEP inline (delegates) | Form event |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, `cmdElegir.Visible` toggle, call `Filtrar m_Error` | **`PECAL_Abrir_Inicializar`** | Form lifecycle init |
| 12 | `ListaFiltrados_Click` | Private | Get selection from ListBox, enable `ComandoEditar`/`ComandoEliminar` | **`PECAL_Seleccionar_Cargar`** | Selection state sync |
| 13 | `ListaFiltrados_DblClick` | Private | If `cmdElegir` visible → `cmdElegir_Click`. Else if `ComandoEditar.Enabled` → `ComandoEditar_Click` | **`PECAL_DobleClick_AbrirEdicion`** | Dispatcher |
| 14 | `m_FormPECAL_Alta` | Private (WithEvents) | Reset `m_ObjEntorno.PECALES`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |
| 15 | `m_FormPECAL_Editado` | Private (WithEvents) | Reset `m_ObjEntorno.PECALES`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

**MsgBox/InputBox occurrences in event handlers:** 1 user-prompt in `ComandoEliminar_Click` (line 187: "¿Desea realmente borrar la PECAL seleccionada?") — atomizable via `p_PromptResult`.

## 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, cmdElegir, ComandoAlta, ComandoEditar, ComandoEliminar, ListaFiltrados, ListaFiltrados.Column(0), ListaFiltrados.Selected, Name, OpenArgs, PECAL` (12 unique references for 10 unique control names).

**`Name="X"` in `.form.txt`:** (matches via preflight audit) — `ComandoAyuda, ComandoAlta, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, PECAL, cmdElegir, cmdSalir, lblTitulo` — PASS.

## 1.3 Rule #8 audit (planned helper names)

Planned `PECAL_*` exports (8):

```
PECAL_Alta_Abrir
PECAL_Buscar_Filtrar
PECAL_Editar_Abrir
PECAL_Eliminar_Borrar
PECAL_Limpiar_Reset
PECAL_Abrir_Inicializar
PECAL_Seleccionar_Cargar
PECAL_DobleClick_AbrirEdicion
```

**Result:** PASS — **0 collisions on 24 planned names** (12 for OrganoContratacion entity + 12 for PECAL entity). See batch audit script in `organo-contratacion-alta-thin.md` §1.3.

## 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites.

## 1.5 Binario health pre-flight (rule A)

Doctor green.

## 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs.

## 1.7 Module-level declaration ordering — verified

All helpers follow `vba-access` §10.1 ordering.

## 1.8 Atom count target

8 actions × ~3 scenarios average = **~19 atoms minimum + RunAll wrapper**.

| Action | Atoms planned |
|---|---|
| `PECAL_Abrir_Inicializar` | 2 |
| `PECAL_Buscar_Filtrar` | 3 |
| `PECAL_Seleccionar_Cargar` | 2 |
| `PECAL_Alta_Abrir` | 2 |
| `PECAL_Editar_Abrir` | 2 |
| `PECAL_Eliminar_Borrar` | 3 |
| `PECAL_Limpiar_Reset` | 2 |
| `PECAL_DobleClick_AbrirEdicion` | 3 |
| **TOTAL (this form)** | **~19 atoms + RunAll wrapper** |

## 1.9 Pattern validation

This form's audit pattern replicates Phase 2.1's `Form_FormLugarEjecucionGestion.cls` (commit `979a57a`):
- 8 actions extracted (Gestion pattern)
- Public methods `Filtrar` and `ComandoBuscar_Click` deleted from form per rule 11B
- Module-level state (`m_PECALSeleccionada`, `m_FormPECAL`) passed ByRef to helpers
- Form's `ComandoEliminar_Click` checks JSON value to dispatch `RaiseEvent Eliminar` only on actual deletion
