# Audit — Form_FormRACSGestion (Phase 2.1 / PR-2 + REWORK / PR-R2)

> **Change**: `forms-thin-coverage` (Phase 2.1 — PR-2, batch of 4 forms; REWORK / PR-R2)
> **Form**: `Form_FormRACSGestion`
> **Date original**: 2026-06-25
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

---

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.1 original (8 helpers / 19 atoms) | MERGED via PR #29 — `979a57a refactor(sdd): Phase 2.1 PR-2 — thin 4 forms (LugarEjecucion, GradosClasificacion, RAC, RACS)` |
| Rework / PR-R2 (5 helpers / 20 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING write-gate authorization |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.1 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. El pilot (PR #33 commit 6838014) demostró la corrección; ahora se escala a las 4 formas de Phase 2.1.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `RACS_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `RACS_Buscar_Listar` | `(ByVal p_RACs As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `RACS_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_RACs As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `RACS_Eliminar_Borrar` | `(ByVal p_RAC As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `RACS_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `RACS_Alta_Abrir` (was opening `FormRAC`)
- `RACS_Editar_Abrir` (was opening `FormRAC`)
- `RACS_Limpiar_Reset` (was resetting state)

**Naming convention:** `RACS_` (plural) used for this list/gestion form to avoid collision with `RAC_*` helpers extracted for the Alta/Edición form (`Form_FormRAC`) in `modRACHelper`. The two modules have different roles; this disambiguates the public surface per rule #9.

### 2.2 Test atoms (20 atoms, ZERO form references)

20 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `RACS_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `RACS_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filtered-rows, edge-long-filter, adversarial-filter-with-semicolons, adversarial-semicolon-in-name |
| `RACS_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `RACS_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `RACS_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormRACSGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — clears `m_ObjRACActivo`, closes `FormRAC` if open, opens it.
- `ComandoEditar_Click` — sets `m_ObjRACActivo = m_RACSeleccionado`, opens `FormRAC` for edit.
- `ComandoLimpiar_Click` — `Me.RAC.Value = Null`, reset selection (pure UI action, no helper).

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modRACSHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modRACSHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_RACSHelper.bas` | **0** |
| `Forms(...)` in `Test_RACSHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_RACSHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_RACSHelper.bas` | **0** |
| `CreateObject` in `Test_RACSHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out (Phase 2.2 organo-pecal — PR-R3)

Same anti-pattern lives in `modOrganoContratacionHelper`, `modPECALHelper`, `modSuministradorHelper`, `modUsuariosHelper`, and their *Alta* siblings. Phase 2.2 (PR-R3) should apply this same rework.

---

## 4. ORIGINAL AUDIT (Phase 2.1 / PR-2 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **16** (14 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline (UI lifecycle adapter) | Pure orchestration |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, disable buttons, set RowSource, fetch `m_ObjEntorno.RACs`, filter by `Me.RAC`, populate ListBox, fire `ListaFiltrados_Click` | **`RACS_Buscar_Filtrar`** | Note: prefix is `RACS_` (plural) to avoid collision with `RAC_*` helpers extracted for Form_FormRAC |
| 4 | `ComandoAlta_Click` | Private | Clear `m_ObjRACActivo`, close `FormRAC` if open, open it, set `m_FormRAC` | **`RACS_Alta_Abrir`** | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline — no helper extracted | 1 call |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error`, `AvanceCerrar` | **`RACS_Buscar_Filtrar`** (wrapper logic stays in caller) | The Public wrapper around `Filtrar`; thin |
| 7 | `ComandoEliminar_Click` | Private | Get selection, prompt user, call `Helper_EntidadCRUD.EliminarEntidadGenerico`, fire `Eliminar` event, refresh | **`RACS_Eliminar_Borrar`** | `p_PromptResult` injected |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.RAC = Null`, call `ComandoBuscar_Click` | **`RACS_Limpiar_Reset`** | Trivial wrapper |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjRACActivo`, open `FormRAC` | **`RACS_Editar_Abrir`** | Selection-aware open |
| 10 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, `cmdElegir.Visible` toggle, call `Filtrar m_Error` | **`RACS_Abrir_Inicializar`** | Form lifecycle init |
| 11 | `ListaFiltrados_Click` | Private | Get selection from ListBox, enable `ComandoEditar`/`ComandoEliminar` | **`RACS_Seleccionar_Cargar`** | Selection state sync |
| 12 | `ListaFiltrados_DblClick` | Private | Dispatch to `cmdElegir_Click` or `ComandoEditar_Click` | **`RACS_DobleClick_AbrirEdicion`** | Dispatcher |
| 13 | `RAC_KeyDown` | Private | Detect Enter key, call `ComandoBuscar_Click` | KEEP inline (delegates) | Form event |
| 14 | `m_FormRAC_Alta` | Private (WithEvents) | Call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |
| 15 | `m_FormRAC_Editado` | Private (WithEvents) | Call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

**MsgBox/InputBox occurrences:** 1 user-prompt in `ComandoEliminar_Click` (line 187) — atomizable via `p_PromptResult`. The rest are error-path `MsgBox` calls.

### 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, cmdElegir, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, Name, OpenArgs, RAC` (12 unique).

**`Name ="X"` in `.form.txt`:** `ComandoAyuda, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, RAC, cmdElegir, cmdSalir, lblTitulo` (11 controls). **PASS — `RAC` matches.**

### 1.3 Rule #8 audit (planned helper names)

Planned `RACS_*` exports (8):

```
RACS_Alta_Abrir
RACS_Buscar_Filtrar
RACS_Editar_Abrir
RACS_Eliminar_Borrar
RACS_Limpiar_Reset
RACS_Abrir_Inicializar
RACS_Seleccionar_Cargar
RACS_DobleClick_AbrirEdicion
```

**Naming convention:** `RACS_` (plural) used for this list/gestion form to avoid collision with `RAC_*` helpers extracted for the Alta/Edición form (`Form_FormRAC`). The two modules have different roles; this disambiguates the public surface per rule #9.

**Result:** PASS — **0 collisions on 39 planned names** (PR-2 batch audit).

### 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites (`Forms("FormRAC")` is a built-in Forms collection lookup, not a method call on another form's `.cls`).

### 1.5 Binario health pre-flight (rule A)

Doctor green (passed in PR-2 batch audit).

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs against the full manifest in step 1.8.

### 1.7 Module-level declaration ordering — verified

`RACOperaciones` (existing) — declarations at top, public after private ✅. New module follows same pattern.

### 1.8 Atom count target (pre-rework)

8 actions × ~3 scenarios average = **~19 atoms minimum + RunAll wrapper** for this form.

| Action | Atoms planned |
|---|---|
| `RACS_Abrir_Inicializar` | 2 |
| `RACS_Buscar_Filtrar` | 3 |
| `RACS_Seleccionar_Cargar` | 2 |
| `RACS_Alta_Abrir` | 2 |
| `RACS_Editar_Abrir` | 2 |
| `RACS_Eliminar_Borrar` | 3 |
| `RACS_Limpiar_Reset` | 2 |
| `RACS_DobleClick_AbrirEdicion` | 3 |
| **TOTAL (this form)** | **~19 atoms + RunAll wrapper** |
