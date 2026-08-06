# Audit — Form_FormGradosClasificacionGestion (Phase 2.1 / PR-2 + REWORK / PR-R2)

> **Change**: `forms-thin-coverage` (Phase 2.1 — PR-2, batch of 4 forms; REWORK / PR-R2)
> **Form**: `Form_FormGradosClasificacionGestion`
> **Date original**: 2026-06-25
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

---

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.1 original (8 helpers / 19 atoms) | MERGED via PR #29 — `979a57a refactor(sdd): Phase 2.1 PR-2 — thin 4 forms (LugarEjecucion, GradosClasificacion, RAC, RACS)` |
| Rework / PR-R2 (5 helpers / 20 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING write-gate authorization |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.1 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. El pilot (PR #33 commit 6838014) demostró la corrección; ahora se escala a las 4 formas de Phase 2.1.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `GradosClasificacion_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `GradosClasificacion_Buscar_Listar` | `(ByVal p_Grados As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `GradosClasificacion_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_Grados As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `GradosClasificacion_Eliminar_Borrar` | `(ByVal p_GradoClasificacion As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `GradosClasificacion_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `GradosClasificacion_Alta_Abrir` (was opening `FormGradoClasificacion`)
- `GradosClasificacion_Editar_Abrir` (was opening `FormGradoClasificacion`)
- `GradosClasificacion_Limpiar_Reset` (was resetting state)

**Property case mismatch notes (schema-first, preserved from original audit)**:
- Form control name:  `GradoClasificacion`  (PascalCase)
- Table column:       `IDGradoClasificacion` (uppercase 'D')
- Entity property:    `IdGradoClasificacion` (lowercase 'd' — canonical class name)
- Constructor param:  `p_IdGradoClasificacion` (lowercase 'd')

The Dictionary stub keys use the TABLE column name (`IDGradoClasificacion`). When wrapping the Dictionary into the entity for CRUD operations, the helper uses the class property name (`IdGradoClasificacion`).

### 2.2 Test atoms (20 atoms, ZERO form references)

20 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `GradosClasificacion_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `GradosClasificacion_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filtered-rows, edge-long-filter, adversarial-filter-with-semicolons, adversarial-semicolon-in-name |
| `GradosClasificacion_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `GradosClasificacion_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `GradosClasificacion_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormGradosClasificacionGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormGradoClasificacion"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormGradoClasificacion"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoLimpiar_Click` — `Me.GradoClasificacion.Value = Null`, reset selection (pure UI action, no helper)

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modGradosClasificacionHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modGradosClasificacionHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_GradosClasificacionHelper.bas` | **0** |
| `Forms(...)` in `Test_GradosClasificacionHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_GradosClasificacionHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_GradosClasificacionHelper.bas` | **0** |
| `CreateObject` in `Test_GradosClasificacionHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out (Phase 2.2 organo-pecal — PR-R3)

Same anti-pattern lives in `modOrganoContratacionHelper`, `modPECALHelper`, `modSuministradorHelper`, `modUsuariosHelper`, and their *Alta* siblings. Phase 2.2 (PR-R3) should apply this same rework.

---

## 4. ORIGINAL AUDIT (Phase 2.1 / PR-2 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline (UI lifecycle adapter) | Pure orchestration |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, disable buttons, set RowSource, fetch `m_ObjEntorno.GradosClasificaciones`, filter by `Me.GradoClasificacion`, populate ListBox, fire `ListaFiltrados_Click` | **`GradosClasificacion_Buscar_Filtrar`** | Same pattern as LugarEjecucion_Buscar_Filtrar |
| 4 | `ComandoAlta_Click` | Private | Clear `m_ObjGradoClasificacionActivo`, close `FormGradoClasificacion` if open, open it, set `m_FormGradoClasificacion` | **`GradosClasificacion_Alta_Abrir`** | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline (delegates to existing legacy `AbrirAyuda`) — no helper extracted | 1 call |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error`, `AvanceCerrar` | **`GradosClasificacion_Buscar_Filtrar`** (wrapper logic stays in caller) | The Public wrapper around `Filtrar`; thin |
| 7 | `ComandoEliminar_Click` | Private | Get selection, prompt user, call `Helper_EntidadCRUD.EliminarEntidadGenerico`, fire `Eliminar` event, refresh | **`GradosClasificacion_Eliminar_Borrar`** | `p_PromptResult` injected |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.GradoClasificacion = Null`, call `ComandoBuscar_Click` | **`GradosClasificacion_Limpiar_Reset`** | Trivial wrapper |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjGradoClasificacionActivo`, open `FormGradoClasificacion` | **`GradosClasificacion_Editar_Abrir`** | Selection-aware open |
| 10 | `GradoClasificacion_Exit` | Private | Call `ComandoBuscar_Click` | KEEP inline (delegates) | Form event |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, `cmdElegir.Visible` toggle, call `Filtrar m_Error` | **`GradosClasificacion_Abrir_Inicializar`** | Form lifecycle init |
| 12 | `ListaFiltrados_Click` | Private | Get selection from ListBox, enable `ComandoEditar`/`ComandoEliminar` | **`GradosClasificacion_Seleccionar_Cargar`** | Selection state sync |
| 13 | `ListaFiltrados_DblClick` | Private | Dispatch to `cmdElegir_Click` or `ComandoEditar_Click` | **`GradosClasificacion_DobleClick_AbrirEdicion`** | Dispatcher |
| 14 | `m_FormGradoClasificacion_Alta` | Private (WithEvents) | Reset `m_ObjEntorno.GradosClasificaciones`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |
| 15 | `m_FormGradoClasificacion_Editado` | Private (WithEvents) | Reset `m_ObjEntorno.GradosClasificaciones`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

**MsgBox/InputBox occurrences:** 1 user-prompt in `ComandoEliminar_Click` (line 190) — atomizable via `p_PromptResult`. The rest are error-path `MsgBox` calls for trap-only display.

### 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, cmdElegir, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, GradoClasificacion, ListaFiltrados, Name, OpenArgs` (12 unique).

**`Name ="X"` in `.form.txt`:** `ComandoAyuda, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, GradoClasificacion, ListaFiltrados, cmdElegir, cmdSalir, lblTitulo` (11 controls). **PASS — `GradoClasificacion` matches.**

### 1.3 Rule #8 audit (planned helper names)

Planned `GradosClasificacion_*` exports (8):

```
GradosClasificacion_Alta_Abrir
GradosClasificacion_Buscar_Filtrar
GradosClasificacion_Editar_Abrir
GradosClasificacion_Eliminar_Borrar
GradosClasificacion_Limpiar_Reset
GradosClasificacion_Abrir_Inicializar
GradosClasificacion_Seleccionar_Cargar
GradosClasificacion_DobleClick_AbrirEdicion
```

**Result:** PASS — **0 collisions** (verified in the PR-2 batch audit covering all 39 planned names).

### 1.4 Inter-form census (rule 11)

**Result for this form:** 0 inter-form callsites (`Forms("FormGradoClasificacion")` is a built-in Forms collection lookup, not a method call on another form's `.cls`).

### 1.5 Binario health pre-flight (rule A)

Doctor green (passed in PR-2 batch audit).

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs against the full manifest in step 1.8.

### 1.7 Module-level declaration ordering — verified

`GradoClasificacionOperaciones` (existing) — declarations at top, public after private ✅. New module follows same pattern.

### 1.8 Atom count target (pre-rework)

8 actions × ~3 scenarios average = **~19 atoms minimum + RunAll wrapper** for this form.

| Action | Atoms planned |
|---|---|
| `GradosClasificacion_Abrir_Inicializar` | 2 |
| `GradosClasificacion_Buscar_Filtrar` | 3 |
| `GradosClasificacion_Seleccionar_Cargar` | 2 |
| `GradosClasificacion_Alta_Abrir` | 2 |
| `GradosClasificacion_Editar_Abrir` | 2 |
| `GradosClasificacion_Eliminar_Borrar` | 3 |
| `GradosClasificacion_Limpiar_Reset` | 2 |
| `GradosClasificacion_DobleClick_AbrirEdicion` | 3 |
| **TOTAL (this form)** | **~19 atoms + RunAll wrapper** |
