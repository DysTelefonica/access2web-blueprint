# Audit — Form_FormLugarEjecucionGestion (Phase 2.1 / PR-2 + REWORK / PR-R2)

> **Change**: `forms-thin-coverage` (Phase 2.1 — PR-2, batch of 4 forms; REWORK / PR-R2)
> **Form**: `Form_FormLugarEjecucionGestion`
> **Date original**: 2026-06-25
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

---

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.1 original (8 helpers / 19 atoms) | MERGED via PR #29 — `979a57a refactor(sdd): Phase 2.1 PR-2 — thin 4 forms (LugarEjecucion, GradosClasificacion, RAC, RACS)` |
| Rework / PR-R2 (5 helpers / 20 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING write-gate authorization |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.1 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless porque el COM abría un form modal real. Regla rota: `access-vba-e2e-methodology` #1 (forms son wiring UI thin, helpers son datos puros). El pilot (PR #33 commit 6838014) demostró la corrección; ahora se escala a las 4 formas de Phase 2.1.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `LugarEjecucion_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `LugarEjecucion_Buscar_Listar` | `(ByVal p_Lugares As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `LugarEjecucion_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_Lugares As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `LugarEjecucion_Eliminar_Borrar` | `(ByVal p_LugarEjecucion As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `LugarEjecucion_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `LugarEjecucion_Alta_Abrir` (was opening `FormLugarEjecucion`)
- `LugarEjecucion_Editar_Abrir` (was opening `FormLugarEjecucion`)
- `LugarEjecucion_Limpiar_Reset` (was resetting state)

### 2.2 Test atoms (20 atoms, ZERO form references)

20 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `LugarEjecucion_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `LugarEjecucion_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filtered-rows, edge-long-filter, adversarial-filter-with-semicolons, adversarial-semicolon-in-name |
| `LugarEjecucion_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `LugarEjecucion_Eliminar_Borrar` | 3 | sad-cancelled, happy-deleted, sad-no-entity |
| `LugarEjecucion_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormLugarEjecucionGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormLugarEjecucion"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormLugarEjecucion"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoLimpiar_Click` — `Me.LUGAREJECUCION.Value = Null`, reset selection (pure UI action, no helper)

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modLugarEjecucionHelper.bas` | **0** (only comments documenting removal) |
| `DoCmd.OpenForm` in `modLugarEjecucionHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_LugarEjecucionHelper.bas` | **0** |
| `Forms(...)` in `Test_LugarEjecucionHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_LugarEjecucionHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_LugarEjecucionHelper.bas` | **0** |
| `CreateObject` in `Test_LugarEjecucionHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out (Phase 2.2 organo-pecal — PR-R3)

The `ByRef p_Form As Object` anti-pattern still lives in `modOrganoContratacionHelper`, `modPECALHelper`, `modSuministradorHelper`, `modUsuariosHelper`, and their *Alta* siblings. Phase 2.2 (PR-R3, next batch of 4 forms) should apply the same rework. The validated pattern is documented in this audit and in `pilot-oficina-programa.md` §2.

---

## 4. ORIGINAL AUDIT (Phase 2.1 / PR-2 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **15** (13 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event, close form | KEEP inline (UI lifecycle adapter, raises event consumed by parent) | Pure orchestration; no testable business logic |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, disable buttons, set RowSource, fetch `m_ObjEntorno.LugaresEjecucion`, filter by `Me.LugarEjecucion`, populate ListBox, fire `ListaFiltrados_Click` | **`LugarEjecucion_Buscar_Filtrar`** (replaces this method) | The biggest extraction; pure-function filter logic |
| 4 | `ComandoAlta_Click` | Private | Clear `m_ObjLugarEjecucionActiva`, close `FormLugarEjecucion` if open, open it, set `m_FormLugarEjecucion` | **`LugarEjecucion_Alta_Abrir`** | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline (delegates to existing legacy `AbrirAyuda`) — no helper extracted | 1 call; pilot does not extract this either |
| 6 | `ComandoBuscar_Click` | **Public** (DELETED) | Call `Filtrar m_Error`, `AvanceCerrar` | **`LugarEjecucion_Buscar_Filtrar`** (wrapper logic stays in caller) | The Public wrapper around `Filtrar`; thin |
| 7 | `ComandoEliminar_Click` | Private | Get selection from ListBox if Nothing, prompt user via `MsgBox`, call `Helper_EntidadCRUD.EliminarEntidadGenerico`, fire `Eliminar` event, refresh | **`LugarEjecucion_Eliminar_Borrar`** | `p_PromptResult` injected for atom assertion |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.LugarEjecucion = Null`, call `ComandoBuscar_Click` | **`LugarEjecucion_Limpiar_Reset`** | Trivial wrapper |
| 9 | `ComandoEditar_Click` | Private | Get selection, set `m_ObjLugarEjecucionActiva`, open `FormLugarEjecucion`, set `m_FormLugarEjecucion` | **`LugarEjecucion_Editar_Abrir`** | Selection-aware open |
| 10 | `LUGAREJECUCION_KeyDown` | Private | Detect Enter key, call `ComandoBuscar_Click` | KEEP inline (delegates) | Form event |
| 11 | `Form_Open` | Private | `Ajustar Me`, configure buttons, set `Caption`, admin check, `cmdElegir.Visible` toggle, call `Filtrar m_Error` | **`LugarEjecucion_Abrir_Inicializar`** | Form lifecycle init |
| 12 | `ListaFiltrados_Click` | Private | Get selection from ListBox, enable `ComandoEditar`/`ComandoEliminar` | **`LugarEjecucion_Seleccionar_Cargar`** | Selection state sync |
| 13 | `ListaFiltrados_DblClick` | Private | If `cmdElegir` visible → `cmdElegir_Click`. Else if `ComandoEditar.Enabled` → `ComandoEditar_Click` | **`LugarEjecucion_DobleClick_AbrirEdicion`** | Dispatcher |
| 14 | `m_FormLugarEjecucion_Alta` | Private (WithEvents) | Reset `m_ObjEntorno.LugaresEjecucion`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |
| 15 | `m_FormLugarEjecucion_Editado` | Private (WithEvents) | Reset `m_ObjEntorno.LugaresEjecucion`, call `ComandoBuscar_Click` | KEEP inline (thin wrapper; calls helper) | WithEvents callback |

**Public methods to delete (rule 11B):** 2 — `Filtrar`, `ComandoBuscar_Click`.

**MsgBox/InputBox occurrences in event handlers:** 1 user-prompt in `ComandoEliminar_Click` (line 188: "¿Desea realmente borrar el lugar de ejecución seleccionado?") — atomizable via `p_PromptResult`. The rest are error-path `MsgBox` calls for trap-only display.

### 1.2 Form controls audit (rule 11A)

**`Me.X` references in `.cls`:** `Caption, cmdElegir, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, LUGAREJECUCION, LugarEjecucion, Name, OpenArgs` (13 unique).

**`Name ="X"` in `.form.txt`:** `ComandoAyuda, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, LUGAREJECUCION, cmdElegir, cmdSalir, lblTitulo` (11 controls). **PASS — `LugarEjecucion` is the field's `.Value` (control name `LUGAREJECUCION` in form.txt).**

Note: the `.cls` accesses `Me.LugarEjecucion` and `Me.LUGAREJECUCION` — the former is the form's control-bound field reference, the latter is the control name. Both refer to the same control. **No mismatch.**

### 1.3 Rule #8 audit (planned helper names)

Planned `LugarEjecucion_*` exports (8) — matches pilot pattern:

```
LugarEjecucion_Alta_Abrir
LugarEjecucion_Buscar_Filtrar
LugarEjecucion_Editar_Abrir
LugarEjecucion_Eliminar_Borrar
LugarEjecucion_Limpiar_Reset
LugarEjecucion_Abrir_Inicializar
LugarEjecucion_Seleccionar_Cargar
LugarEjecucion_DobleClick_AbrirEdicion
```

Audit script (against all 39 planned names across the 4 forms in PR-2):

```powershell
$names = @( 'LugarEjecucion_*', 'GradosClasificacion_*', 'RAC_*', 'RACS_*' )  # 39 names total
Get-ChildItem -Path src\modules, src\classes, src\forms -Filter *.bas -ErrorAction SilentlyContinue
Get-ChildItem -Path src\classes, src\forms -Filter *.cls -ErrorAction SilentlyContinue
# For each: Select-String -Pattern '^(Public)\s+(Sub|Function)\s+(\w+)' and check membership in $names
```

**Result:** PASS — **0 collisions on 39 planned names.**

### 1.4 Inter-form census (rule 11)

```powershell
Get-ChildItem src/forms -Filter 'Form_*.cls' | ForEach-Object {
    $f = $_.Name; $c = (Select-String -Path $_.FullName -Pattern 'Form_Form\w+\.').Count
    if ($c -gt 0) { "$f -> $c llamadas inter-form" }
}
```

**Result for this form:** 0 inter-form callsites (`Forms("FormLugarEjecucion")` is a built-in Forms collection lookup, not a method call on another form's `.cls` — accepted by rule #11 boundary).

### 1.5 Binario health pre-flight (rule A)

`dysflow_dysflow_doctor` result (2026-06-25):

| Check | Status |
|---|---|
| access-db-path | ok — `configured` |
| access-open | ok — `opened` |

Doctor green. Proceeding.

### 1.6 Project smoke test (rule B) — baseline

To be documented after `dysflow_test_vba` runs against the full manifest in step 1.8 (cannot run before user-compile per rule #12).

### 1.7 Module-level declaration ordering — verified

- `modFormInteractionHelper.bas` (existing) — declarations at top, public after private ✅
- `modTestingCoreHelper.bas` (Phase 0, PR-0 merged) — declarations at top, public after private ✅
- New modules will follow the same pattern per `vba-access` §10.1

### 1.8 Atom count target (pre-rework)

8 actions × ~3 scenarios average (happy + sad + adversarial/edge) = **~24 atoms minimum** for this form. Across all 4 forms in PR-2 we target ~32 atoms (some forms share scenarios).

| Action | Atoms planned |
|---|---|
| `LugarEjecucion_Abrir_Inicializar` | 2 (happy-no-admin, sad-form-not-loadable) |
| `LugarEjecucion_Buscar_Filtrar` | 3 (happy, filter-no-match, sad-env-fail) |
| `LugarEjecucion_Seleccionar_Cargar` | 2 (happy-load, empty-selection) |
| `LugarEjecucion_Alta_Abrir` | 2 (happy-open, sad-open-fail) |
| `LugarEjecucion_Editar_Abrir` | 2 (happy-edit, no-selection) |
| `LugarEjecucion_Eliminar_Borrar` | 3 (happy-yes-deleted, sad-no-prompt-cancelled, sad-no-selection) |
| `LugarEjecucion_Limpiar_Reset` | 2 (happy-clear, sad-no-form) |
| `LugarEjecucion_DobleClick_AbrirEdicion` | 3 (dispatch-choose, dispatch-edit, no-edit-disabled) |
| **TOTAL (this form)** | **~19 atoms + RunAll wrapper** |
