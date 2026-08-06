# Pilot Audit — Form_FormOficinasProgramaGestion (Phase 1 PR-1 + REWORK-PILOT 2026-06-26)

> **Change**: `forms-thin-coverage` (Phase 1 — pilot)
> **Form**: `Form_FormOficinasProgramaGestion`
> **Date original**: 2026-06-25
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

---

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 1 original (8 helpers / 23 atoms) | MERGED via PR #32 — `c4542e4 refactor(oficina-programa): thin form + TDD atoms (pilot)` |
| Rework-pilot (5 helpers / 22 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING write-gate authorization |

**Re-work motivation** (user-reported 2026-06-26): el pilot original aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless porque el COM abría un form modal real. Regla rota: `access-vba-e2e-methodology` #1 (forms son wiring UI thin, helpers son datos puros).

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `OficinaPrograma_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `OficinaPrograma_Buscar_Listar` | `(ByVal p_Oficinas As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `OficinaPrograma_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByRef p_Oficinas As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `OficinaPrograma_Eliminar_Borrar` | `(ByRef p_OficinaPrograma As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `OficinaPrograma_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `OficinaPrograma_Alta_Registrar` (was opening `FormOficinaPrograma`)
- `OficinaPrograma_Editar_Actualizar` (was opening `FormOficinaPrograma`)
- `OficinaPrograma_Limpiar_Reset` (was resetting state)

### 2.2 Test atoms (22 atoms, ZERO form references)

22 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `OficinaPrograma_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `OficinaPrograma_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filtered-rows, edge-long-filter, adversarial-filter-with-semicolons, adversarial-semicolon-in-name |
| `OficinaPrograma_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `OficinaPrograma_Eliminar_Borrar` | 5 | sad-cancelled, happy-deleted, sad-no-entity, sad-n-a-sentinel, adversarial-double-delete |
| `OficinaPrograma_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormOficinasProgramaGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormOficinaPrograma"`, `FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormOficinaPrograma"`, `FormularioAbierto`, `Forms(...)`
- `ComandoLimpiar_Click` — `Me.OficinaPrograma.Value = Null`, reset selection (pure UI action, no helper)

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modOficinaProgramaHelper.bas` | **0** (only comments documenting removal) |
| `DoCmd.OpenForm` in `modOficinaProgramaHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_OficinaProgramaHelper.bas` | **0** |
| `Forms(...)` in `Test_OficinaProgramaHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_OficinaProgramaHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_OficinaProgramaHelper.bas` | **0** |
| `CreateObject` in `Test_OficinaProgramaHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `.vbs` / `.ps1` in worktree root | **0** |

### 3.2 Pattern for scale-out

The `ByRef p_Form As Object` anti-pattern still lives in **other** pilot modules (`modLugarEjecucionHelper`, `modPECALHelper`, `modOrganoContratacionHelper`, `modGradosClasificacionHelper`, `modRACHelper`, `modRACSHelper`, `modSuministradorHelper`, `modUsuariosHelper`, and their *Alta* siblings). Phase 2 (next sprints) should apply this same rework to each.

---

## 4. INFRASTRUCTURE NOTES (for the next agent / future iterations)

| Item | Note |
|---|---|
| `.dysflow/project.json` in rework-pilot worktree | MISSING — global config falls back to `00_EXPEDIENTES_staging\Expedientes.accdb`. Per-call `accessPath` / `backendPath` / `destinationRoot` / `projectRoot` overrides required when working in this worktree. |
| Backend `Expedientes_datos.accdb` | NOT present in rework-pilot worktree — must be copied from `00_EXPEDIENTES_staging\Expedientes_datos.accdb` for tests that touch `TbOficinasPrograma` via `CurrentDb`. |
| MCP write gate | ENABLED — `dysflow_import_modules` / `compile_vba` return `MCP_INPUT_INVALID: apply is not allowed`. Either restart MCP server with `--enable-writes`, or add `allowWrites: true` to `.dysflow/project.json` (requires explicit user authorization per AGENTS.md). |
| Previous failed import | `dysflow-22636e52-bf63-4de5-9bee-fb59aa0008f4` (2026-06-26T08:51:51) tried `import_modules` of just the form and failed — the helper wasn't updated yet, so the form's references were stale. |
| Orphan MSACCESS processes | 4× `-Embedding` PIDs were alive at session start; brief file lock cleared mid-session. They do NOT appear in `dysflow_access_force_cleanup_orphaned`. |

---

## 5. NEXT STEPS (for the user)

To complete the rework-pilot PR:

1. **Authorize writes**: either (a) restart the MCP server with `--enable-writes`, or (b) explicitly tell the agent to add `allowWrites: true` to `.dysflow/project.json` in the rework-pilot worktree.
2. Copy `Expedientes_datos.accdb` into the rework-pilot worktree (so `CurrentDb` reaches the test tables).
3. Agent then runs the full sync loop: `import_modules` (helper + test, importMode:Code, compile:true) → `compile_vba` → user compiles the form in VBE → `test_vba filter:Test_OficinaProgramaHelper` → `verify_code` on the 3 modules.
4. Agent commits + opens PR `rework/oficina-programa-pure-data` → `staging`.