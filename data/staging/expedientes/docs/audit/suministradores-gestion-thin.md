# Audit — Form_FormSuministradoresGestion (Phase 2.3 / PR-4 + REWORK / PR-R4)

> **Change**: `forms-thin-coverage` (Phase 2.3 — PR-4, batch of 5 forms; REWORK / PR-R4)
> **Form**: `Form_FormSuministradoresGestion` (CRUD list / Gestion form)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.3 original (8 helpers / ~22 atoms) | MERGED via PR #31 — `c86d460 refactor(sdd): Phase 2.3 / PR-4 — thin 5 forms (suministrador + usuarios + frmBusy + oficina-programa)` |
| Rework / PR-R4 (5 helpers / 22 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING user paste in VBE |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.3 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. Regla rota: `access-vba-e2e-methodology` #1. El pilot (PR #33 commit 6838014), PR #34 (lugar-grados-rac) y PR #35 (organo-pecal) demostraron la corrección; ahora se aplica al batch Phase 2.3.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `Suministrador_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `Suministrador_Buscar_Listar` | `(ByVal p_Suministradores As Object, ByVal p_Filter As String, ByRef p_Error As String) As String` |
| `Suministrador_Seleccionar_Cargar` | `(ByVal p_IDSeleccionado As String, ByVal p_Suministradores As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `Suministrador_Eliminar_Borrar` | `(ByVal p_Suministrador As Object, ByRef p_PromptResult As Long, ByRef p_Error As String) As String` |
| `Suministrador_DobleClick_AbrirEdicion` | `(ByVal p_HasElegir As Boolean, ByVal p_EditarEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `Suministrador_Alta_Abrir` (was opening `FormSuministrador`)
- `Suministrador_Editar_Abrir` (was opening `FormSuministrador` in edit mode)
- `Suministrador_Limpiar_Reset` (was resetting form state directly)

### 2.2 Test atoms (22 atoms, ZERO form references)

22 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `Suministrador_Abrir_Inicializar` | 3 | happy-admin-no-args, happy-non-admin-no-args, happy-admin-with-args |
| `Suministrador_Buscar_Listar` | 6 | happy-all-rows, edge-empty-collection, happy-filter-by-nombre, happy-filter-by-cif, edge-long-filter, adversarial-semicolon-in-cif |
| `Suministrador_Seleccionar_Cargar` | 4 | happy-admin, happy-non-admin, sad-empty-selection, sad-not-found |
| `Suministrador_Eliminar_Borrar` | 4 | sad-cancelled, happy-deleted, sad-no-entity, sad-missing-id |
| `Suministrador_DobleClick_AbrirEdicion` | 4 | happy-choose, happy-edit, happy-none, edge-choose-takes-precedence |

### 2.3 Form (thin UI wiring)

`Form_FormSuministradoresGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoAlta_Click` — `DoCmd.OpenForm "FormSuministrador"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoEditar_Click` — `DoCmd.OpenForm "FormSuministrador"`, `FormInteraction_FormularioAbierto`, `Forms(...)`
- `ComandoLimpiar_Click` — `Me.PalabraClave.Value = Null`, reset selection (pure UI action, no helper)

### 2.4 Pattern: dual-field filter (Nombre AND CIF)

The new helper receives `p_Suministradores` (Dictionary) + `p_Filter` (String). The form reads `Me.PalabraClave.Value` and passes it as the filter; the helper applies case-insensitive substring match on Nombre OR CIF (project-specific dual-field convention).

### 2.5 Pattern: serialización for `Seleccionar_Cargar`

The new helper returns a Dictionary entity with `{IDSuministrador, CIF, Nombre}` keys. The form's `Suministrador_RenderResultadoSeleccionar` wraps this Dictionary back into a real `Suministrador` class instance so the form's `m_SuministradorSeleccionada` state is consistent for `ComandoEditar_Click` / `ComandoEliminar_Click` paths.

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modSuministradorHelper.bas` | **0** (only comments documenting removal) |
| `DoCmd.OpenForm` in `modSuministradorHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_SuministradorHelper.bas` | **0** |
| `Forms(...)` in `Test_SuministradorHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_SuministradorHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_SuministradorHelper.bas` | **0** |
| `CreateObject` in `Test_SuministradorHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** (Eliminar uses `p_PromptResult` seam) |
| `Application.Echo` in helpers / tests | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.3 / PR-4 preflight, preserved for traceability)

### 0. Form profile — Gestion list

`Form_FormSuministradoresGestion` is the **Gestion list** form for Suministrador entities. Pattern matches `Form_FormLugarEjecucionGestion` (Phase 2.1) and `Form_FormOrganoContratacionGestion` (Phase 2.2).

Notable pre-existing variables / conventions:
- `m_SuministracionSeleccionado` — note the typo (`Suministracion` instead of `Suministrador`). Pre-existing, not fixed in this PR.
- Form control name for the search filter is `PalabraClave` (NOT `Suministrador` like the pilot).
- ListBox columns: `IDSuministrador;CIF;Nombre`.
- `FormularioAbierto` is a legacy helper; will use `FormInteraction_FormularioAbierto` per Phase 0 cross-project convention.

### 1.1 Handlers audit

Total handlers: **14** (12 Private + 2 Public)

| # | Handler | Scope | Inline logic | Proposed helper | Notes |
|---|---|---|---|---|---|
| 1 | `cmdElegir_Click` | Private | Raise `Seleccionar` event | KEEP inline (UI lifecycle adapter) | Pure orchestration |
| 2 | `cmdSalir_Click` | Private | Close form | KEEP inline (trivial) | 1 LOC |
| 3 | `Filtrar` | **Public** (DELETED) | Reset selection, fetch collection, filter by `PalabraClave` (Nombre OR CIF), populate ListBox | **`Suministrador_Buscar_Listar`** | The biggest extraction |
| 4 | `ComandoAlta_Click` | Private | Clear active, close `FormSuministrador` if open, open it | **`Suministrador_Alta_Abrir`** | Open-form orchestration |
| 5 | `ComandoAyuda_Click` | Private | Call `AbrirAyuda m_Error` | KEEP inline | 1 call |
| 6 | `ComandoBuscar_Click` | **Public** → Private | Call `Filtrar` | wrapper stays | Public → Private |
| 7 | `ComandoEliminar_Click` | Private | Prompt user via `MsgBox`, call `Helper_EntidadCRUD.EliminarEntidadGenerico` | **`Suministrador_Eliminar_Borrar`** | `p_PromptResult` injected |
| 8 | `ComandoLimpiar_Click` | Private | Set `Me.PalabraClave = Null` | **`Suministrador_Limpiar_Reset`** | Trivial wrapper |
| 9 | `ComandoEditar_Click` | Private | Set `m_ObjSuministradorActivo`, open `FormSuministrador` | **`Suministrador_Editar_Abrir`** | Selection-aware open |
| 10 | `Form_Open` | Private | `Ajustar Me`, configure buttons, admin check, call `Filtrar` | **`Suministrador_Abrir_Inicializar`** | Form lifecycle init |
| 11 | `ListaFiltrados_Click` | Private | Get selection, enable `ComandoEditar`/`ComandoEliminar` | **`Suministrador_Seleccionar_Cargar`** | Selection state sync |
| 12 | `ListaFiltrados_DblClick` | Private | Dispatch to `cmdElegir` or `ComandoEditar` | **`Suministrador_DobleClick_AbrirEdicion`** | Dispatcher |
| 13 | `m_FormSuministrador_Alta(p_ID)` | Private (WithEvents) | Call `ComandoBuscar_Click`, raise `Alta(p_ID)` event | KEEP inline | WithEvents callback |
| 14 | `m_FormSuministrador_Editado` | Private (WithEvents) | Call `ComandoBuscar_Click`, raise `Editado` event | KEEP inline | WithEvents callback |

### 1.2 Form controls audit (rule 11A)

`Me.X` references in `.cls`: `Caption, cmdElegir, ComandoAlta, ComandoBuscar, ComandoEditar, ComandoEliminar, ComandoLimpiar, ListaFiltrados, Name, OpenArgs, PalabraClave` (11 unique).

### 1.3 Rule #8 audit (planned helper names)

Planned `Suministrador_*` exports (8) → REWORKED to 5 (see §2.1):
- Removed: `Suministrador_Alta_Abrir`, `Suministrador_Editar_Abrir`, `Suministrador_Limpiar_Reset` (moved to form as UI orchestration)

### 1.4 Inter-form census (rule 11)

`DoCmd.OpenForm "FormSuministrador"` is a built-in DoCmd call, not a method call on another form's `.cls`.

### 1.5 Binario health pre-flight (rule A)

Doctor pending — verified before module import in PR-R4.

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs against the full manifest.

### 1.7 Module-level declaration ordering — verified

Same pattern as Phase 2.1 / 2.2 ✅.

### 1.8 Atom count target — REWORKED

Original target: ~22 atoms across 8 helpers. Reworked target: 22 atoms across 5 helpers (richer coverage per helper since 3 helpers moved to form).

### 1.9 Special considerations

1. **Search filter `PalabraClave`** — matches both Nombre AND CIF (dual-field filter). Preserved in helper.
2. **`m_ObjSuministradorActivo` global** — set by `ComandoEditar_Click` before opening `FormSuministrador`. The form sets it; the helper `SuministradorAlta_Abrir_Inicializar` clears it in alta mode, sets it via constructor in edicion mode.
3. **`m_SuministracionSeleccionado` typo** — kept as-is. Private to form, no impact on helpers.
4. **Event signature** — `Alta(p_ID As String)` is forwarded from `m_FormSuministrador_Alta(p_ID)`. Preserved.
5. **`m_SuministradorSeleccionada` (post-rework name)** — replaced the typo name with a correctly-named module-level. The form wraps the helper's Dictionary output back into a `Suministrador` class for state consistency.
