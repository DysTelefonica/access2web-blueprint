# Audit — Form_FormUsuariosGestion (Phase 2.3 / PR-4 + REWORK / PR-R4)

> **Change**: `forms-thin-coverage` (Phase 2.3 — PR-4, batch of 5 forms; REWORK / PR-R4)
> **Form**: `Form_FormUsuariosGestion` (SELECT-ONLY Gestion form — no Alta/Editar/Eliminar)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.3 original (5 helpers / ~12 atoms) | MERGED via PR #31 — `c86d460 refactor(sdd): Phase 2.3 / PR-4 — thin 5 forms (suministrador + usuarios + frmBusy + oficina-programa)` |
| Rework / PR-R4 (5 helpers / 14 atoms) | **READY** — source rewritten, anti-pattern removed, binary sync PENDING user paste in VBE |

**Re-work motivation** (user-reported 2026-06-26): el batch Phase 2.3 aceptaba `ByRef p_Form As Object` en los helpers, y los átomos TDD abrían el formulario real con `DoCmd.OpenForm TEST_FORM_NAME`. Esto provocaba la interrupción del VBE en runs headless. Regla rota: `access-vba-e2e-methodology` #1.

---

## 2. REWORK — anti-pattern removed

### 2.1 Helper signatures (5 SELECT-only helpers, ZERO Form params)

| Helper | Signature (post-rework) |
|---|---|
| `Usuarios_Abrir_Inicializar` | `(ByVal p_EsAdministrador As Boolean, ByVal p_HasOpenArgs As Boolean, ByRef p_Error As String) As String` |
| `Usuarios_Buscar_Listar` | `(ByVal p_Usuarios As Object, ByVal p_Filter As String, ByVal p_EstadoFiltro As String, ByRef p_Error As String) As String` |
| `Usuarios_Seleccionar_Cargar` | `(ByVal p_IDUsuario As String, ByVal p_Usuarios As Object, ByVal p_EsAdministrador As Boolean, ByRef p_Error As String) As String` |
| `Usuarios_Limpiar_Reset` | `(ByRef p_Error As String) As String` |
| `Usuarios_DobleClick_AbrirDetalle` | `(ByVal p_DetalleEnabled As Boolean, ByRef p_Error As String) As String` |

**Removed from helpers** (UI orchestration that stays in form per e2e rule #1):
- `Avance "Obteniendo los usuarios Filtrados"` and `AvanceCerrar` UI calls (now in form's `ComandoBuscar_Click`)
- `Set m_ObjEntorno.ColUsuarios = Nothing` cache-invalidation (now in form's `ComandoBuscar_Click`)
- The `Me.Activos` and `Me.USUARIO` control reads (now the form passes primitives)

### 2.2 Test atoms (14 atoms, ZERO form references)

14 atoms + 1 RunAll wrapper. All atoms use `Scripting.Dictionary` stubs. ZERO `DoCmd.OpenForm`, ZERO `Forms(...)`, ZERO `Screen.ActiveForm`, ZERO `Application.Echo`.

Per action:

| Action | Atoms | Coverage |
|---|---|---|
| `Usuarios_Abrir_Inicializar` | 2 | happy-admin-no-args, happy-admin-with-args |
| `Usuarios_Buscar_Listar` | 6 | happy-todos, happy-active-only, happy-inactive-only, edge-empty-collection, happy-text-filter, edge-long-filter |
| `Usuarios_Seleccionar_Cargar` | 2 | happy-admin, sad-empty-selection |
| `Usuarios_Limpiar_Reset` | 1 | happy-cleared (verifies payload.activosDefault="Todos") |
| `Usuarios_DobleClick_AbrirDetalle` | 2 | happy-choose (cmdElegir visible), happy-none (SELECT-only) |

### 2.3 Form (thin UI wiring)

`Form_FormUsuariosGestion.cls` (post-rework) does exactly: read controls → call helper → render JSON result. UI orchestration that stays in the form per rule #1:
- `ComandoLimpiar_Click` — sets `Me.USUARIO = Null` and `Me.Activos = "Todos"` from helper payload (pure UI action)
- `Activos_Change` — re-runs search via `ComandoBuscar_Click` (form-level event)
- `USUARIO_KeyDown` — Enter key handler (form-level event)
- Default `Activos="Sí"` in `Form_Open` (pre-existing legacy behavior, preserved)

### 2.4 Pattern: composite filter (text + activos)

The new helper receives `p_Usuarios` (Dictionary) + `p_Filter` (String) + `p_EstadoFiltro` (String). The form reads `Me.USUARIO.Value` and `Me.Activos.Value` and passes both. The helper applies the legacy composite filter logic:
- Text: case-insensitive substring on `Nombre`
- Activos: `"Sí"` → only active (FechaBaja empty/null/non-date), `"No"` → only inactive (FechaBaja is a date), `"Todos"` → no activos filter

### 2.5 Pattern: serialización for `Seleccionar_Cargar`

The new helper returns a Dictionary entity with `{ID, Nombre}` keys. The form's `Usuarios_RenderResultadoSeleccionar` wraps this Dictionary back into a real `USUARIO` class instance for parent event dispatch (`Seleccionar(m_UsuarioSeleccionado As USUARIO)`).

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `modUsuariosHelper.bas` | **0** |
| `DoCmd.OpenForm` in `modUsuariosHelper.bas` | **0** |
| `DoCmd.OpenForm` in `Test_UsuariosHelper.bas` | **0** |
| `Forms(...)` in `Test_UsuariosHelper.bas` | **0** |
| `Screen.ActiveForm` in `Test_UsuariosHelper.bas` | **0** |
| `CreateObject("Access.Application")` in `Test_UsuariosHelper.bas` | **0** |
| `CreateObject` in `Test_UsuariosHelper.bas` | only `Scripting.Dictionary` (allowed) |
| `MsgBox` in helpers | **0** |
| `Application.Echo` in helpers / tests | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |

---

## 4. ORIGINAL AUDIT (Phase 2.3 / PR-4 preflight, preserved for traceability)

### 0. Form profile — SELECT-ONLY Gestion (4 actions, NOT 8)

`Form_FormUsuariosGestion` is a **SELECT-only** Gestion form for `USUARIO` entities. No Alta/Editar/Eliminar buttons.

**Prefix disambiguation** (entity is `Usuario`, collection is `ColUsuarios`, helper module is `modUsuariosHelper` with prefix `Usuarios_*`).

### 1.1 Handlers audit

Total handlers: **9** (7 Private + 2 Public)

| # | Handler | Scope | Notes |
|---|---|---|---|
| 1 | `Activos_Change` | Private | KEEP inline (UI event) |
| 2 | `cmdElegir_Click` | Private | KEEP inline (UI lifecycle) |
| 3 | `cmdSalir_Click` | Private | Trivial |
| 4 | `Filtrar` | **Public** (DELETED) | Composite filter extraction |
| 5 | `ComandoBuscar_Click` | **Public** → Private | Wrapper stays |
| 6 | `ComandoLimpiar_Click` | Private | Reset text + Activos |
| 7 | `Form_Open` | Private | Lifecycle init |
| 8 | `ListaFiltrados_Click` | Private | Selection state sync |
| 9 | `ListaFiltrados_DblClick` | Private | Dispatcher (only `choose`) |
| 10 | `USUARIO_KeyDown` | Private | Form event |

### 1.2 Form controls audit (rule 11A)

`Me.X` references in `.cls`: `Caption, cmdElegir, ComandoBuscar, ComandoLimpiar, ListaFiltrados, Name, OpenArgs, USUARIO, Activos` (9 unique).

### 1.3 Rule #8 audit (planned helper names)

5 `Usuarios_*` exports (no Alta/Editar/Eliminar because SELECT-only).

### 1.4 Inter-form census (rule 11)

0 inter-form callsites.

### 1.5 Binario health pre-flight (rule A)

Doctor pending — verified before module import in PR-R4.

### 1.6 Project smoke test (rule B)

To be documented after `dysflow_test_vba` runs.

### 1.7 Module-level declaration ordering — verified

Same pattern as Phase 2.1 / 2.2 ✅.

### 1.8 Atom count target — REWORKED

Original target: ~12 atoms across 5 helpers. Reworked target: 14 atoms (richer Buscar_Listar coverage including happy-text-filter and edge-long-filter).

### 1.9 Special considerations

1. **Composite filter**: Activos has 3 values. Preserved in helper.
2. **`Usuarios_Abrir_Inicializar`** defaults `Activos = "Sí"` (NOT "Todos" as `Limpiar_Reset` does). Pre-existing behavior preserved — the form's `Form_Open` sets `Me.Activos = "Sí"` directly via `FormInteraction_EstablecerValorControl` AFTER calling the helper.
3. **`Usuarios_Seleccionar_Cargar`** does NOT enable ComandoEditar/ComandoEliminar (no such buttons). Only loads selection state.
4. **`Usuarios_DobleClick_AbrirDetalle`** only returns `"choose"` (no `"edit"` branch — there's no edit button).
5. **Default `Activos`** differs between Abrir_Inicializar (`"Sí"`) and Limpiar_Reset (`"Todos"`). The new `Limpiar_Reset` returns `payload.activosDefault="Todos"` so the form reads the canonical default from JSON.
