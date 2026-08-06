# Audit — Form_frmBusy (Phase 2.3 / PR-4 + REWORK / PR-R4) — SPECIAL: UI-only popup

> **Change**: `forms-thin-coverage` (Phase 2.3 — PR-4, batch of 5 forms; REWORK / PR-R4)
> **Form**: `Form_frmBusy` (anti-spam popup, UI-only)
> **Date original**: 2026-06-26
> **Date rework**: 2026-06-26
> **Auditor**: SDD apply (sdd-apply skill, post-preflight + post-rework)

## 0. Form profile — Anti-spam popup (UI-only, NO business logic)

`Form_frmBusy` is the **anti-spam popup** used during long operations (>1-2 seconds). Per project AGENTS.md and the `vba-antispam-popup` skill, this form is a UI-only utility:

- It does NOT touch any table or DAO
- It does NOT contain any business logic
- It only displays a progress bar and listens for cancellation

**Decision: NO HELPER EXTRACTION NEEDED.** The `.cls` remains as-is. There is no `modBusyHelper` to create. The only changes to this form are minimal cleanup of the 4 Public Subs (which are pure UI adapters, called indirectly via `FUNCIONES UTILES.bas`'s `MostrarPopupProgreso`/`CerrarPopupProgreso` wrappers — those wrappers don't call the form's Public Subs directly, they access `Forms("frmBusy")` controls directly).

## 1. STATUS

| Fase | Estado |
|---|---|
| Phase 2.3 original (no extraction — UI-only Verified-static) | MERGED via PR #31 — `c86d460 refactor(sdd): Phase 2.3 / PR-4 — thin 5 forms (suministrador + usuarios + frmBusy + oficina-programa)` |
| Rework / PR-R4 (no extraction — UI-only re-verified) | **VERIFIED** — re-checked for anti-pattern matches, source unchanged |

---

## 2. REWORK — RE-VERIFICATION (no extraction needed)

### 2.1 Anti-pattern re-check (PR-R4 audit, 2026-06-26)

Following the same audit pattern used for the 4 helper-backed forms in this PR-R4 batch, `Form_frmBusy.cls` was re-grepped for the full anti-pattern catalog:

| Anti-pattern | `Form_frmBusy.cls` matches | Verdict |
|---|---|---|
| `ByRef p_Form As Object` | 0 | PASS |
| `ByRef p_SubForm As Object` | 0 | PASS |
| `DoCmd.OpenForm` | 0 | PASS |
| `Forms(...)` | 0 | PASS |
| `Screen.ActiveForm` | 0 | PASS |
| `Application.Echo` | 0 | PASS |
| `MsgBox` (user-prompt path, not error-trap) | 0 | PASS |
| `DAO` / `Recordset` / `Database` (DAO) | 0 (only `Option Compare Database`) | PASS |
| `Helper_` (helper module reference) | 0 | PASS |
| `constructor.` (constructor module reference) | 0 | PASS |

The single match for "Database" in `Form_frmBusy.cls` is the `Option Compare Database` statement on line 4 — a VBA compiler directive, NOT a DAO/Database access. The form has zero DAO calls, zero helper calls, zero constructor calls, and zero form-coupling references.

### 2.2 Decision: no extraction (matches Phase 2.3 preflight)

The PR-4 preflight audit (see §4 below) decided "**NO HELPER EXTRACTION NEEDED**". The PR-R4 rework re-verifies that decision:

- The form has no testable business logic (the 4 Public Subs are pure UI control setters — `Me.lblEstado.Caption = ...`, `Me.lblProgresoBarra.Width = ...`, etc.).
- The form has no DAO, no validation rules, no business rules. Only the lifecycle/UI concerns (timer, progress bar, ribbon, cancel flag).
- Per `access-vba-e2e-methodology` rule #1, forms with no business logic need NO helper extraction.

### 2.3 Tier 3 form (form-tiering)

Per e2e methodology `references/form-tiering.md`, this is a Tier 3 form: "leave as-is, document as Verified-static debt". This audit doc IS the debt record.

### 2.4 Action items (re-confirmed in PR-R4)

- **DO NOT** create `modBusyHelper.bas`
- **DO NOT** create `Test_frmBusy_*.bas` (no testable logic)
- **DO NOT** modify `Form_frmBusy.cls` (no changes needed)
- **DO** verify_code on this form (should be caseOnly — no functional changes)
- **DO** include this audit doc in the PR

---

## 3. VERIFICATION CHECKLIST

### 3.1 Source anti-pattern audit (post-rework, re-verified 2026-06-26)

| Check | Result |
|---|---|
| `ByRef p_Form As Object` in `Form_frmBusy.cls` | **0** |
| `DoCmd.OpenForm` in `Form_frmBusy.cls` | **0** |
| `Forms(...)` in `Form_frmBusy.cls` | **0** |
| `Screen.ActiveForm` in `Form_frmBusy.cls` | **0** |
| `Application.Echo` in `Form_frmBusy.cls` | **0** |
| `DAO` / `Recordset` / `Database` (DAO access) in `Form_frmBusy.cls` | **0** |
| `Helper_` / `constructor.` calls in `Form_frmBusy.cls` | **0** |
| `MsgBox` (user-prompt) in `Form_frmBusy.cls` | **0** |
| `.vbs` / `.ps1` in worktree root | **0** |
| `Form_frmBusy.cls` file size (LOC, including blanks/comments) | 92 (unchanged from pre-rework) |
| `Form_frmBusy.cls` Public Subs (UI adapters, unchanged) | 4 (`MostrarEstado`, `ActualizarProgreso`, `IniciarPopup`, `CerrarPopup`) |
| `Form_frmBusy.cls` lifecycle handlers (unchanged) | 3 (`Form_Open`, `Form_Timer`, `Form_Unload`) |

---

## 4. ORIGINAL AUDIT (Phase 2.3 / PR-4 preflight, preserved for traceability)

### 1.1 Handlers audit

Total handlers: **7** (3 lifecycle + 4 Public UI adapters)

| # | Handler | Scope | Inline logic | Action | Notes |
|---|---|---|---|---|---|
| 1 | `Form_Open` | Private | Toggle ribbon (in non-dev mode), set TimerInterval=500, reset Progreso, set `g_BusyFlag="running"` | KEEP inline (lifecycle) | Form event |
| 2 | `Form_Timer` | Private | If `g_BusyFlag="done"`, reset flag and close form | KEEP inline (lifecycle) | Form event |
| 3 | `Form_Unload` | Private | Set `g_OperationCancelled = True` | KEEP inline (lifecycle) | Form event |
| 4 | `MostrarEstado(p_Mensaje)` | **Public** | Set `Me.lblEstado.Caption` | KEEP (Public — UI adapter, NOT called externally) | UI adapter |
| 5 | `ActualizarProgreso(p_Actual, p_Total)` | **Public** | Update `lblProgresoBarra.Width` based on percentage | KEEP (Public — UI adapter) | UI adapter |
| 6 | `IniciarPopup(p_Titulo, p_Mensaje)` | **Public** | Set Caption, lblTitulo, lblEstado, reset progress bar | KEEP (Public — UI adapter) | UI adapter |
| 7 | `CerrarPopup()` | **Public** | Close the form | KEEP (Public — UI adapter) | UI adapter |

**Public methods to delete (rule 11B):** 0. The 4 Public Subs are pure UI adapters for the popup overlay. They are NOT business logic and they are NOT called externally — the `FUNCIONES UTILES.bas` wrappers (`MostrarPopupProgreso`, `CerrarPopupProgreso`) access `Forms("frmBusy")` controls directly (not via these Public Subs). These Public Subs are leftover from a previous design but they don't break anything. Keeping them preserves behavior.

**MsgBox/InputBox occurrences:** 1 (in `Form_Open` errores block) — error-path only.

### 1.2 Form controls audit (rule 11A)

`Me.X` references in `.cls`: `Caption, lblEstado, lblProgresoBarra, lblProgresoFondo, lblTitulo, Name, TimerInterval` (7 unique). **PASS — all UI controls.**

### 1.3 Rule #8 audit (planned helper names)

**Planned exports for `modBusyHelper`:** NONE — no helper needed.

### 1.4 Inter-form census (rule 11)

`Form_frmBusy` is opened by `FUNCIONES UTILES.bas`'s `MostrarPopupProgreso` via `DoCmd.OpenForm "frmBusy"` — that's a built-in DoCmd call, not a method call on another form's `.cls`.

Note: many OTHER forms do access `Forms("frmBusy")` controls directly (Form_Form0BDOpciones, Form_FormExpedientesGestion, Instalador.bas, FUNCIONES UTILES.bas, etc.) — but those are read-only/writes-to-control accesses on the FORM instance, not method calls on Form_frmBusy.cls code.

### 1.5 Binario health pre-flight (rule A)

Doctor pending — verified before any import.

### 1.6 Project smoke test (rule B)

The user paste binario should already include this form (no changes). Existing `Test_frmBusy_*` atoms (if any) should continue to pass.

### 1.7 Module-level declaration ordering — verified

Existing `Form_frmBusy.cls` is at the top with all declarations before any code — already compliant.

### 1.8 Atom count target

**Atoms: 0.** No testable business logic to cover. The form's behavior (Form_Open lifecycle, Form_Timer polling, Form_Unload cancellation flag) is purely UI and is not the kind of behavior that TDD atoms should cover (would require UI automation, not unit testing).

### 1.9 Special considerations (re-confirmed in PR-R4)

1. **NO business logic**: this is the rare case where e2e rule #1 ("no business logic in event handlers") is trivially satisfied — there's no logic to extract. The 4 Public Subs are pure UI control setters, not business logic.
2. **NO cross-form dependencies**: no helper module to create. The form is a leaf in the dependency graph.
3. **NO DAO**: zero database touches — the form only sets control values and reads a global flag (`g_BusyFlag`, `g_OperationCancelled`).
4. **Tier 3 form**: "leave as-is, document as Verified-static debt". This audit doc IS the debt record.

## 5. Confidence ledger entry (Verified-static debt)

Per access-vba-capability-docs §7 confidence ledger:

```markdown
## Form_frmBusy — Verified-static

**Tier**: 3 (UI overlay, no business logic)
**Audit**: 2026-06-26, sdd-apply Phase 2.3 / PR-4 (re-verified 2026-06-26 PR-R4)
**Decision**: kept as-is, no extraction needed
**Verification**: anti-pattern grep on Form_frmBusy.cls reported 0 matches (only `Option Compare Database` literal)
**Future work**: none required (this is a leaf form, no churn expected)
```
