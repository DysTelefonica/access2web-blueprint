# Bug Report

## Pre-flight Checks
- [x] I have searched existing issues and this is not a duplicate
- [x] I understand this issue needs `status:approved` before a PR can be opened

## Bug Description

Three regressions discovered in the same consumer session immediately after verifying PR #853 (round-7 fix for #852) and PR #854 (round-6 fix for #849) landed in `v2.10.0`:

1. **Bulk-read tools now fail with `VBA_MANAGER_FAILED: ... No se pudo deshabilitar AutoExec/StartupForm mediante DAO ... OpenDatabase ... 'No es una contraseña válida'`.** This is a NEW v2.10.0 safety gate that I did not see in v2.9.2. It affects every tool that enumerates via `OpenDatabase` (e.g. `list_vba_modules`). Targeted tools like `inspect_form({formName: X})` still work. No MCP-level bypass parameter exists; the error message itself suggests a CLI flag `--allow-startup-execution` that is not exposed through the MCP `tools[*].input`.

2. **`import_modules({moduleNames: [...], dryRun:false})` returns a misleading envelope when the per-module operation succeeded.** Outer wrapper reports `VBA_MANAGER_FAILED exit code 1`, while the inner `DYSFLOW_RESULT` has `status:"ok"`, `error:null`, `rollbackApplied:false`, `fallbackUsed:false`. Same call can be either `VBA_MANAGER_FAILED` outer (success but loud failure) or full clean envelope depending on form. This makes it impossible for consumers to programmatically tell whether the op actually succeeded without parsing the nested result.

3. **`MSACCESS.EXE` zombies are not auto-cleaned.** During the smoke for #852, two `MSACCESS.EXE` processes (PIDs 27288, 39736) were spawned by the failed `import_modules` and stayed alive holding the `.accdb` lock. `list_access_operations` returned them as `status:"failed"` but `access_force_cleanup_orphaned({accessPath, no confirmPid})` reported `[]` (refusing to enumerate them because they "aren't orphaned by the criterion"), and only `access_force_cleanup_orphaned({confirmPid: <pid>})` per-PID would retire them. Consumer should not need explicit PIDs to clean up dysflow-spawned processes.

4. **`v2.10.0` release notes body is empty on GitHub.** Without seeing what PR #853 and PR #854 changed, consumers cannot reason about the new gating or any related behavior changes. The current `AdapterVersion: 2.10.0` runtime is a black box.

Together these make v2.10.0 unsafe for production-bulk automation. The consumer (`gestion_riesgos`, branch `staging`) has paused its bulk-import-of-63-forms plan (#115) until a patch release is published.

## Steps to Reproduce

1. Consumer = `gestion_riesgos`, projectId `00-gestion-riesgos-staging`, local `Gestion_Riesgos.accdb` (CSV protected via env `ACCESS_VBA_PASSWORD`, no MSACCESS process alive, no `.laccdb` lock file).
2. Live runtime adapter version: `2.10.0` (verified via `dysflow.get_capabilities`).
3. Reproducer A (gate regression):
   ```js
   await tools.dysflow.list_vba_modules({
     projectId: "00-gestion-riesgos-staging",
     accessPath: "Gestion_Riesgos.accdb",
     typeFilter: "form"
   });
   ```
   Actual: `VBA_MANAGER_FAILED: CRITICAL: No se pudo deshabilitar AutoExec/StartupForm mediante DAO. Detalle: Excepción al llamar a "OpenDatabase" con los argumentos "4": "No es una contraseña válida". Se aborta la apertura para evitar ejecucion no desatendida. Si estás en un entorno controlado de testing y aceptás ejecutar startup code, reintentá con --allow-startup-execution.`
4. Reproducer B (envelope inconsistency — depends on which form name is passed first):
   ```js
   await tools.dysflow.import_modules({
     projectId: "00-gestion-riesgos-staging",
     accessPath: "Gestion_Riesgos.accdb",
     moduleNames: ["Form_FormRiesgoBiblioteca"],
     importMode: "auto",
     dryRun: false,
     verbose: true
   });
   ```
   Actual outer envelope: `VBA_MANAGER_FAILED with exit code 1`. Inner `DYSFLOW_RESULT` payload: `{"module":"Form_FormRiesgoBiblioteca","status":"ok","phase":null,"error":null,"durationMs":903,"rollbackApplied":false,"fallbackUsed":false,"fallbackReason":null}`. Compare to a clean success: `import_modules({moduleNames:["Form_frmSplash"], ...})` returns the same shape but as top-level object `{"result":{"status":"ok","durationMs":645,...},"operation":"import_modules","dryRun":false,"willModifyAccess":true,...}` — no `VBA_MANAGER_FAILED` wrapper.
5. Reproducer C (zombie cleanup):
   ```js
   // After B above, expect zombies held in registry
   await tools.dysflow.list_access_operations({});
   // Returns ops with status:"failed" but live processes (PIDs 27288, 39736) attached to MSACCESS.
   
   await tools.dysflow.access_force_cleanup_orphaned({ projectId: "00-gestion-riesgos-staging", accessPath: "Gestion_Riesgos.accdb" });
   // Actual: [] (no candidates).
   
   // No bypass unless consumer knows the exact PIDs.
   await tools.dysflow.access_force_cleanup_orphaned({ projectId: "00-gestion-riesgos-staging", accessPath: "Gestion_Riesgos.accdb", confirmPid: 27288 });
   // Actual: "ORPHAN_CLEANUP_PID_GONE: PID 27288 is no longer running" (one zombie already exited; the other 39736 required a separate explicit confirmPid).
   ```

## Expected Behavior

Per the three regressions separately:

1. `list_vba_modules` should succeed when `access-open` is OK (as reported by `doctor`). The AutoExec-disable pre-check should only run if actual startup code is detected, and should have an MCP-level bypass parameter.

2. `import_modules` should return consistent envelope shape across all calls:
   - On per-module success: top-level `{result: {status:"ok", ...}, operation:"import_modules", ...}` (the clean shape).
   - On failure: top-level `{error: {code, message, ...}, ...}` with the `code` envelope. Never a `VBA_MANAGER_FAILED exit code 1` outer that wraps a successful inner `status:"ok"`.

3. `access_force_cleanup_orphaned({accessPath, no confirmPid})` should enumerate **all** `MSACCESS.EXE` (or `pwsh.exe`) processes that touch the project's `accessPath`, including ones dysflow itself spawned and that hold an orphan lock. The current policy "only orphaned headless" misses them.

4. `v2.10.0` release should ship with a populated body describing which PRs were included and their user-visible changes.

## Actual Behavior

Covered above.

## Operating System

Windows

## Agent / Client

OpenCode + dysflow MCP

## Shell

PowerShell 7

## Relevant Logs

```text
// Step 3 — bulk-read
VBA_MANAGER_FAILED: CRITICAL: No se pudo deshabilitar AutoExec/StartupForm mediante DAO.
Detalle: Excepción al llamar a "OpenDatabase" con los argumentos "4": "No es una contraseña válida".
Se aborta la apertura para evitar ejecucion no desatendida.
Si estás en un entorno controlado de testing y aceptás ejecutar startup code, reintentá con --allow-startup-execution.

// Step 4 — inconsistent envelope (outer fail, inner ok)
VBA_MANAGER_FAILED: import_modules failed with exit code 1: Accion: Import
Base de datos: [PATH]
Carpeta: [PATH]
[1/1] Importando: Form_FormRiesgoBiblioteca
DYSFLOW_RESULT {"module":"Form_FormRiesgoBiblioteca","status":"ok","phase":null,"error":null,"durationMs":903,"rollbackApplied":false,"fallbackUsed":false,"fallbackReason":null}

// Step 4 (clean success — different envelope shape, no outer wrapper) — compare:
{
  "result": {"module":"Form_frmSplash","status":"ok","phase":null,"error":null,"durationMs":645,"rollbackApplied":false,"fallbackUsed":false,"fallbackReason":null},
  "operation":"import_modules","dryRun":false,"willModifyAccess":true,
  "requestedProjectId":"00-gestion-riesgos-staging","resolvedProjectId":"00-gestion-riesgos-staging",
  "configSource":"repo-config",
  "projectRoot":"C:\00repos\codigo\00_GESTION_RIESGOS_staging",
  "accessPath":"C:\00repos\codigo\00_GESTION_RIESGOS_staging\Gestion_Riesgos.accdb",
  "backendPath":"C:\00repos\codigo\00_GESTION_RIESGOS_staging\Gestion_Riesgos_Datos.accdb",
  "destinationRoot":"C:\00repos\codigo\00_GESTION_RIESGOS_staging\src"
}

// Step 5 — zombie cleanup
ORPHAN_CLEANUP_PID_GONE: PID 27288 is no longer running.
```

## Acceptance criteria

The following five points define the consumer-observable definition of "done" for this issue. Each is testable from the consumer side without requiring internal maintainer knowledge. Until **all five** pass, the consumer considers v2.10.0 unstable and stays on v2.9.2.

1. **Stable bulk-read without autoexec gate when not needed.** Calling `dysflow.list_vba_modules({projectId, accessPath, typeFilter:"form"})` against a project whose `ACCESS_VBA_PASSWORD` env var is set and whose `.accdb` has no live startup macro **must succeed and return the same shape as v2.9.2** (e.g., `{modules: [...], summary: {...}}`). If the safety gate is necessary, the MCP wrapper **must expose** a documented parameter (e.g. `allowStartupExecution: true`) **OR** auto-detect that no startup macro is present and skip the gate.

2. **Single, consistent envelope shape from `import_modules`.** Every `import_modules({dryRun:false, ...})` call **must return one and only one** envelope shape:
   - On full success: `{result: {module, status:"ok", phase, error, durationMs, ...}, operation:"import_modules", willModifyAccess:true/false, ...}`.
   - On per-module failure: `{result: {module, status:"error", error:{code,message}, rollbackApplied, fallbackUsed, fallbackReason, ...}, operation:"import_modules", ...}`.
   - On top-level failure: `{error: {code:"<TYPED_CODE>", message:"<verbatim>"}, ...}` — typed code, not raw `VBA_MANAGER_FAILED exit code N`.
   **Never** an outer `VBA_MANAGER_FAILED` wrapper around a per-module `status:"ok"`. The consumer's smoke reproducer from Step 4 must return the success shape on the same call.

3. **Zombie cleanup enumerates dysflow-spawned processes without consumer-supplied `confirmPid`.** `access_force_cleanup_orphaned({accessPath, no confirmPid})` after a failed `import_modules` must include any `MSACCESS.EXE` process holding the access path with an `laccdb` lock file present, even if the registry shows the op as `status:"failed"`. The current criterion (only "orphaned headless" + "wrong-path" + "Dysflow-owned" filters) is too narrow and misses live-but-failed ops.

4. **Release notes populated.** `v2.10.1` (or whichever release fixes this) ships with a non-empty release body enumerating the user-visible changes — at minimum: what the AutoExec gate does, what trigger conditions gate it, the new MCP parameter(s) if any, and which PRs (#853, #854, this one) are included.

5. **A new test in the dysflow CI suite proves each of items 1–3 above.** The test fixtures must use the live MCP `tools/call` interface (no internal-only hooks). Test names must be informative: `bulk_read_skips_unneeded_autoexec_gate`, `import_modules_envelope_consistent_on_success`, `orphan_cleanup_enumerates_dysflow_spawned_zombies`.

## Additional Context

- Consumer-side rollback safeguards used in this session (which proved the regression): full `Gestion_Riesgos.accdb` byte-restore from `*.bak-20260713-183839-pre-sequential-relink` (SHA `EF489F2F24D7F7A7D3C66D5CA4C584B727CDE8825DB422AC819454629FDFCCA9`, 58,327,040 bytes) verified after every failed smoke.
- This issue is `round-8` of the maintainer-prompt series for `gestion_riesgos`. Antecedents: #849 round-6 (#849 closed by PR #854), #852 round-7 (#852 closed by PR #853). Both landed in `v2.10.0`. Round-8 is about user-visible regressions of v2.10.0 itself, NOT about source bugs.
- Related consumer-side cleanup: per project rule "no generic killers", the consumer avoided `Stop-Process`/`taskkill`/`pkill` and instead used the dysflow API. Once item 3 of acceptance criteria is met, that workaround becomes unnecessary.

## Suggested Version Bump

**patch** → `v2.10.1` (fixes to v2.10.0 are bug-fix level, no breaking changes).

## Reinforcement

Items 1–3 of acceptance criteria were discovered in a single 18-minute consumer smoke session — without per-item tests in CI, every consumer that upgrades to `v2.10.0` will pay the same discovery tax. The 3 reproducer code blocks above are CI-pasteable today; please ensure they become RED tests that fail on v2.10.0 master and pass on the patch.
