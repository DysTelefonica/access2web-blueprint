# Maintainer prompt — dysflow round 13 — `contextId` is bound as `projectId` when `projectId` is omitted

## Mode

Regression/contract bug hunt, medium variant.

## Consumer context

- Consumer: `DysTelefonica/GESTION_RIESGOS`
- Consumer worktree: `C:/00repos/codigo/00_GESTION_RIESGOS_staging`
- Runtime: dysflow MCP `2.19.0`
- Active project config reported by `get_capabilities`: `projectId=00-gestion-riesgos-staging`, `status=valid`, `toolsVisible=89`
- Prior rounds: round 12 covered write-gate UX gaps. This round is a separate optional-field binding defect in a read-only form tool.

## Verified symptom

`form_list_controls` succeeds when `projectId` is explicitly supplied or when both `projectId` and `contextId` are omitted. It fails when `projectId` is omitted but `contextId` is supplied: the runtime reports the `contextId` value as the requested `projectId`.

The tool schema declares `projectId?` and `contextId?` as independent named optional fields. Omitting one must not shift or reinterpret the other.

## Literal reproduction

### Control call that succeeds

```js
await tools.dysflow.form_list_controls({
  projectId: "00-gestion-riesgos-staging",
  contextId: "issue-129-preflight-staging-indicador",
  sourcePath: "C:/00repos/codigo/00_GESTION_RIESGOS_staging/src/forms/Form_FormIndicador.form.txt",
  limit: 5000
});
```

Observed: returns `formName: "FormIndicador"`, `totalCount: 34`.

### RED call

```js
await tools.dysflow.form_list_controls({
  contextId: "issue-129-preflight-main-indicador",
  sourcePath: "C:/00repos/codigo/00_GESTION_RIESGOS/src/forms/Form_FormIndicador.form.txt",
  limit: 5000
});
```

Literal result:

```text
CONFIG_PROJECT_ID_MISMATCH: Requested projectId 'issue-129-preflight-main-indicador' does not match repo config id '00-gestion-riesgos-staging' in [PATH]
```

The requested `projectId` in the error is exactly the caller's `contextId`.

### Second control call that succeeds

```js
await tools.dysflow.form_list_controls({
  sourcePath: "C:/00repos/codigo/00_GESTION_RIESGOS/src/forms/Form_FormIndicador.form.txt",
  limit: 5000
});
```

Observed: returns `formName: "FormIndicador"`, `totalCount: 38`.

## Preliminary fault boundary

Verified at the MCP contract boundary: named optional arguments are not kept independent for this shape. The internal root cause and affected tool set are not yet verified. Audit the shared argument normalization/dispatch path rather than patching only `form_list_controls` if the defect is centralized.

## What already works and must not regress

- Explicit `projectId` plus `contextId` works.
- Omitting both fields works.
- `sourcePath` can read a caller-selected `.form.txt` without COM or filesystem mutation.
- `form_list_controls` returns stable control names, types, geometry and event-binding flags.
- Existing `CONFIG_PROJECT_ID_MISMATCH` validation must remain enforced when a real, explicit mismatching `projectId` is supplied.

## Required TDD RED tests

1. `form_list_controls` with `contextId` present and `projectId` omitted must preserve `contextId` and resolve project identity from config; it must not emit `CONFIG_PROJECT_ID_MISMATCH` naming the context ID.
2. The same call with an explicit mismatching `projectId` must still emit `CONFIG_PROJECT_ID_MISMATCH` for that actual project ID.
3. Add a parameterized contract test over other MCP tools sharing the same optional `projectId`/`contextId` prelude, proving omission does not shift named fields.
4. Assert the operation/diagnostic trace retains the supplied `contextId` unchanged.

## Minimum fix

Correct the shared input binding/normalization so object fields are read by name and optional omission never produces positional shifting. Keep project identity validation unchanged for explicit values. Do not weaken strict-context or write gates.

## Discipline and guardrails

- Start with the RED tests above.
- Do not change consumer source or project config to hide the defect.
- Do not weaken `CONFIG_PROJECT_ID_MISMATCH`.
- Do not introduce positional compatibility that can reinterpret arbitrary later fields as `projectId`.
- Keep the fix scoped to the MCP input/adapter seam unless tests prove a deeper boundary.
- Use conventional commits; no AI attribution.

## Acceptance output

- PR with regression tests and the minimal fix.
- Changelog entry describing the optional-field binding correction.
- Version bump and release containing the fix.
- PR body includes the exact RED-before/GREEN-after command and results.
- Confirmation of whether other tools with the same input prelude were affected.

## Quick verification

```text
get_capabilities -> adapterVersion >= fixed release
form_list_controls({contextId:"ctx-only",sourcePath:"<valid form path>"}) -> succeeds and retains contextId
form_list_controls({projectId:"definitely-wrong",contextId:"ctx",sourcePath:"<valid form path>"}) -> CONFIG_PROJECT_ID_MISMATCH naming definitely-wrong
```
