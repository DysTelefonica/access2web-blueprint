# Maintainer prompt — dysflow round 14 — `apply_form_design_plan.plan` has no usable runtime schema or example

## Mode

Documentation/contract regression, medium variant.

## Consumer context

- Consumer: `DysTelefonica/GESTION_RIESGOS`
- Runtime: dysflow MCP `2.19.0`
- Workflow: issue #129 must add four controls and resize three existing controls in one guarded Access form transaction, without guessing `.form.txt` metadata.
- Prior round: round 13 / dysflow #1021 covers a separate `contextId` binding defect.

## Verified symptom

The MCP schema exposes `apply_form_design_plan.plan` only as an opaque object with an empty description and no nested properties. The installed per-tool example is still a TODO scaffold and explicitly says not to guess arguments. A consumer therefore cannot construct the documented atomic multi-operation plan from runtime-owned documentation.

## Literal evidence

`schema({toolName:"apply_form_design_plan"})` returns:

```json
"plan": {
  "type": "object",
  "required": true,
  "description": ""
}
```

Installed file:

```text
C:/Users/adm1/.agents/skills/dysflow-usage/assets/examples/apply-form-design-plan.md:5
TODO: replace this scaffold with a runtime-verified usage contract.
```

The same file states:

```text
Line 18: Do not guess arguments. Read the live tool schema before replacing the TODO payload.
Line 23: TODO: capture the live structured result without inventing fields.
```

The tool description promises operation kinds such as add-control, move-control, rename-control, set-property, delete-control and note, but neither the live schema nor the installed example defines each operation's required fields.

## What already works and must not regress

- `form_list_controls`, `analyze_form_ui` and `render_form_preview` provide safe read-only perception.
- Single-control tools publish typed top-level arguments.
- `apply_form_design_plan` remains safe-by-default and dry-run by default.
- The atomic single-write/single-import gate and rollback behavior must remain intact.

## Required TDD RED tests

1. Contract/schema test: `schema({toolName:"apply_form_design_plan"})` must expose nested `plan.formName` and `plan.operations[]` schemas.
2. Each supported operation kind must publish a discriminated schema with required fields and allowed values.
3. Documentation test must fail while `assets/examples/apply-form-design-plan.md` contains TODO placeholders.
4. Example verification test must execute the published dry-run payload against a fixture form and assert no filesystem/binary mutation.
5. Add an acceptance example with multiple operations in one plan, because atomic batching is the tool's primary value.

## Minimum fix

Publish the complete plan contract through the MCP schema/catalog and replace the TODO scaffold with a runtime-verified dry-run example. Include return fields needed to audit the preview (`mode`, applied/planned operations, advisories, filesystem/import status). Do not require consumers to inspect maintainer source code.

If `commitScope:"source"` is intentionally unsupported for this tool, state that explicitly and document the safe workflow when source changes must be prepared before human compile without importing the Access binary.

## Discipline and guardrails

- Start with RED contract/docs tests.
- Do not weaken write gates, dry-run defaults, strict context, rollback or human-compile discipline.
- Do not invent a second plan shape only for docs; schema, implementation and example must share one contract.
- Use conventional commits; no AI attribution.

## Acceptance output

- PR with contract tests, complete nested JSON schema and runtime-verified example.
- Changelog entry and version bump.
- Release containing the corrected schema/docs bundle.
- PR body includes RED-before/GREEN-after commands and exact results.

## Consumer verification

```text
schema({toolName:"apply_form_design_plan"}) -> nested discriminated plan schema
installed apply-form-design-plan.md -> no TODO placeholders
published dry-run multi-operation example -> ok, mode=dry-run, no source/binary write
```
