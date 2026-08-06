# Maintainer prompt — DysTelefonica/team-skills round 1 — refresh `dysflow-usage` and `access-form-ui-builder` docs to match dysflow v2.19.0

## Mode

`hygiene` (docs-only refresh of the consumer-facing skill surface), long variant. No runtime code change; the runtime fix is in DysTelefonica/dysflow round 16 (this issue is its docs-side companion).

## Routing — read this first

This issue is filed against **`DysTelefonica/team-skills`** because every file mentioned below lives in `C:\Proyectos\skills\` (remote = `DysTelefonica/team-skills`). The `dysflow-usage` skill and the `access-form-ui-builder` skill are tracked in the team-skills monorepo, not the dysflow runtime repo. Verified at session start:

```text
C:\Proyectos\skills\.git\config -> remote.origin.url = https://github.com/DysTelefonica/team-skills.git
C:\Proyectos\skills\skills\dysflow-usage\.git\config      -> DysTelefonica/team-skills.git
C:\Proyectos\skills\skills\access-form-ui-builder\.git\config -> DysTelefonica/team-skills.git
```

The runtime-contract fix (publish non-opaque nested schemas for `generate_form_design_plan`, `copy_form_ui_pattern`, `verify_form_ui`) is filed separately against `DysTelefonica/dysflow` as round 16. This issue covers the docs that ship with team-skills and cannot live in the dysflow runtime repo.

No `compose_apply`, `compile_vba`, or runtime-vs-binary fix belongs in this round.

## Consumer context

- Consumer: `DysTelefonica/GESTION_RIESGOS` (the staging worktree at `C:/00repos/codigo/00_GESTION_RIESGOS_staging`)
- Runtime fact: `get_capabilities` reports `adapterVersion=2.19.0`, `toolsVisible=89`, `projectId=00-gestion-riesgos-staging`, `status=valid`
- Skills surface at issue-time:
  - `dysflow-usage` last_verified at v2.19.0 (matches runtime)
  - `access-form-ui-builder` last_verified at **v2.7.0** (STALE — 12 minor versions behind)
- Prior team-skills-side issues: none on file. The companion runtime fixes are DysTelefonica/dysflow `#1021` (round 13), `#1022` (round 14), the `form_duplicate_control` round 15 issue, and the round 16 contract publication. Every docs change in this issue aligns to the contract revealed by those runtime fixes.

## Sub-gaps (eight todos, all read-only authoring)

The audit verified fourteen TODO placeholder files under `C:\Proyectos\skills\skills\dysflow-usage/assets/examples/` (≤ 712 bytes each, all carrying the literal string `TODO: replace this scaffold with a runtime-verified usage contract.`). They split into:

**Form-tool examples (13)** — targeted by dysflow round 16 (and `#1022` for `apply-form-design-plan.md`):

1. `apply-form-design-plan.md` — replace with a runtime-verified multi-operation dry-run payload. Round 14 / `#1022` covered the schema publication; the consumer-facing example must mirror that contract. The plan shape now exposes `plan.formName` and `plan.operations[]`; include at least one `add-control`, one `set-property`, and one `move-control` operation in the worked example.
2. `analyze-form-layout.md` — runtime-verified example that returns `analyze_form_layout({sourcePath, scope?}) -> { overlapWarnings, alignmentWarnings, offSectionControls, tabOrderVsVisualOrder }`.
3. `analyze-form-ui.md` — runtime-verified example for `analyze_form_ui` returning semantic roles.
4. `copy-form-ui-pattern.md` — runtime-verified example for the read-only "lift a proven layout from a reference form" verb. Aligns to dysflow round 16.
5. `diff-form-preview.md` — before/after visual diff sample.
6. `form-align-controls.md` — preview-only and committed apply call.
7. `form-distribute-controls.md` — preview-only and committed apply call.
8. `form-set-property.md` — single-property mutation sample.
9. `generate-form-design-plan.md` — runtime-verified dry-run plan. Aligns to dysflow round 16.
10. `map-form-behavior.md` — control-to-handler-to-callpath example, with the `CodeGraphBehaviorEvidence[]` input the skill describes.
11. `render-form-preview.md` — SVG/ASCII preview sample.
12. `verify-form-bindings.md` — `ControlSource`/`RowSource` resolution against the real backend schema.
13. `verify-form-ui.md` — contract + geometry/tab-order/property checks. Aligns to dysflow round 16.

**Database-tool example (1):**

14. `compact-repair.md` — runtime-verified `compact_repair` sample, including its explicit target selector and the `live-tools-must-not-default-to-frontend` reminder surfaced in the dysflow-usage SKILL.

**Stale skill floor (1):**

- Bump `C:\Proyectos\skills\skills\access-form-ui-builder\SKILL.md` frontmatter from `last_dysflow_version: "2.7.0"` / `requires: "dysflow MCP >= 2.7"` to `last_dysflow_version: "2.19.0"` / `requires: "dysflow MCP >= 2.19"`. The skill body content itself already references the post-`#812`/`#816` verbs (`form_align_controls`, `form_distribute_controls`, `analyze_form_layout`, `render_form_preview`) and the `dryRun:true`/`apply:true` convention, so a floor bump plus an explicit "last_verified_v2_19_0" note is sufficient — no body rewrite required.

**Stale contract in `create-form-from-template.md` (1):**

The file currently mixes two flows:
- Line 11 mentions the **clone flow** with `sourceForm` + `targetForm` against the pure `form-source-resolver`.
- Line 19 calls the tool with `specPath: "<repo>/specs/forms/NewForm.json"` (JSON spec flow).

The clone flow is the more common runtime path (per the round 15 consumer evidence: a manual clone with a fresh GUID is the desired behavior; `create_form_from_template` already regenerates the form GUID per `#600`). Rewrite the file so the call block at line 15-24 shows BOTH flows with header comments separating them, and the anti-patterns section explicitly states the path-containment rule applies to `specPath` only.

**Missing examples (consumer parity file list, but no example file):**

Add `.md` files for these runtime verbs the audit called but found no example for:

- `form_duplicate_control` (mirror of the round-15 runtime fix; the docs-side companion MUST use a different sample GUID from the runtime source so the example does not silently regress alongside the runtime fix)
- `form_serialize`
- `compare_form`
- `lint_form_code`

These four are all read-only or plan-only verbs; the example discipline matches the existing per-tool `assets/examples/<tool>.md` skeleton (When to use, Call, Anti-patterns, Result shape, Live verification, Cross-reference).

## Required TDD RED tests (docs side)

1. The literal string `TODO: replace this scaffold with a runtime-verified usage contract.` is absent from every file in `dysflow-usage/assets/examples/`. Today it appears in fourteen files.
2. The literal string `TODO_SOURCE_PATH` is absent from every file in `dysflow-usage/assets/examples/`. Today it appears in the same fourteen.
3. Every per-tool example with a `dryRun`/`apply` flag pair documents the live dry-run payload AND the equivalent apply:true payload.
4. `access-form-ui-builder/SKILL.md` frontmatter `last_dysflow_version` field is `>= 2.19.0` and `requires` is `>= 2.19`.
5. `access-form-ui-builder/SKILL.md` references no verb not present in `dysflow-usage/SKILL.md` (no drift between the two skills).
6. `create-form-from-template.md` separates the two flows (clone + spec) with a header-per-flow layout and no leftover mixed-prose paragraphs.
7. New example files exist for `form_duplicate_control`, `form_serialize`, `compare_form`, `lint_form_code`. Each follows the same skeleton (When to use, Call, Anti-patterns, Result shape, Live verification, Cross-reference).
8. The hard-rule reminder from the dysflow-usage SKILL ("`compact_repair` defaults to the frontend. Pass its explicit target selector only when the intended database is the backend; never rely on path fallback when the target matters.") is documented once in `compact-repair.md`'s Anti-patterns section.

## Minimum fix (six-file PR shape, all docs-only)

1. **Rewrite the 14 TODO examples** listed above using the existing `assets/examples/<tool>.md` skeleton (`When to use`, `Call`, `Anti-patterns`, `Result shape`, `Live verification`, `Cross-reference`). Each rewrite is a runtime-verified dry-run payload unless the verb is write-class, in which case both the dry-run and the `apply:true` payloads appear. Use the verified dysflow v2.19.0 contract for each.
2. **Add 4 new examples**: `form-duplicate-control.md`, `form-serialize.md`, `compare-form.md`, `lint-form-code.md`. Use the round-15 RED probe sample (without the asserted bug) to seed `form-duplicate-control.md`.
3. **Bump `access-form-ui-builder/SKILL.md`** frontmatter to `last_dysflow_version: "2.19.0"` and `requires: "dysflow MCP >= 2.19"`. No body rewrite.
4. **Rewrite `create-form-from-template.md`** so the two flows are visually separated. Keep both pre-existing paths valid; do not invent a third shape.
5. **No DRY dance** between this round and dysflow round 16: the contract changes only when round-16 lands and releases. Until then, every example uses the current round-16 RED schema observations as a fact-of-life disclaimer (a one-liner at the top of each affected example).
6. **No live-runtime execution.** The audit is read-only; the doc rewrites must work from `get_capabilities` and the existing RED probes already in the consumer's staging worktree.

## What already works and must not regress

- The `dysflow-usage` SKILL.md (master skill) is already aligned to v2.19.0. Do not touch its structure.
- The other per-tool examples that are NOT in the TODO list (e.g. `access-cleanup.md`, `composition-patterns.md`, `get-capabilities.md`, `import-modules.md`, etc.) remain unchanged.
- The `access-form-ui-builder` SKILL.md body content (the perceive→act→verify loop) remains intact; only frontmatter and a one-line "last_verified_v2_19_0" note is needed.
- The hard-rule reminder paths and anti-patterns already in `dysflow-usage` are preserved.
- The team-skills git history of `docs/prompts/` (consumer-side archive) is NOT modified.

## Discipline and guardrails

- All changes are docs. No runtime code, no Access binary touch.
- The rewrites must avoid framing the still-TODO parts as if they were final contracts (do not mark schema items as `additionalProperties: false` in prose; defer that to round-16).
- Keep the per-tool skeleton consistent across the 14 rewrites and the 4 new files.
- Do not introduce a third tool-name flavor (e.g. `formSetProperty` / `form_set_property` / `form.set_property`); pick one and stick to it.
- Use conventional commits; no AI attribution.
- Do NOT touch any tracked file outside `dysflow-usage/assets/examples/` and `access-form-ui-builder/SKILL.md` and `dysflow-usage/assets/examples/create-form-from-template.md` for this round.

## Acceptance output

- PR with the 18 example files (14 rewrites + 4 new) and the `access-form-ui-builder` frontmatter bump.
- Changelog entry naming the v2.19.0 sync and listing the example files.
- Version bump of the skills monorepo (minor, not patch — surface change).
- PR body includes the RED-before / GREEN-after for tests 1-8 above.
- Confirmation that no consumer-facing skill other than the two named was modified.

## Quick verification

```text
rg -n "TODO: replace this scaffold with a runtime-verified usage contract\." \
   C:\Proyectos\skills\skills\dysflow-usage\assets\examples | wc -l
  -> 0
rg -n "TODO_SOURCE_PATH" \
   C:\Proyectos\skills\skills\dysflow-usage\assets\examples | wc -l
  -> 0
access-form-ui-builder/SKILL.md last_dysflow_version field
  -> starts with "2." and the minor >= 19
ls dysflow-usage/assets/examples/form-duplicate-control.md \
   dysflow-usage/assets/examples/form-serialize.md \
   dysflow-usage/assets/examples/compare-form.md \
   dysflow-usage/assets/examples/lint-form-code.md | wc -l
  -> 4
```
