---
name: deterministic-quality-harness
description: "Trigger: deterministic code quality, ratchet gate, shrink-only baseline, fail-loud gate. Enforce criteria via mechanical self-policing gates."
license: Apache-2.0
metadata:
  author: "ardelperal"
  version: "1.2"
---

## Activation Contract

Load when authoring or auditing CI workflows that enforce code quality; adding a new lint rule, complexity gate, dependency scan, secret scan, or coverage floor; setting or shrinking a ratchet baseline; wiring a quality gate into a project for the first time; diagnosing why a "running" gate never actually fails.

Do not load for generic code style advice, human review checklists, or naming conventions that do not touch a gate.

## Hard Rules

1. **A gate that cannot fail is not a gate.** Never wrap a scanner with `|| true`, `continue-on-error: true`, or a narrow `grep` that swallows real findings. Red CI beats a false negative.
2. **Pin exact versions for tools that govern rules.** `ruff==0.15.21`, never `ruff>=0.6`. Drift graduates rules; the gate enforces the version it was designed against.
3. **Shrink-only baselines.** Every ratchet count decreases only via explicit script edit. New rule: FAIL. New occurrence: FAIL. Lower count: NOTE (lock-in).
4. **Wiring pin via test.** Every gate in `ci.yml` has a test that parses the workflow YAML and asserts the step is present. Untested wiring drifts.
5. **No lint ignore of antipatterns.** Each `B008`-style ignore is a smell; document the reason inline or fix the code. `Annotated[T, Depends(get_x)]` outlives the ignore.
6. **Image scanners pinned by digest.** `python:3.12-slim-bookworm@sha256:...`. Tags mutate silently; digests pin byte-for-byte content.
7. **No `--ignore-vuln` proliferation.** Each entry corresponds to an open issue; removing the entry closes the issue. Allowlists are shrink-only by ID.
8. **Tests carry the contract via `pytest_sessionfinish`.** `config.exitstatus = 1` does not change exit code; mutate `session.exitstatus`. Verify with `pytester`.
9. **Pure-layer boundaries enforced by AST.** Hexagonal purity uses `scripts/check_layers.py` with `PURE_LAYERS` and `ALLOWED_IMPORTS`; regex breaks on whitespace, comments, and refactors.
10. **Document drift immediately.** When a documented convention diverges from code, fix one or the other in the same PR that detects the divergence. Drift compounds.
11. **Ownership is declared and segregated.** Every gate names the actor that runs it and what that actor explicitly does not own. The author of a change is never the judge of its own gate. Unnamed ownership means the writer self-certifies — the least deterministic arrangement available.
12. **Absolute ceiling first, ratchet second.** Every gate declares a fixed threshold plus the mandatory action taken when it is crossed (`CRAP <= 6`; a file above 100 mutation sites is split before handoff). A ratchet is the ramp toward that ceiling, never the destination: every `BASELINE` records its target value and target date. A gate whose verdict depends on the distribution of other code — `top-N` selection — is not a gate, because the same function passes or fails based on its neighbours.
13. **Verification runs in a fixed order and fixes between steps.** Declare the sequence once, never reorder it, and resolve everything a step reports before running the next. Order is load-bearing: a duplication fix reintroduces complexity, so complexity is re-checked after duplication, not before. Same code plus a different fix order must not yield a different final state.
14. **The testable/untestable boundary is a versioned, gated artifact.** Modules that open GUIs, drive external devices, emit system errors, or hang under automation are declared untestable and excluded from coverage, mutation, complexity, and duplication tooling. Keep that boundary as thin as possible. Without a declared denominator a coverage floor measures a different thing every week.
15. **Pin the environment, not just the tools.** Runner image by exact label (`ubuntu-24.04`, never `ubuntu-latest`), every GitHub Action by commit SHA, every dependency by lockfile with hashes. Pinning `ruff` while resolving the transitive graph at install time leaves the gate green today and red on Tuesday with no code change.
16. **Every gate publishes an indicator, not only a verdict.** Pass/fail says whether to merge; the number says whether the codebase is getting better or worse, and it is what a ratchet's target date is measured against. Each gate emits a machine-readable envelope (`gate`, `status`, `indicators`, `ceilings`, `findings`); one aggregator merges them into a single report and renders it where reviewers already look. An indicator with no declared ceiling and no declared good direction is decoration.
17. **A gate must depend on nothing but the code.** Wall-clock time, `PYTHONHASHSEED`, locale-dependent output encoding, network reachability, filesystem iteration order, and parallel-execution interleaving all make the same commit yield different output. Close each one explicitly: report the commit instead of a timestamp, hash with `hashlib` instead of `hash()`, pin the output encoding, sort every collection before reporting. Two runs over one commit must be byte-identical, or "the report changed" stops meaning "the code changed".

## Decision Gates

| Need | Action |
|------|--------|
| Add a lint rule with known noise | Open a policy issue first; choose block, fix, or ratchet before adding to `select`. |
| Existing project carries noise from day one | Adopt shrink-only ratchet with explicit `BASELINE` snapshot. |
| Test of a gate cannot run (empty fixtures) | Remove the gate or add fixtures; empty gate equals false guarantee. |
| Gate runs locally but not in CI | Wire it or delete the local script; CI is the source of truth. |
| PR exceeds 400-line review budget | Split unless `size:exception` is recorded in the PR body with reason. |
| New dependency introduces CVE | File issue, link to allowlist entry, add to `pip-audit --ignore-vuln`. |

## Execution Steps

0. **Start from `assets/`, never from this prose.** Copy the templates listed under Assets, keep their provenance lines, and adjust only the marked configuration blocks. Re-deriving a gate by hand is how two instances of the same harness end up disagreeing.
1. **Inventory the gate surface.** List every `make` target, `ci.yml` step, `scripts/check_*.py`, and pre-commit hook. Confirm each exits `1` on a real violation.
2. **Pin versions exactly.** Add tool versions to `pyproject.toml` or `.tool-versions`. Reject `>=` for any tool that governs rules.
3. **Wire each gate into CI.** Add the step in `.github/workflows/ci.yml` and a test in `tests/test_ci_workflow.py` that parses the workflow and asserts presence of the step name.
4. **Choose BASELINE policy.** Snapshot current count per rule for legacy projects. Greenfield starts empty: any occurrence fails.
5. **Add the smoke test for the gate itself.** A test invokes the gate with a known violation and asserts exit code `1`. Without it, regressions slip in silently.
6. **Document and audit quarterly.** Update the quality gates doc with name, mechanism, exit code, `BASELINE`, override path. Quarterly: confirm no `|| true` slipped in, `BASELINE` counts only decreased or held steady.

## Output Contract

Return:
- Gate surface inventory (gate, file, exit code on failure, `BASELINE`).
- Wiring pin tests added or verified.
- Pinned tool versions.
- Any `BASELINE` change with reason and source issue.
- Any drift between docs and code, fixed in the same PR.

## Assets

`assets/` ships the reference implementation. Instantiate from it — never re-derive a workflow or
a gate script from this prose, because prose re-derived twice produces two different harnesses,
and a quality harness that is not itself reproducible cannot make anything else reproducible.

| Asset | Destination | Makes executable |
|---|---|---|
| `assets/ci.yml` | `.github/workflows/ci.yml` | Rules 1, 4, 13, 15 |
| `assets/scripts/check_layers.py` | `scripts/check_layers.py` | Rules 9, 12, 14 |
| `assets/scripts/check_complexity.py` | `scripts/check_complexity.py` | Rule 12 (absolute ceiling, never `top-N`) |
| `assets/scripts/check_crap.py` | `scripts/check_crap.py` | Rule 12, and the complexity-vs-coverage axis |
| `assets/scripts/check_dry.py` | `scripts/check_dry.py` | Rule 17 (deterministic clone detection) |
| `assets/scripts/quality_report.py` | `scripts/quality_report.py` | Rules 13, 16 |
| `assets/scripts/check_pr_size.py` | `scripts/check_pr_size.py` | Review-budget decision gate |
| `assets/scripts/check_branch_name.py` | `scripts/check_branch_name.py` | Naming discipline |
| `assets/tests/test_ci_workflow.py` | `tests/test_ci_workflow.py` | Rule 4, and re-asserts 1, 13 and 15 |
| `assets/tests/test_gate_smoke.py` + `assets/tests/fixtures/**` | `tests/` | Execution Step 5, Rule 17 |
| `assets/pyproject.fragment.toml` | merged into `pyproject.toml` | Rules 2, 14, 15 |

### Ceilings that ship

| Gate | Ceiling | Where it comes from |
|---|---|---|
| layers | 0 violations | The architecture decision the project already made |
| complexity | `CC <= 15` per function, global | Conventional; the cheap early signal |
| CRAP | `CRAP <= 6` per function | `swarm-forge` `cleaner.prompt`, verbatim |
| DRY | 0 duplicate blocks of 5+ statements | Greenfield default; ratchet it for legacy |
| PR size | 400 changed lines | Review-budget decision gate |

**Read this before adopting**: `CRAP <= 6` collapses to `CC <= 6` at full coverage, so it dominates
the complexity ceiling of 15 — a function at complexity 7 cannot pass however well tested it is.
That is upstream's number and it is what makes the harness demand small functions rather than
merely well-tested large ones. Raising it is a legitimate local decision; making it silently is not.

## Indicators

Verdicts gate the merge. Indicators tell you which way the codebase is moving, and they are what
a ratchet's target date is measured against. Every gate emits the same envelope:

```json
{"gate": "crap", "status": "pass",
 "indicators": {"max_crap": 4.0, "functions_over_ceiling": 0, "line_coverage_pct": 91.2},
 "ceilings": {"max_crap": 6.0, "functions_over_ceiling": 0},
 "findings": [{"file": "app/x.py", "line": 12, "detail": "..."}]}
```

`scripts/quality_report.py` merges them into `quality-report.json` and renders the table into the
CI job summary. The published set:

| Indicator | Ceiling | Good direction | Meaning |
|---|---|---|---|
| `layers.violations` | 0 | lower | Architecture boundary crossings |
| `complexity.max_complexity` | 15 | lower | Highest cyclomatic complexity of any function |
| `complexity.functions_over_ceiling` | 0 | lower | How many functions are over |
| `crap.max_crap` | 6 | lower | Worst complexity-against-coverage score |
| `crap.line_coverage_pct` | — | higher | Statements the suite actually executes |
| `dry.duplicate_groups` | 0 | lower | Distinct duplicated blocks |
| `dry.duplicated_ratio_pct` | — | lower | Duplicated share of all statements |
| `pr_size.changed_lines` | 400 | lower | Reviewable lines in this change |

## Determinism checklist

Every source of nondeterminism the harness closes, and how. Work through this list when adding a
gate — a gate that fails any row will eventually produce two different verdicts for one commit,
and the first time it does, the team stops trusting all of them.

| Source | Closed by |
|---|---|
| Tool version drift | Exact pins (Rule 2) |
| Transitive dependency resolution | Hash-checked lockfile (Rule 15) |
| Runner image drift | Exact label, never `*-latest` (Rule 15) |
| Action tag mutation | Commit-SHA pins, asserted by `test_ci_workflow.py` (Rule 15) |
| Scanner image tag mutation | Digest pins (Rule 6) |
| Wall-clock time in output | Report the commit SHA instead (Rule 17) |
| `PYTHONHASHSEED` salting | `hashlib.sha256` over a canonical dump, never `hash()` (Rule 17) |
| Locale-dependent output encoding | `_pin_output_encoding()` in every gate (Rule 17) |
| Filesystem iteration order | `sorted()` on every walk and every report (Rule 17) |
| Fix-order sensitivity | Fixed gate order pinned in code by `quality_report.GATES` (Rule 13) |
| Verdict depending on other code | Absolute ceilings, never `top-N` (Rule 12) |
| Coverage denominator drift | Declared testable/untestable boundary (Rule 14) |
| A ratchet that never ends | `BASELINE` entries carry `target` and `target_date` (Rule 12) |

Every asset carries a `HARNESS-PROVENANCE` line naming the skill version that produced it. Keep
it: it is what lets a later audit separate an intentional local edit from silent drift. Read
`assets/README.md` before copying anything, and re-resolve the action SHAs and scanner digests —
they are correct as of the version stamp, not forever.

## References

- `docs/calidad-de-codigo-y-ci.md` — quality gates single source in `access2web-blueprint`.
- `docs/auditoria-harnesses-clean-code.md` — Clean Code and swarm-forge mapping with Phase 2 proposals.
- `https://github.com/unclebob/swarm-forge` — role-segregated agent harnesses. Rules 11–14 are ported from `constitution/articles/engineering.prompt` (`main`) and `roles/{cleaner,architect,hardender,QA}.prompt` (`six-pack`); the tmux/Babashka/worktree orchestration is not imported.
- **Deliberate divergence from swarm-forge.** Its `engineering.prompt` requires installing the *latest upstream* version of every governing tool and forbids cached or vendored copies. Rule 2 requires the opposite. Upstream optimizes for tool freshness; this harness optimizes for a reproducible verdict, which is the stated goal. Do not "fix" rule 2 toward upstream.
- `https://github.com/ardelperal/APAP_WEB` — live reference implementation (issues #380, #381, #393, #424, #428, #436, #437, #441, #442, #443 are the canonical examples).
