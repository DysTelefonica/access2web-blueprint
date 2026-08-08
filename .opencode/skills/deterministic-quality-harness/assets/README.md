# Assets — reference implementation of the harness

These files are **templates**, not the live gates. The live copies belong in the consuming
repository (`scripts/`, `tests/`, `.github/workflows/`). Two live copies of the same script is
drift, and drift is Hard Rule 10.

## Instantiation contract

1. Copy the file into the consuming repo at the path listed below.
2. Keep the `HARNESS-PROVENANCE` line at the top of every copied file. It records which version of
   this skill produced it, so a later audit can tell an intentional local edit from silent drift.
3. Adjust only the clearly marked configuration block at the top. Everything below it is mechanism.
4. Run `pytest tests/test_gate_smoke.py` immediately. A gate that has never been observed failing is
   not yet a gate (Execution Step 5).

| Asset | Destination in the consuming repo |
|---|---|
| `ci.yml` | `.github/workflows/ci.yml` |
| `scripts/check_layers.py` | `scripts/check_layers.py` |
| `scripts/check_complexity.py` | `scripts/check_complexity.py` |
| `scripts/check_crap.py` | `scripts/check_crap.py` |
| `scripts/check_dry.py` | `scripts/check_dry.py` |
| `scripts/quality_report.py` | `scripts/quality_report.py` |
| `scripts/check_pr_size.py` | `scripts/check_pr_size.py` |
| `scripts/check_branch_name.py` | `scripts/check_branch_name.py` |
| `tests/test_ci_workflow.py` | `tests/test_ci_workflow.py` |
| `tests/test_gate_smoke.py` | `tests/test_gate_smoke.py` |
| `tests/fixtures/**` | `tests/fixtures/**` |
| `pyproject.fragment.toml` | merged into `pyproject.toml` |

## What each asset enforces

| Asset | Hard Rule it makes executable |
|---|---|
| `ci.yml` | 1 (no `\|\| true`, no `continue-on-error`), 4 (every gate wired), 15 (runner and actions pinned) |
| `check_layers.py` | 9 (AST, not regex), 12 (ratchet with target and expiry date), 14 (declared testable boundary) |
| `check_complexity.py` | 12 (absolute global ceiling — never `top-N`) |
| `check_crap.py` | 12, plus the complexity-against-coverage axis; fails closed without coverage data |
| `check_dry.py` | 17 (sha256 over a normalised AST — never `hash()`, which is salted per process) |
| `quality_report.py` | 13 (gate order pinned in code), 16 (indicators, not just verdicts) |
| `check_pr_size.py` | Decision Gate "PR exceeds 400-line review budget" |
| `check_branch_name.py` | Naming discipline, allowlisted defaults only |
| `test_ci_workflow.py` | 4 (wiring pin), and re-asserts 1, 13 and 15 mechanically |
| `test_gate_smoke.py` | Execution Step 5 (each gate observed exiting 1 on a real violation), 17 |

## Resolving the action pins

`ci.yml` ships with real commit SHAs resolved at authoring time. Re-resolve them before adopting,
and on every action upgrade:

```bash
gh api repos/actions/checkout/releases/latest --jq .tag_name
gh api repos/actions/checkout/git/ref/tags/<tag> --jq .object.sha
```

Never move an action back to a floating tag to "fix" a failure. Hard Rule 15 exists because a
floating tag turns a green gate red with no code change.

## Running the gates locally

```bash
pytest --cov --cov-report=json:coverage.json   # CRAP consumes coverage.json
python scripts/quality_report.py               # layers -> complexity -> CRAP -> DRY, in order
```

`quality_report.py` writes `quality-report.json` and prints the indicator table. Any single gate
also runs standalone, with `--root` to point it elsewhere and `--json` for its raw envelope:

```bash
python scripts/check_crap.py --root . --json
```

## Deliberate scope limits

- **Mutation testing does not ship here.** It is the one remaining measurement that verifies the
  tests actually assert something, and CRAP is a proxy for it, not a replacement: CRAP trusts the
  coverage number, and coverage counts a line as covered whether or not any assertion looked at
  it. Upstream gives the affordable recipe — differential mutation against a manifest, one file
  at a time. Do not read this absence as "not needed"; read it as the next thing to build.
- **No coverage floor gate ships here.** Coverage is published as an indicator, but the floor
  belongs to `pytest --cov-fail-under` so that one tool owns one verdict — and it is meaningless
  until the consuming repo declares its testable/untestable boundary (Hard Rule 14).
- **`CRAP <= 6` dominates `CC <= 15`.** At full coverage CRAP equals complexity, so no function
  above complexity 6 can pass. That is upstream's number, kept deliberately. If it is too strict
  for your codebase, raise `MAX_CRAP` explicitly and record why — do not discover it by accident.
