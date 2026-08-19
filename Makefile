# HARNESS-PROVENANCE: deterministic-quality-harness v1.5 — assets/Makefile
#
# The local entrypoint. `make verify` is THE definition of green (Hard Rule 19):
# one command that runs every gate CI applies to a pull request, in the order
# ci.yml applies them (Hard Rule 13).
#
# Why this file exists at all: a harness whose gate list lives only in ci.yml
# has no local expression of its own contract. Developers and agents run some
# smaller subset, get a green they did not earn, and discover the real gate on
# the runner. The subset always drifts downward, because nothing measures it.
# `tests/test_ci_workflow.py::test_make_verify_runs_every_local_gate` measures it.
#
# Adjust only the configuration block below. Everything after it is mechanism.

# --- configuration ----------------------------------------------------
PYTHON ?= python
PACKAGE ?= app
# ----------------------------------------------------------------------

RUFF ?= $(PYTHON) -m ruff
MYPY ?= $(PYTHON) -m mypy
PYTEST ?= $(PYTHON) -m pytest

.PHONY: help verify format lint typecheck test check-workflows quality-report mutation clean

help:
	@echo "verify           - THE green-PR gate: every gate ci.yml runs on a pull request"
	@echo "format           - ruff format --check --config app/pyproject.toml (mechanical, fix first)"
	@echo "lint             - ruff check --config app/pyproject.toml ."
	@echo "typecheck        - mypy over $(PACKAGE)/"
	@echo "test             - pytest with coverage.json"
	@echo "quality-report   - layers -> complexity -> mutation sites -> DRY, in order"
	@echo "mutation         - the cosmic-ray session + ratchet (slow; CI runs it weekly)"
	@echo ""
	@echo "Run 'make verify' before opening a PR. Its gate list is pinned to ci.yml by"
	@echo "tests/test_ci_workflow.py::test_make_verify_runs_every_local_gate."

# Verification order is Hard Rule 13 and is identical to ci.yml's:
# mechanical fixes first, then tests (which produce coverage.json), then the
# code gates that consume it. Never reorder these; the order is load-bearing.
# --config is load-bearing: the ruff contract (line-length 100, select list,
# tests/fixtures exclude) lives in app/pyproject.toml. Run from the repository
# root without it, ruff finds no configuration and silently uses its own
# defaults — a different line length, a narrower rule set, no exclusions.
format:
	$(RUFF) format --check --config app/pyproject.toml .

lint:
	$(RUFF) check --config app/pyproject.toml .

typecheck:
	$(MYPY) --explicit-package-bases $(PACKAGE)/

# -c app/pyproject.toml --rootdir=app is load-bearing: pytest's contract
# (pythonpath, testpaths, the coverage-gate plugin by its `app.` path) is
# written for rootdir=app. Without them, every test fails with
# ModuleNotFoundError: app. cwd stays at the repo root so coverage.json lands
# where the downstream gates look for it. --cov-fail-under=69 is the current
# floor; pyproject.toml's [tool.coverage.run] omit is what makes that number
# comparable week to week (Hard Rule 14).
test:
	$(PYTEST) -c app/pyproject.toml --rootdir=app --cov --cov-report=json:coverage.json --cov-report=term --cov-fail-under=69

check-workflows:
	$(PYTHON) scripts/check_workflows.py

# quality_report.py runs layers -> complexity -> mutation_sites -> DRY
# in one process because their order is pinned in code (GATES), not in YAML
# where a reviewer can swap two steps without noticing.
quality-report:
	$(PYTHON) scripts/quality_report.py

# verify — THE definition of green (Hard Rule 19).
#
# Deliberately NOT included, each for a reason that is pinned by
# test_make_verify_excludes_what_a_workstation_cannot_run:
#   - check_pr_size / check_branch_name  need the PR payload (merge-base, labels)
#   - check_mutation + cosmic-ray        weekly schedule; far too slow per PR
#   - pip-audit / gitleaks / trivy       Docker- and network-dependent scanners
#
# Adding a gate to ci.yml without adding it here fails the parity test. That is
# the only reason the two lists will still match a year from now.
verify: format lint typecheck test check-workflows quality-report
	@echo "verify: all CI pull-request gates passed."

# Slow, scheduled, and Linux-only in most setups. Kept out of `verify` on
# purpose (Hard Rule 19's exclusion list), available on demand.
mutation:
	cosmic-ray baseline mutation.toml
	cosmic-ray init mutation.toml mutation.sqlite
	cosmic-ray exec mutation.toml mutation.sqlite
	$(PYTHON) scripts/check_mutation.py mutation.sqlite

clean:
	rm -rf .pytest_cache .ruff_cache .mypy_cache coverage.json quality-report.json
