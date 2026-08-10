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

.PHONY: help verify format-check lint typecheck test quality mutation clean

help:
	@echo "verify        - THE green-PR gate: every gate ci.yml runs on a pull request"
	@echo "format-check  - ruff format --check (mechanical, fix first)"
	@echo "lint          - ruff check ."
	@echo "typecheck     - mypy over $(PACKAGE)/"
	@echo "test          - pytest with coverage.json, which the CRAP gate consumes"
	@echo "quality       - layers -> complexity -> CRAP -> mutation sites -> DRY, in order"
	@echo "mutation      - the cosmic-ray session + ratchet (slow; CI runs it weekly)"
	@echo ""
	@echo "Run 'make verify' before opening a PR. Its gate list is pinned to ci.yml by"
	@echo "tests/test_ci_workflow.py::test_make_verify_runs_every_local_gate."

# Verification order is Hard Rule 13 and is identical to ci.yml's:
# mechanical fixes first, then tests (which produce coverage.json), then the
# code gates that consume it. Never reorder these; the order is load-bearing.
format-check:
	$(RUFF) format --check

lint:
	$(RUFF) check .

typecheck:
	$(MYPY) $(PACKAGE)/

test:
	$(PYTEST) --cov --cov-report=json:coverage.json

# quality_report.py runs layers -> complexity -> CRAP -> mutation_sites -> DRY
# in one process because their order is pinned in code (GATES), not in YAML
# where a reviewer can swap two steps without noticing.
quality:
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
verify: format-check lint typecheck test quality
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
