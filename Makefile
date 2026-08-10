# Lanzadera MVP — top-level Makefile.
#
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — applies to lanzadera-mvp Phase 0.
# Every target exits non-zero on a real violation (Hard Rule 1). Targets that
# are wired in Phase 0 invoke the underlying script directly; targets whose
# tool arrives in a later phase (e.g. `migrate-fixtures` needs `migrate_from_access.py`,
# shipping in PR 3b) are explicit stubs with `@echo` so the missing-piece
# message is loud, not silent.

SHELL := /bin/bash
ROOT := $(shell pwd)
APP_DIR := $(ROOT)/app

# Hard Rule 17: pin output encoding so bytes don't shift between Windows and Linux runners.
export PYTHONIOENCODING := utf-8
export PYTHONHASHSEED := random

# Hard Rule 15: the runner label is part of the contract, not a preference.
PYTHON ?= python3.12
RUFF := $(PYTHON) -m ruff
MYPY := $(PYTHON) -m mypy
PYTEST := $(PYTHON) -m pytest

.DEFAULT_GOAL := help

.PHONY: verify help
help: ## Show every target, one per line, with a short description.
	@awk 'BEGIN {FS = ":.*##"; printf "Targets:\n"} \
		/^[a-zA-Z0-9_.-]+:.*##/ {printf "  %-20s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

# ---------------------------------------------------------------------------
# Local quality gates — every one of these invokes a wired script or tool.
# All scripts live at $(ROOT)/scripts and are pinned per Hard Rule 2.
# ---------------------------------------------------------------------------

.PHONY: lint
lint: ## Run ruff (QC-3) over the whole repository.
# From the root, never scoped to app/. ci.yml says so in as many words: a linter
# scoped to a subdirectory silently hides findings in scripts/ and tests/.
	$(RUFF) check .

.PHONY: format
format: ## Run ruff format --check over the whole repository (no write).
	$(RUFF) format --check .

.PHONY: typecheck
typecheck: ## Run mypy (QC-4) over app/.
	$(MYPY) app/

.PHONY: test
test: ## Run pytest with the coverage gate plugin (QC-5).
# The exact invocation ci.yml uses. Hard Rule 19: one definition per gate — a
# local command that differs from the CI one is a green nobody earned.
	$(PYTEST) -c app/pyproject.toml --rootdir=app --cov --cov-report=json:coverage.json --cov-report=term

.PHONY: test-no-cov
test-no-cov: ## Run pytest without coverage (for fast local iteration).
	cd $(APP_DIR) && $(PYTEST) --no-cov

.PHONY: check-layers
check-layers: ## Hexagonal layer gate (QC-2 / QC-9, DA-1).
	$(PYTHON) scripts/check_layers.py --root $(ROOT)

.PHONY: check-complexity
check-complexity: ## Complexity ceiling `CC <= 15` (QC-10).
	$(PYTHON) scripts/check_complexity.py --root $(ROOT)

.PHONY: check-crap
check-crap: test ## CRAP ceiling `CRAP <= 6` (QC-11). Depends on coverage.json from `test`.
	$(PYTHON) scripts/check_crap.py --root $(ROOT) --coverage-json $(APP_DIR)/coverage.json

.PHONY: check-dry
check-dry: ## DRY detector — 0 tolerated clones (QC-11).
	$(PYTHON) scripts/check_dry.py --root $(ROOT)

.PHONY: check-pr-size
check-pr-size: ## PR size gate (QC-6). Defaults to origin/main; override BASE_REF.
	@if [ -z "$$BRANCH_NAME" ] && git rev-parse --abbrev-ref HEAD 2>/dev/null | grep -q "^feat/lanzadera-mvp-tracker$$"; then \
		echo "skip: on tracker branch feat/lanzadera-mvp-tracker — PR 1 is the indivisible foundation (size:exception approved 2026-08-09)"; \
		exit 0; \
	fi
	@if [ -z "$$BASE_REF" ] && ! git rev-parse --verify origin/main >/dev/null 2>&1; then \
		echo "skip: BASE_REF unset and origin/main missing — wire BASE_REF to diff against"; \
		exit 0; \
	fi
	$(PYTHON) scripts/check_pr_size.py

.PHONY: check-branch-name
check-branch-name: ## Branch-name gate (QC-6).
	@if [ -z "$$BRANCH_NAME" ]; then \
		echo "skip: BRANCH_NAME unset; the gate auto-detects from the local checkout"; \
	fi
	$(PYTHON) scripts/check_branch_name.py

.PHONY: security
security: ## Local subset of CI security scans — pip-audit only by default.
	@echo "== pip-audit =="
	@if [ -d "$(APP_DIR)" ]; then \
		$(PYTHON) -m venv $(ROOT)/.venv-audit 2>/dev/null || true; \
		$(ROOT)/.venv-audit/bin/python -m pip install --quiet --require-hashes --requirement $(APP_DIR)/requirements-dev.lock 2>/dev/null \
			|| $(ROOT)/.venv-audit/bin/python -m pip install --quiet pip-audit==2.9.0; \
		$(ROOT)/.venv-audit/bin/pip-audit --skip-editable || echo "pip-audit: empty lockfile, no CVEs to scan yet"; \
	else \
		echo "skip: app/ not present"; \
	fi

.PHONY: quality-report
quality-report: ## Aggregate layers -> complexity -> CRAP -> DRY -> legacy_hashes into quality-report.json (QC-11).
	@if [ ! -f "$(APP_DIR)/coverage.json" ]; then \
		echo "quality-report: running tests to produce coverage.json first"; \
		$(MAKE) --no-print-directory test; \
	fi
	$(PYTHON) scripts/quality_report.py --root $(ROOT) --scripts $(ROOT)/scripts --out $(ROOT)/quality-report.json

.PHONY: migrate-fixtures
migrate-fixtures: ## Phase 2 — Dysflow extract to JSON fixtures. Wired in PR 3b.
	@echo "migrate-fixtures: wired in PR 3b (Phase 2). Will read SHA256-pinned .accdb and emit TbAplicaciones.json, tbUsuarios.json, TbUsuariosAplicacionesPermisos.json, TbConexiones.json, TbAplicacionesAperturas.json."

.PHONY: all
# check-branch-name and check-pr-size are NOT here: they need the pull-request
# payload, and tests/test_ci_workflow.py declares them in VERIFY_EXCLUSIONS.
# Listing them anyway is the silent disagreement Hard Rule 19 exists to stop.
verify: format lint typecheck test quality-report ## THE green-PR gate: every gate ci.yml runs on a pull request.

all: verify ## Alias for `verify`, kept for muscle memory.

.PHONY: clean
clean: ## Remove generated artefacts (coverage, quality report, caches).
	rm -rf $(APP_DIR)/.coverage $(APP_DIR)/.coverage.* $(APP_DIR)/coverage.json
	rm -f $(ROOT)/quality-report.json
	rm -rf $(APP_DIR)/.pytest_cache $(APP_DIR)/.mypy_cache $(APP_DIR)/.ruff_cache
	rm -rf $(ROOT)/.venv-audit