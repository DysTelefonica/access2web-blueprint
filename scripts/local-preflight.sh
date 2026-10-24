#!/usr/bin/env bash
# Preflight local canónico: los gates de pull request del pr-gate de
# .github/workflows/ci.yml en modo local, en su mismo orden, más ruff. Es la
# ÚNICA lista local (Hard Rule 19): el target `verify` del Makefile delega aquí
# y el test `test_make_verify_runs_every_local_gate` fija que no se separe de
# ci.yml. La suite, mypy, la cobertura, la mutación y los escáneres quedan solo
# en CI: son lentos o necesitan Docker, red o el contexto del run, y se declaran
# en VERIFY_EXCLUSIONS de ese mismo test.
set -euo pipefail

ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
python .agents/skills/ci-pattern/assets/pr-size/check_pr_size.py
python .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py local --body-file tests/fixtures/issue-spec-preflight-body.md
python scripts/check_walkthrough_schema.py --strict
