#!/usr/bin/env bash
# Preflight local canónico: los mismos gates, en el mismo orden, que el job
# `quality` de .github/workflows/ci.yml. Es la ÚNICA lista de comandos (Hard
# Rule 19); el target `verify` del Makefile delega aquí y el test
# `test_make_verify_delegates_to_the_local_preflight` fija que las dos listas
# no se separen. Lo que CI corre y una estación de trabajo no puede —el
# payload del PR, el presupuesto semanal de mutación, los escáneres que
# necesitan Docker o red— queda fuera y se declara en ese mismo test.
set -euo pipefail

ruff format --check --config app/pyproject.toml .
ruff check --config app/pyproject.toml .
mypy --explicit-package-bases app/
python scripts/check_workflows.py
pytest -c app/pyproject.toml --rootdir=app -m "not integration" --cov --cov-report=json:coverage.json --cov-report=term --cov-fail-under=69
python scripts/quality_report.py
