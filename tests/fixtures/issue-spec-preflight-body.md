Closes #999

# Problema y contexto

El preflight local ejercita el gate issue-spec en modo `local --body-file` con la lista de secciones declaradas en `scripts/required-sections.json`. Sin esto, el gate solo se valida dentro del workflow `pr-event` (que necesita `GITHUB_EVENT_PATH` y `GITHUB_TOKEN`); un cambio de canon que rompe la regla 7 solo aparece en CI.

# Evidencia verificable

- `python3 .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py local --body-file tests/fixtures/issue-spec-preflight-body.md --required-sections scripts/required-sections.json` debe devolver `exit 0`.
- `python3 .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py local --body-file <body-sin-una-seccion> --required-sections scripts/required-sections.json` debe devolver `exit 1` con la sección ausente en el hallazgo.

# Alcance y no objetivos

Adopta la regla 7 (`--required-sections`) en el preflight local de BP. No cambia el modo `pr-event` del workflow (Refs #960, PR-2 cablea `--required-sections` en `ci.yml:140`).

# Criterios de aceptación

- El preflight (`scripts/local-preflight.sh`) invoca el gate en modo `local` con un fixture versionado y la configuración de secciones.
- El test `test_make_verify_runs_every_local_gate` sigue verde: el comando añadido es prefijo de una entrada de `REQUIRED_COMMANDS`.
- Los 35 tests del canon (sustituidos en este PR-3) pasan en local.

# Plan de validación

`pytest -c app/pyproject.toml --rootdir=app tests/test_check_issue_spec.py` (35 tests) y `bash scripts/local-preflight.sh` (gate de preflight, exit 0). La CI del PR corre los 35 tests en el job `unit` y el comando de workflow en `pr-gate`.

# Dependencias y riesgos

Refs #960 (cadena). El fixture versionado viaja con el repo (mismo boundary que `tests/fixtures/issue-canonical-794.md`); no se borra. Si el catálogo `ci-pattern` actualiza las 6 secciones del contrato, el fixture se actualiza en la misma PR.
