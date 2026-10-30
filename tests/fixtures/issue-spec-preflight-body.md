Closes #999

# Problema y contexto

El preflight local ejercita el gate issue-spec en modo `local --body-file`. Sin esto, el gate solo se valida dentro del workflow `pr-event` (que necesita `GITHUB_EVENT_PATH` y `GITHUB_TOKEN`); un cambio de canon que rompe el gate solo aparece en CI.

# Evidencia verificable

- `python3 .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py local --body-file tests/fixtures/issue-spec-preflight-body.md` debe devolver `exit 0`.
- `python3 .agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py local --body-file <body-sin-referencia>` debe devolver `exit 1` con el hallazgo de la referencia ausente.

# Alcance y no objetivos

Ejercita el gate issue-spec en el preflight local de BP, en modo `local`. La regla 7 (`--required-sections`) se retiró con team-skills#1103: el canon ya solo comprueba el enlace a una issue y la flag vieja se acepta como no-op obsoleta (team-skills#1121). No cambia el modo `pr-event` del workflow (Refs #960).

# Criterios de aceptación

- El preflight (`scripts/local-preflight.sh`) invoca el gate en modo `local` con un fixture versionado.
- El test `test_make_verify_runs_every_local_gate` sigue verde: el comando añadido es prefijo de una entrada de `REQUIRED_COMMANDS`.
- Los 35 tests del canon (sustituidos en este PR-3) pasan en local.

# Plan de validación

`bash scripts/local-preflight.sh` (gate de preflight, exit 0). La CI del PR corre el comando de workflow en `pr-gate`. (Los 35 tests del canon se retiraron de este repositorio con BP#1040: viven y se ejecutan en `DysTelefonica/team-skills`.)

# Dependencias y riesgos

Refs #960 (cadena). El fixture versionado viaja con el repo (mismo boundary que `tests/fixtures/issue-canonical-794.md`); no se borra. Si el catálogo `ci-pattern` actualiza las 6 secciones del contrato, el fixture se actualiza en la misma PR.
