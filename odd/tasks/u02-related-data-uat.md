# U02 — Related-Data UAT (issue #262)

## Outcome

Test de aceptación end-to-end para CAP-008..014 (related-data vertical) que verifica los 7 servicios R01..R07 (`register_hito`, `register_modificado`, `anexos`, `register_anualidad`, `register_responsable`, `register_juridica`, `register_suministrador`) componiendo sobre un audit log compartido.

## Branch

`test/262-uat-related-data` desde `main @ ec4c694` (post cleanup batch).

## Budget

+150/-25 = 175 líneas, 1 archivo.

## Spec

`openspec/changes/expedientes-web-migration/specs/related-data.md` (CAP-008..014, 3 escenarios por cap).

U02 cubre la composición happy-path + 1 failure-mode para cada uno de los 7 verticales. No re-testea cada servicio individualmente (eso es trabajo de R01..R07).

## Patrón (mirror U07)

- 1 dataclass fake `_AuditLog` compartida entre todas las pruebas.
- Tests parametrizados donde aplique (1 happy-path por cap = 7 tests, 1 failure = 7 más = 14 total).
- NO `MagicMock`.
- Idioma: castellano peninsular en docstrings (per `documentation-alan-style`).

## Hard rules

- Test-only — sin cambios en código de producción.
- `EXEDIENTES_APP_ID = 19` no se referencia (no emite audit events propios; los servicios ya lo hacen).
- Conventional commit: `test(exp): add related-data UAT (CAP-008..014, U02)`.

## Verificación local pre-PR

```bash
python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/expedientes/application/test_related_data_uat.py -v --no-cov
python3.12 -m ruff check --config app/pyproject.toml tests/expedientes/application/test_related_data_uat.py
python3.12 -m ruff format --check --config app/pyproject.toml tests/expedientes/application/test_related_data_uat.py
python3.12 -m mypy --explicit-package-bases tests/expedientes/application/test_related_data_uat.py
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main
python3 scripts/check_dry.py --json  # status MUST be "pass"
```

## Cierre

Commit + push + abrir PR con título `test(exp): add related-data UAT (CAP-008..014, U02)`. PR body: Summary / Hard rules / Tests / Files / Budget. Closes #262.

## Estado

- [ ] worktree `test/262-uat-related-data` desde main
- [ ] tests RED escritos
- [ ] 7 gates verde
- [ ] commit + push + PR
- [ ] merge + close issue + delete WT
