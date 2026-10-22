# U08 — Runtime UAT (issue #281)

## Outcome

Test de aceptación end-to-end para CAP-047..050 que verifica la COMPOSICIÓN de los 4 métodos de RuntimePolicyService (T01) sobre un audit log compartido. No re-testea cada método individual (eso es trabajo de T01); verifica que la cadena completa funciona junta y que el audit es coherente.

## Branch

`test/281-uat-runtime` desde `main @ e62d9fb` (post XCUT #703 merge, incluye T01).

## Budget

+110/-25 = 135 líneas (`additions + deletions`), 1 archivo nuevo.

## Spec

`openspec/changes/expedientes-web-migration/specs/runtime.md` (CAP-047..050, 3 escenarios por cap = 12 escenarios en total). U08 NO cubre los 12 — sería duplicar T01. Cubre los **happy-path compuestos**:

1. **Full chain**: check_readiness → load_config → rebuild_cache → resolve_backend → 4 audit events en orden con `capacidad` correcta.
2. **Per-context isolation**: dos servicios sobre el mismo audit log emiten audit events con `actor_id` distintos sin contaminación.
3. **Dependency failure**: `audit_log.append` lanza → `RuntimeDependencyError` chained.
4. **Audit chain integrity**: cada capacidad emite `event_type="runtime.<cat>.recorded"` y `capacidad=EXP-CAP-04X` correcto.
5. **Payload round-trip**: `load_config` y `resolve_backend` surface sus campos typed en `payload`.

## Patrón (mirror de U07)

- Dataclass fakes inline: `RecordingAuditLog` (lista events), `RuntimePolicyService` instanciado con la misma log.
- NO `MagicMock`.
- Composición: una sola instancia del service, llamadas encadenadas sobre el mismo audit log.
- `from __future__ import annotations` SOLO en `service.py` (no en el test).

## Verificación local pre-PR

```bash
python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/expedientes/application/test_runtime_uat.py -v --no-cov
python3.12 -m ruff check --config app/pyproject.toml tests/expedientes/application/test_runtime_uat.py
python3.12 -m ruff format --check --config app/pyproject.toml tests/expedientes/application/test_runtime_uat.py
python3.12 -m mypy --explicit-package-bases tests/expedientes/application/test_runtime_uat.py
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main
python3 scripts/check_dry.py --json  # status MUST be "pass"
```

## Cierre

Commit único con mensaje Conventional, push, abrir PR con título `test(exp): add runtime UAT (CAP-047..050, U08)`. PR body: Summary / Hard rules / Tests / Files / Budget. Closes #281.

## Estado

- [ ] worktree `test/281-uat-runtime` creado desde main @ e62d9fb
- [ ] tests RED escritos
- [ ] verificación de los happy-path compuestos pasan
- [ ] todos los gates verdes
- [ ] commit + push + PR abierto
