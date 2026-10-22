# U07 — Access-Control UAT (issue #280)

## Outcome

Test de aceptación end-to-end para CAP-043..046 que verifica la COMPOSICIÓN de A01 (Authorization) + A02 (CurrentPrincipal) + A03 (SessionState) + A04 (Audit) un servicio de happy-path realista. No re-testea cada capability individual (eso es trabajo de A01..A04); verifica que la cadena completa funciona junta y que el audit es coherente.

## Branch

`test/280-uat-access-control` desde `main @ 6f42d9b` (post A04 merge).

## Budget

+120/-30 = 150 líneas (`additions + deletions`), 1 archivo nuevo.

## Spec

`openspec/changes/expedientes-web-migration/specs/access-control.md` (CAP-043..046, 3 escenarios por cap = 12 escenarios en total). U07 NO cubre los 12 — sería duplicar A01..A04. Cubre los **4 happy-path compuestos**:

1. **Auth + load + bind + audit**: actor con permiso efectivo → A02 load principal → A03 bind session → A01 authorize → A04 emite 4 audit events en orden.
2. **Deny + audit**: actor sin permiso → A02 load (ok) → A03 bind (ok) → A01 deny → A04 emite 3 audit events (load, bind, denied).
3. **Sin principal → fail-closed + audit**: Lanzadera lanza → A02 fails to `principal.failed` → A03 NO bindea → A01 nunca se llama → A04 emite 1 audit event.
4. **Per-context isolation**: dos servicios (dos request contexts) emiten audit events con `actor_id` distintos sin contaminación cruzada.

## Patrón (mirror de los tests A01..A04)

- Dataclass fakes inline: `RecordingAuditLog` (lista events), `FakeLanzadera` (configurable: raises o returns Principal).
- NO `MagicMock`.
- `from __future__` SOLO si lo tienen los archivos de implementación referenciados (en este caso, solo `service.py` lo tiene).
- Tests parametrizados con `@pytest.mark.parametrize` cuando aplique.

## Verificación local pre-PR

```bash
python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/expedientes/application/test_access_control_uat.py -v --no-cov
python3.12 -m ruff check --config app/pyproject.toml tests/expedientes/application/test_access_control_uat.py
python3.12 -m ruff format --check --config app/pyproject.toml tests/expedientes/application/test_access_control_uat.py
python3.12 -m mypy --explicit-package-bases app/src/modules/expedientes/application/authorization/ app/src/modules/expedientes/application/current_principal/ app/src/modules/expedientes/application/session_state/ app/src/modules/expedientes/application/audit/
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main
python3 scripts/check_dry.py --json  # status MUST be "pass"
```

## Cierre

Commit único con mensaje Conventional, push, abrir PR con título `test(exp): add access-control UAT (CAP-043..046, U07)`. PR body: Summary / Hard rules / Tests / Files / Budget. Cierra issue #280.

## Estado

- [ ] worktree `test/280-uat-access-control` creado desde main @ 6f42d9b
- [ ] tests RED escritos
- [ ] verificación de que los 4 happy-path pasan
- [ ] todos los gates verdes
- [ ] commit + push + PR abierto
