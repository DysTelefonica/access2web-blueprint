# H01 — CAP-051 HPS adapter (issue #248)

## Outcome

Implementar `HpsService` para CAP-051 (alta/consulta HPS autenticada e idempotente). Service wrap del `HpsPort` con deny-by-default, audit, y dedupe por clave idempotente vía `IdempotencyPort`. 

## Branch

`feat/248-hps-adapter` desde `main @ b246de9`.

## Spec

`openspec/changes/expedientes-web-migration/specs/integrations.md` CAP-051:

- **Camino feliz**: credencial válida + clave idempotente → alta (o consulta) sin duplicación.
- **Validación/autorización**: deny-by-default en actor ausente, payload inválido, o dependencia no resoluble.
- **Concurrencia/fallo**: dos requests concurrentes con misma clave idempotente → una sola ejecución (la segunda recibe el primer resultado). Fallo del adapter → error diagnosticable, UoW rollback.

## Patrón (mirror A04)

| Archivo | Líneas (~) | Responsabilidad |
|---|---|---|
| `ports/hps.py` | 25 | `HpsPort` Protocol + `HpsSubmission`/`HpsResult` dataclasses |
| `application/hps/__init__.py` | 12 | re-exports |
| `application/hps/command.py` | 45 | `HpsSubmitCommand` + `HpsQueryCommand` + errores + `HpsResult` |
| `application/hps/_evidence.py` | 30 | `EXPEDIENTES_APP_ID = 19` + audit builders |
| `application/hps/service.py` | 90 | `HpsService` con `submit` / `query` |
| `tests/expedientes/application/test_hps.py` | 130 | 8-10 tests con fakes inline |

## Service contract

```python
class HpsService:
    def __init__(self, *, hps: Any, audit_log: Any, idempotency: Any) -> None:
        self._hps = hps
        self._audit_log = audit_log
        self._idempotency = idempotency

    async def submit(self, command: HpsSubmitCommand) -> HpsResult: ...
    async def query(self, command: HpsQueryCommand) -> HpsResult: ...
```

`submit`:
1. Valida actor_id (deny-by-default).
2. Hashea la firma del comando.
3. Busca en `idempotency.get(key)` — si existe, devuelve el resultado previo (sin llamar al adapter).
4. Llama a `hps.submit(...)` (inyectado).
5. Persiste el resultado vía `idempotency.record(...)`.
6. Emite UN audit event.
7. Devuelve `HpsResult`.

`query`: similar pero sin persistencia (las queries son naturalmente idempotentes por referencia).

## Hard rules

- `EXPEDIENTES_APP_ID = 19` solo en `_evidence.py`.
- Duck-typed ports: `Any` (Protocol evitado por gate de capas).
- NO `MagicMock` — dataclass fakes inline (ver A04 tests).
- `from __future__ import annotations` SOLO en `service.py`.
- Una sola emisión de audit por llamada.

## Verificación local pre-PR

```bash
python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/expedientes/application/test_hps.py -v --no-cov
python3.12 -m ruff check --config app/pyproject.toml app/src/modules/expedientes/ports/hps.py app/src/modules/expedientes/application/hps/ tests/expedientes/application/test_hps.py
python3.12 -m ruff format --check --config app/pyproject.toml <same paths>
python3.12 -m mypy --explicit-package-bases app/src/modules/expedientes/application/hps/
python3 scripts/check_layers.py
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main  # ≤400 (plan 175)
python3 scripts/check_dry.py --json  # status MUST be "pass"
```

## Cierre

Commit + push + PR con título `feat(exp): add HpsService adapter (CAP-051, H01)`. Closes #248.

## Estado

- [ ] port escrito
- [ ] service + commands + _evidence + __init__
- [ ] tests RED escritos
- [ ] verificación 7 gates verdes
- [ ] commit + push + PR
