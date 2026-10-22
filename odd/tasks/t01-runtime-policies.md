# T01 — CAP-047..050 Runtime Policies (issue #247)

## Outcome

Implementar `RuntimePolicyService` con 4 métodos (check_readiness, load_config, rebuild_cache, resolve_backend), uno por capability. Cada método emite UN `ExpedienteAuditEvent` con `event_type="runtime.<cap>.recorded"`. Deny-by-default. ~210 líneas total.

## Branch

`feat/247-runtime-policies` desde `main @ 92b9c1c` (post U07 merge).

## Spec a cumplir

`openspec/changes/expedientes-web-migration/specs/runtime.md`:

- **CAP-047 Readiness**: reportar dependencias y config faltantes diagnosticablemente.
- **CAP-048 Configuración tipada**: validar config por entorno, inyectar deps.
- **CAP-049 Caché reconstruible**: invalidar y reconstruir read-model.
- **CAP-050 Binding de backend**: seleccionar backend por despliegue, nunca interacción.

Los 4 specs comparten el patrón deny-by-default + audit + actor/fecha + 3 escenarios (camino feliz / validación / concurrencia).

## Patrón (mirror exacto de A04)

| Archivo | Líneas (~) | Responsabilidad |
|---|---|---|
| `application/runtime_policies/__init__.py` | 10 | re-exports |
| `application/runtime_policies/command.py` | 50 | 4 commands + 4 reports + exception hierarchy |
| `application/runtime_policies/_evidence.py` | 35 | `EXPEDIENTES_APP_ID = 19` + 4 audit event builders |
| `application/runtime_policies/service.py` | 80 | `RuntimePolicyService` con 4 métodos |
| `tests/expedientes/application/test_runtime_policies.py` | 120 | tests con dataclass fakes inline |

## Service contract

```python
class RuntimePolicyService:
    def __init__(self, *, audit_log: Any, deps: RuntimeDeps | None = None) -> None: ...
    
    async def check_readiness(self, command: ReadinessCommand) -> ReadinessReport: ...
    async def load_config(self, command: ConfigCommand) -> RuntimeConfig: ...
    async def rebuild_cache(self, command: CacheCommand) -> CacheReport: ...
    async def resolve_backend(self, command: BindingCommand) -> BindingReport: ...
```

Cada método:
1. Valida el command (deny-by-default en `actor_id` ausente/vacío).
2. Emite UN audit event con `event_type="runtime.<cap>.recorded"`.
3. Devuelve un resultado frozen dataclass con la información del check.

`RuntimeDeps` es un dataclass opcional con `backend: Any`, `cache: Any`, `config_source: Any` que el caller inyecta. Si es None, los métodos devuelven un report "unknown" o "default" (degraded mode).

## Hard rules

- `from __future__ import annotations` SOLO en `service.py` (patrón A01/A02/A03/A04).
- `EXPEDIENTES_APP_ID = 19` SOLO en `_evidence.py` (ratchet dup:79d8c5976f0f).
- Duck-typed `audit_log: Any` (Protocol evitado por gate de capas).
- NO `MagicMock` — dataclass fakes inline.
- Un audit event por llamada (no loops, no double-emit).
- Frozen dataclasses para commands y reports.
- Conventional commit: `feat(exp): add RuntimePolicyService (CAP-047..050, T01)`.

## Verificación local pre-PR

```bash
python3.12 -m pytest -c app/pyproject.toml --rootdir=app tests/expedientes/application/test_runtime_policies.py -v --no-cov
python3.12 -m ruff check --config app/pyproject.toml app/src/modules/expedientes/application/runtime_policies/ tests/expedientes/application/test_runtime_policies.py
python3.12 -m ruff format --check --config app/pyproject.toml app/src/modules/expedientes/application/runtime_policies/ tests/expedientes/application/test_runtime_policies.py
python3.12 -m mypy --explicit-package-bases app/src/modules/expedientes/application/runtime_policies/
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main
python3 scripts/check_dry.py --json  # status MUST be "pass"
```

## Cierre

Commit único, push, abrir PR. PR body: Summary / Hard rules / Tests / Files / Budget. Closes #247.

## Estado

- [ ] worktree `feat/247-runtime-policies` desde main @ 92b9c1c
- [ ] tests RED escritos
- [ ] implementación GREEN
- [ ] 7 gates verdes
- [ ] commit + push + PR
