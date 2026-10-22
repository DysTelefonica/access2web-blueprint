# A04 — CAP-046 Auditoría (issue #246)

## Outcome

Implementar `AuditService` para CAP-046 con 4 categorías: access / change / export / integration. Cada llamada emite **un solo** `ExpedienteAuditEvent` con `event_type` prefijado `audit.<categoria>.recorded`. Deny-by-default en `actor_id` ausente o vacío. 4 archivos fuente + 1 test, ~315 líneas, dentro del budget de 400.

## Branch

`feat/246-audit` desde `main @ 93e41b9`. A03 (#720) y A02 (#719) siguen abiertos pero A04 no depende de ellos: el `AuditLogPort` ya existe en `app/src/modules/expedientes/ports/audit_log.py` (con `ExpedienteAuditEvent` y `ChangeRecord`).

## Spec a cumplir

`openspec/changes/expedientes-web-migration/specs/access-control.md` líneas 58–62:

> ### Requirement: EXP-CAP-046 — Auditoría
> El sistema MUST registrar accesos, cambios, exportaciones e integraciones con usuario/fecha. Debe preservar invariantes, auditar actor/fecha y aplicar deny-by-default.
>
> #### Scenario: camino feliz
> GIVEN actor autenticado y canal de auditoría disponible
> WHEN realiza acceso, cambio, exportación o integración
> THEN se registra usuario y fecha de la operación, con salida determinista y auditable.
> #### Scenario: validación y autorización
> GIVEN datos inválidos, permiso ausente o dependencia no resoluble
> WHEN solicita auditoría
> THEN deniega sin efecto parcial, devuelve error diagnosticable y conserva auditoría.
> #### Scenario: concurrencia o fallo
> GIVEN dos solicitudes concurrentes o un fallo de dependencia durante auditoría
> WHEN se procesa la operación
> THEN se preservan invariantes, no se duplica ni se pierde evidencia y queda reintentable.

## Patrón (mirror exacto de A02/A03)

| Archivo | Líneas (~) | Responsabilidad |
|---|---|---|
| `application/audit/__init__.py` | 1 | marker |
| `application/audit/command.py` | 40 | `AuditCommand` (frozen), `AuditResult` (frozen), jerarquía de errores, `EXPEDIENTES_APP_ID = 19` |
| `application/audit/_evidence.py` | 55 | builders: `access_recorded_event`, `change_recorded_event`, `export_recorded_event`, `integration_recorded_event` + helpers para `audit.<cat>.failed` |
| `application/audit/service.py` | 70 | `AuditService` con 4 métodos (record_access/change/export/integration), duck-typed `audit_log: Any` |
| `tests/expedientes/application/test_audit.py` | 150 | tests async con fakes dataclass inline (sin `MagicMock`) |

## Tests mínimos (cubrir los 3 escenarios del spec)

- **Camino feliz**: `record_access` con actor válido emite 1 evento `audit.access.recorded`, mismo actor, app_id 19, payload con `target_id` y `category=access`; idem para change/export/integration.
- **Validación**: `actor_id is None` → `AuditAuthorizationError`; `actor_id == UUID(int=0)` → idem; canal de auditoría caído (`append` lanza) → `AuditDependencyError`, NO se duplica ni pierde evidencia (la mutación del caller debe rollback).
- **Concurrencia**: dos `AuditService` sobre el mismo `audit_log` no se interfieren; `record_access` simultáneo produce 2 eventos distintos (cada uno con su `id`).
- **Invariantes**: `AuditCommand`/`AuditResult` son frozen dataclass; el módulo no tiene globals mutables fuera de `EXPEDIENTES_APP_ID` (rastreable con `inspect.getsource`).

## Hard rules

- Sin `MagicMock` — usar dataclass fake inline (`RecordingAuditLog` con lista de eventos capturados).
- Una sola emisión por llamada — no loops, no double-audit.
- `EXPEDIENTES_APP_ID = 19` vive solo en `_evidence.py` (no duplicar en `command.py` para evitar el ratchet DRY dup:79d8c5976f0f que mordió a A03).
- Duck-typed `audit_log: Any` con Protocol evitado (gate de capas `scripts/check_layers.py`).
- Conventional commit: `feat(exp): add AuditService adapter (CAP-046, A04)`.
- ≥ 85% coverage global; rama `application/audit/` con coverage local ≥ 95% (CRITICAL_HELPERS).

## Verificación local pre-PR

```bash
python3 scripts/check_layers.py
python3 -m pytest tests/expedientes/application/test_audit.py -v --no-cov
python3.12 -m ruff check app/src/modules/expedientes/application/audit/ tests/expedientes/application/test_audit.py
python3.12 -m ruff format --check app/src/modules/expedientes/application/audit/ tests/expedientes/application/test_audit.py
python3 -m mypy app/src/modules/expedientes/application/audit/
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main
```

## Cierre

Commit único con mensaje Conventional, push de la rama, abrir PR apuntando a `main`. PR body en castellano peninsular formal siguiendo la plantilla de #719/#720 (Summary / Hard rules / Tests / Files). Vincular a #246 en el body: `Closes #246` si A03 ya mergeó (no todavía), o `Refs #246` si todavía está abierto.

## Estado

- [ ] worktree `feat/246-audit` creado
- [ ] tests RED escritos
- [ ] implementación GREEN
- [ ] REFACTOR + ruff + mypy
- [ ] commit y push
- [ ] PR abierto
