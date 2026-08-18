[← Back to capabilities-index](../capabilities-index.md) · [← Back to DOCS](../../../DOCS.md)

# Expedientes · access-control (D51 / EXP-CAP-043..046)

Resumen a nivel repo del slice D51 del task plan de la epic #170. La fuente de verdad por Requirement es [`openspec/changes/expedientes-web-migration/specs/access-control.md`](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md).

## Capacidades

| ID | Título | Escenario clave |
|---|---|---|
| [EXP-CAP-043](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-043--deny-by-default) | Deny-by-default | Camino feliz: actor autenticado con permiso efectivo → autorización registrada y auditable. |
| [EXP-CAP-044](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-044--principal-lanzadera) | Principal Lanzadera | Camino feliz: Lanzadera disponible → sesión con principal suministrado. |
| [EXP-CAP-045](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-045--estado-de-sesión) | Estado de sesión | Camino feliz: request y sesión del mismo usuario → estado aislado sin globals mutables. |
| [EXP-CAP-046](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-046--auditoría) | Auditoría | Camino feliz: cualquier decisión → registro con actor/fecha y contexto. |

Cada Requirement tiene además escenarios de validación (datos inválidos, permiso ausente, dependencia no resoluble → denegación sin efecto parcial) y de concurrencia o fallo (preservación de invariantes y reintentabilidad). Ver el spec para los G/W/T completos.

## Cómo se aplica a access2web-blueprint

- **D86** (forma hexagonal preservada del legacy): las decisiones de auth viven en `app/src/modules/expedientes/domain/authorization/` y se exponen vía `AuthorizationPort`; los adapters concretos (HTTP middleware, gRPC interceptor, etc.) se inyectan en `di/`.
- **DA-11** (audit en misma transacción que mutación auth): los eventos canónicos (`auth.bootstrap.set_password`, `auth.login.success`, `auth.login.failure`, `app.open`, `global_admins.bootstrap`) se persisten atómicamente con la mutación.
- **D88+D89** (Argon2id `RFC_9106_LOW_MEMORY`, sin `legacy_hash`): los passwords migrados van con `password_hash = NULL`; la helper `password_hash` está pineada con cobertura 100 % (`CRITICAL_HELPERS`, QC-5).
- **D55** (sin telemetría sensible): nunca `ssid`, `bssid`, `coordinates`, `machine_name`, `ip_address` en logs ni en argumentos CLI.

## Core invariants

- **Deny-by-default es ley**: cualquier capacidad sin permiso efectivo explícito deniega. No hay fallback permisivo. La regla se aplica en el adapter más cercano a la entrada (HTTP middleware, gRPC interceptor, CLI flag).
- **Principal siempre de Lanzadera**: la identidad, el `app_id=19` y los permisos por aplicación se consumen desde Lanzadera vía `IdentityPort`; ningún módulo del Expedientes accede directamente a `tbUsuarios`.
- **Sin globals mutables**: el estado de sesión vive en `request.scope` o `contextvars.ContextVar`; nunca en módulos globales. Tests pinean esto con `Test_NoMutableGlobals`.
- **Auditoría inseparable de la mutación**: el INSERT en `audit_log` va en la misma transacción que la mutación auth; si el primero falla, la mutación hace rollback (DA-11). `scripts/check_legacy_hashes.py` pinea este invariante.

## Contributor checklist

- [ ] Si el PR añade un puerto de autorización, el adapter mantiene la signatura del puerto y el dominio no importa el framework concreto (FastAPI Depends, grpc interceptor, CLI flag).
- [ ] Si el PR añade una regla deny-by-default, hay un test que verifica el camino denegado con effective_permissions ausente o inválido.
- [ ] Si el PR modifica la sesión, el estado vive en `request.scope` o `contextvars`; ningún global mutable nuevo.
- [ ] Si el PR toca audit, el INSERT de audit_log comparte transacción con la mutación auth (DA-11).
- [ ] El PR respeta los 5 gates del workflow (`check_branch_name`, `check_pr_size`, `check_workflows`, ruff, mypy).
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, partir por unidad de trabajo o encadenar.

## Lista de comprobación final

- [ ] Las 4 capabilities referencian anchors al spec fuente (`#requirement-exp-cap-NNN`).
- [ ] Las decisiones D86/DA-11/D88/D89/D55 citadas en §Cómo se aplica están vigentes en `docs/architecture.md`.
- [ ] El tono sigue `skills/documentation-alan-style/SKILL.md` §3 + §8 (castellano peninsular formal con usted).
- [ ] Cross-references desde DOCS.md, CODEBASE-GUIDE.md y `docs/03-aplicaciones/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [capabilities-index](../capabilities-index.md) | Next: [runtime](runtime.md)
