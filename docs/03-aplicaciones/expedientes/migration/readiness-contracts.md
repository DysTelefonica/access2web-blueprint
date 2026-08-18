# Readiness Contracts — Expedientes

Contratos de readiness para la migración de Expedientes. Define los perfiles, adapters y flujos que deben estar operativos antes del cutover.

**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-3, D-EXP-8)
**Specs**: `access-control.md` (EXP-CAP-043..046), `integrations.md` (EXP-CAP-051..056)

## Perfiles de autorización

| Perfil | Permisos | Adapter | Readiness check |
|---|---|---|---|
| Administrador | effective_permissions: `expedientes.admin` | `AuthorizationPort` (Lanzadera) | `GET /readyz` devuelve 200 con `authz: ok` |
| Gestor de área | effective_permissions: `expedientes.gestor` | `AuthorizationPort` (Lanzadera) | `GET /readyz` devuelve 200 con `authz: ok` |
| Responsable calidad | effective_permissions: `expedientes.calidad` | `AuthorizationPort` (Lanzadera) | `GET /readyz` devuelve 200 con `authz: ok` |
| Consulta (solo lectura) | effective_permissions: `expedientes.read` | `AuthorizationPort` (Lanzadera) | `GET /readyz` devuelve 200 con `authz: ok` |

**Deny-by-default**: Si `effective_permissions` está ausente o hay error, el `AuthorizationPort` deniega. No hay fallback permisivo (D-EXP-3).

**Fakes deterministas**: En tests/dev, los fakes de `AuthorizationPort` y `CurrentPrincipalPort` son deterministas. El readiness check falla si se enlazan en producción (D-EXP-3).

## Adapters de integración

| Integración | Puerto | Adapter | Readiness check |
|---|---|---|---|
| Lanzadera (identidad) | `CurrentPrincipalPort` | `LanzaderaPrincipalAdapter` | `GET /readyz` devuelve 200 con `lanzadera: ok` |
| Lanzadera (authz) | `AuthorizationPort` | `LanzaderaAuthzAdapter` | `GET /readyz` devuelve 200 con `authz: ok` |
| HPS | `HpsClientPort` | `HpsClientAdapter` | `GET /readyz` devuelve 200 con `hps: ok` |
| AGEDYS | `AgedysClientPort` | `AgedysClientAdapter` | `GET /readyz` devuelve 200 con `agedys: ok` |
| Gestión de Riesgos | `RiesgosClientPort` | `RiesgosClientAdapter` | `GET /readyz` devuelve 200 con `riesgos: ok` |
| No Conformidades | `NcClientPort` | `NcClientAdapter` | `GET /readyz` devuelve 200 con `nc: ok` |
| Notificación | `NotificationDeliveryPort` | `SmtpNotificationAdapter` | `GET /readyz` devuelve 200 con `notification: ok` |
| Documentos | `DocumentStoragePort` | `S3DocumentStorageAdapter` | `GET /readyz` devuelve 200 con `storage: ok` |

**DTOs anticorrupción**: Cada adapter usa DTOs anticorrupción para aislar el dominio de los cambios en las integraciones (D-EXP-8).

**Timeouts**: Cada adapter tiene timeout configurado. Si el timeout se excede, el adapter devuelve error tipado (D-EXP-8).

**Idempotencia**: Cada adapter implementa idempotencia key para evitar duplicados (D-EXP-8).

**Contract tests**: Cada adapter tiene contract tests que verifican el comportamiento esperado (D-EXP-8).

## Flujos de UI

| Flujo | Perfil | Estados | Readiness check |
|---|---|---|---|
| Alta de expediente | Gestor | Vacío → Alta | `GET /expedientes/new` devuelve 200 |
| Edición de expediente | Gestor | Alta → Edición | `GET /expedientes/{id}/edit` devuelve 200 |
| Baja de expediente | Gestor | Edición → Baja | `POST /expedientes/{id}/delete` devuelve 200 |
| Cambio de tipo | Gestor | Alta → Cambio tipo | `POST /expedientes/{id}/change-type` devuelve 200 |
| Consulta de expediente | Consulta | Cualquier estado | `GET /expedientes/{id}` devuelve 200 |
| Búsqueda avanzada | Consulta | Cualquier estado | `GET /expedientes/search` devuelve 200 |
| Exportación E2E | Gestor | Cualquier estado | `POST /expedientes/e2e/export` devuelve 200 |

**HTMX/Alpine**: Los flujos de UI usan HTMX para actualizaciones de regiones y Alpine solo para estado local/accesibilidad (`aria-busy`). Sin SPA ni paridad de pantalla (D-EXP-5).

**Accesibilidad**: Todos los flujos de UI cumplen con `aria-busy` y accesibilidad básica (D-EXP-5).

## Readiness endpoint

```
GET /readyz
```

**Respuesta exitosa** (200):
```json
{
  "status": "ok",
  "authz": "ok",
  "lanzadera": "ok",
  "hps": "ok",
  "agedys": "ok",
  "riesgos": "ok",
  "nc": "ok",
  "notification": "ok",
  "storage": "ok",
  "database": "ok"
}
```

**Respuesta de fallo** (503):
```json
{
  "status": "degraded",
  "authz": "ok",
  "lanzadera": "error",
  "hps": "ok",
  "agedys": "error",
  "riesgos": "ok",
  "nc": "ok",
  "notification": "ok",
  "storage": "ok",
  "database": "ok"
}
```

**Criterio de readiness**: El readiness check devuelve 200 solo si todos los adapters están operativos. Si algún adapter está degradado, el readiness check devuelve 503.

## Verificación

- [ ] Todos los adapters tienen contract tests
- [ ] El readiness endpoint devuelve 200 con todos los adapters operativos
- [ ] El readiness endpoint devuelve 503 si algún adapter está degradado
- [ ] Los fakes deterministas fallan si se enlazan en producción
- [ ] Los flujos de UI cumplen con HTMX/Alpine y accesibilidad
