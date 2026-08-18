[← Back to capabilities-index](../capabilities-index.md) · [← Back to DOCS](../../../DOCS.md)

# Expedientes · integrations (D53 / EXP-CAP-051..056)

Resumen a nivel repo del slice D53 del task plan. Fuente de verdad: [`openspec/changes/expedientes-web-migration/specs/integrations.md`](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md).

## Capacidades

| ID | Título | Contrato | Sistema destino |
|---|---|---|---|
| [EXP-CAP-051](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-051--hps) | HPS | Outbound JSON/HTTP con identidad compartida y timeout duro | Lanzadera vía XApp handshake (D132) |
| [EXP-CAP-052](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-052--agedys) | AGEDYS | Outbound al servicio externo Agedys (genérico o específico) con logging de intentos y reversibilidad | Servicio externo AGEDYS |
| [EXP-CAP-053](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-053--gestión-de-riesgos) | Gestión de Riesgos | Vinculación conceptual por `IDRiesgo` con orden de migración estricto (D95) | Gestion_Riesgos vía XApp (D132) |
| [EXP-CAP-054](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-054--no-conformidades) | No Conformidades | Vinculación conceptual por `IDNC` con orden de migración estricto (D95) | NoConformidades vía XApp (D132) |
| [EXP-CAP-055](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-055--correo) | Correo | Outbound SMTP/SES vía `NotificationDeliveryPort` con cola-por-tabla (DA-10) | SMTP server / SES |
| [EXP-CAP-056](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-056--documentos) | Documentos | Object storage S3-compatible vía `DocumentStoragePort` con URL prefirmada y borrado lógico primero, físico después | S3-compatible (D16) |

Cada Requirement tiene escenarios de validación (credenciales ausentes, endpoint caído, version mismatch) y de concurrencia o fallo (idempotencia, retry con backoff). Ver el spec.

## Cómo se aplica a access2web-blueprint

- **D86** (forma hexagonal preservada): cada integración vive en `app/src/modules/expedientes/adapters/` y se inyecta vía `di/`. Los puertos (`HpsClientPort`, `AgedysClientPort`, `RiesgosClientPort`, `NcClientPort`, `NotificationDeliveryPort`, `DocumentStoragePort`) se exponen en `ports/`; el dominio no importa los SDKs concretos.
- **D132** (XApp HTTP/JSON handshake cross-app): Lanzadera, Gestion_Riesgos y NoConformidades se consumen vía XApp con timeout duro y retry con backoff. La identidad compartida va en el header `X-Principal` derivado del `IdentityPort`.
- **D14** (esquema por módulo en PostgreSQL): cada app tiene su propio schema (`expedientes`, `lanzadera`, `gestion_riesgos`, `no_conformidades`); las FKs inter-app son conceptuales, no físicas, vía XApp.
- **D16** (S3-compatible para anexos): el `DocumentStoragePort` con `S3CompatibleAdapter` (MinIO en dev, AWS S3 / equivalente en prod).
- **DA-10** (NotificationDeliveryPort con cola-por-tabla): filas en `mail_outbox` con `status='pending'`; dispatcher externo las consume cada ~5 min.

## Core invariants

- **Identidad compartida con timeout duro**: cada llamada cross-app lleva el principal del request y falla explícitamente si el destino no responde en < N segundos (no retry infinito).
- **Borrado lógico antes del físico en documentos**: `DocumentStoragePort.delete` marca el objeto como `pending_purge` y sólo el garbage collector interno lo borra físicamente tras verificación.
- **Correo no bloqueante**: la cola-por-tabla (`mail_outbox`) acepta el mensaje en <10 ms; el dispatcher externo hace el SMTP real. La UI no espera el envío.
- **FKs inter-app siempre via XApp, nunca via SQL directo**: Gestion_Riesgos y NoConformidades se consultan vía `RiesgosClientPort.getById(...)`, no via `getdbLanzadera()`. La acoplamiento directo del legacy es el anti-patrón que D86 viene a romper.

## Contributor checklist

- [ ] Si el PR añade un puerto de integración, el adapter concreto implementa retry con backoff y timeout duro.
- [ ] Si el PR introduce una FK inter-app, va por XApp (puerto), no por SQL directo.
- [ ] Si el PR toca la cola de correo, el dispatcher externo no se acopla a la request (no bloqueante).
- [ ] Si el PR toca documentos, el borrado es lógico primero; el físico lo hace el GC tras verificación.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Lista de comprobación final

- [ ] Las 6 capabilities referencian anchors al spec fuente.
- [ ] Las decisiones D86/D132/D14/D16/DA-10 citadas están vigentes en `docs/architecture.md`.
- [ ] Tono castellano peninsular formal con usted.
- [ ] Cross-references desde DOCS.md y `docs/03-aplicientes/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [runtime](runtime.md) | Next: [uat-cutover-legacy-retirement](uat-cutover-legacy-retirement.md)
