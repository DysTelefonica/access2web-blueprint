[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-051 — HPS

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-051--hps). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Alta y consulta desde HPS mediante contrato autenticado e idempotente; la clave idempotente evita la duplicación del expediente en reintentos; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- D9-D10, D25: ver §Secret manager en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-051--hps` (anchor verificado vía `git grep "### Requirement: EXP-CAP-051"` en `openspec/changes/expedientes-web-migration/specs/integrations.md`).
- **Catálogo PRD**: fila `EXP-CAP-051 / Alta/consulta desde HPS mediante contrato autenticado e idempotente.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «I. Integraciones»); linaje `inventario`, evidencia `VS/INT`, disposición `MODERNIZAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/classes/ExpedienteOperaciones.cls` — `Registrar` transaccional como camino de alta reutilizable por el sistema autorizado, conforme al linaje compartido con [CAP-001-alta-transaccional](CAP-001-alta-transaccional.md).
  - Hipótesis — no verificada desde este entorno: el PRD README no nombra un módulo de integración HPS en la tabla de mapeo hexagonal, de modo que el contrato HPS y su credencial no tienen ancla `ruta:línea` verificable desde este entorno; la credencial se espera resuelta por `SecretManagerPort` con `EnvSecretManagerAdapter` (D9-D10, D25).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-050-binding-backend](CAP-050-binding-backend.md) | Next: [CAP-052-agedys](CAP-052-agedys.md)
