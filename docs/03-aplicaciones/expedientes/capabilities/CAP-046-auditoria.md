[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-046 — Auditoría

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-046--auditoría). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Registro de accesos, cambios, exportaciones e integraciones con usuario y fecha; el registro de auditoría se escribe en la misma transacción que la mutación que lo origina; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- DA-12: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- D55: ver §Audit en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-046--auditoría` (anchor verificado vía `git grep "### Requirement: EXP-CAP-046"` en `openspec/changes/expedientes-web-migration/specs/access-control.md`).
- **Catálogo PRD**: fila `EXP-CAP-046 / Auditoría de accesos, cambios, exportaciones e integraciones.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «G. Identidad, autorización y auditoría»); linaje `CAP-039 + inventario`, evidencia `VS`, disposición `MODERNIZAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `IDUsuarioCreacion` e `IDUsuarioUltimoCambio` — columnas de `TbExpedientes` tipadas `Text(255)` y marcadas ⚠️ (no son FK numéricas) en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Hallazgos críticos del esquema»; la migración agrega FK numérica real.
  - `src/classes/ExpedienteOperaciones.cls` — servicio de casos de uso con `Registrar` transaccional (`BeginTrans` + `CommitTrans` + `Rollback`), donde el audit hereda la atomicidad de la mutación.
  - `IDUsuarioCambio` — la referencia normativa de la FK de auditoría es [CAP-038-trazaibilidad-e2e](CAP-038-trazabilidad-e2e.md) (D132), que fija la regla aplicable a esta capacidad.
  - Hipótesis — no verificada desde este entorno: el linaje `CAP-039 + inventario` declarado por el catálogo PRD no tiene walkthrough propio de auditoría en `data/staging/expedientes/docs/capabilities/`; el canal de auditoría destino se espera en el adapter de audit de `app/src/modules/expedientes/`.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-045-estado-sesion](CAP-045-estado-sesion.md) | Next: [capabilities-index](index.md)
