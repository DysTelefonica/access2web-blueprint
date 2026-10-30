[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-009 — Modificados e historial

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-009--modificados-e-historial). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Registro de los cambios de campo del expediente y de su historial reciente (los últimos N); auditoría inmutable por actor/fecha; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- D27: ver §Decisiones arquitectónicas D-<n> cross-cutting en [docs/architecture.md](../../../architecture.md) — logs estructurados (no tablas TbLog), extensivo al historial reciente.

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-009--modificados-e-historial` (anchor verificado vía `git grep "### Requirement: EXP-CAP-009"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-009 / Modificados, cambios de campo e historial reciente.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje legacy `CAP-017, CAP-039` (consolidación), evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - `src/classes/ExpedienteOperaciones.cls:2165` — entry `Public Function RegistrarModificado(...)` (registro del cambio de campo), referenciado en la lista de entries de §3 «Mapa de implementación» de `data/staging/expedientes/docs/capabilities/CAP-001-alta.md`.
  - Hipótesis — no verificada desde este entorno: helper `src/modules/modExpedienteEntidadesHelper.bas` cubre también modificados como sub-caso de entidades-derivadas (citado en la tabla de descubrimiento de staging de [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)»).
  - Tabla afectada en escritura: `TbExpedientesModificados` (FK = `IDExpediente`).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-008-hitos-temporales](CAP-008-hitos-temporales.md) | Next: [CAP-010-anexos-referencias](CAP-010-anexos-referencias.md)
