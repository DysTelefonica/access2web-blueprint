[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-008 — Hitos temporales

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-008--hitos-temporales). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Mantenimiento de los hitos del expediente con sus reglas temporales (fechas de inicio/fin, dependencias entre hitos); denegación si una transición temporal rompe las reglas del expediente; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-008--hitos-temporales` (anchor verificado vía `git grep "### Requirement: EXP-CAP-008"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-008 / Hitos y sus reglas temporales.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje legacy `CAP-016`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - `src/classes/ExpedienteOperaciones.cls:2033` — entry `Public Function RegistrarHito(...)` (registro de hitos), referenciado en la lista de entries de §3 «Mapa de implementación» de `data/staging/expedientes/docs/capabilities/CAP-001-alta.md`.
  - `src/modules/modExpedienteHitosHelper.bas` — helper de dominio por agregado Hitos (citado en la tabla de descubrimiento de staging de [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)»).
  - Tabla afectada en escritura: `TbExpedientesHitos` (FK = `IDExpediente`).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-007-ordinal-funcional](CAP-007-ordinal-funcional.md) | Next: [CAP-009-modificados-historial](CAP-009-modificados-historial.md)
