[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)

# CAP-006 — Jerarquía acuerdos-lotes-basados

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-006--jerarquía-acuerdos-lotes-basados). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Mantenimiento de la jerarquía acuerdo marco → lote → basado con padre válido e integridad autorreferente; denegación si introduce ciclos o rompe la integridad; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- D70: ver §Caché en [docs/architecture.md](../../../architecture.md) — la materialización de la jerarquía no se cachea como contador.

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-006--jerarquía-acuerdos-lotes-basados` (anchor verificado vía `git grep "### Requirement: EXP-CAP-006"` en `openspec/changes/expedientes-web-migration/specs/lifecycle.md`).
- **Catálogo PRD**: fila `EXP-CAP-006 / Jerarquía de acuerdos marco, lotes y basados, incluida integridad padre-hijo.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «A. Ciclo de vida»); linaje legacy `CAP-045 (basados/lotes)`, evidencia `VS`, disposición `PRESERVAR`. La consolidación del legacy `CAP-045` en dos capacidades (basados/lotes → `EXP-CAP-006`, jurídicas → `EXP-CAP-013`) está documentada en el epígrafe «Corrección del linaje legacy» del mismo ledger.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - Hipótesis — no verificada desde este entorno: la implementación consolida el comportamiento del legacy `CAP-045` (basados/lotes) en `src/classes/ExpedienteOperaciones.cls` junto al entry `Registrar` (`RegistrarAlta` y `RegistrarExpEntidades` ya catalogadas para `CAP-001` y `CAP-002`). No existe walkthrough detallado en `data/staging/expedientes/docs/capabilities/` para este consolidación; cualquier `ruta:línea` adicional debe confirmarse al abrirse el adapter PostgreSQL correspondiente (issue futura, p. ej. `R07` proveedor-UTE).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-003-baja-condicionada](CAP-003-baja-condicionada.md) | Next: [CAP-007-ordinal-funcional](CAP-007-ordinal-funcional.md)
