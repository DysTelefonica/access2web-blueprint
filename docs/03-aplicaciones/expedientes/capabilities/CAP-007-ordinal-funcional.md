[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)

# CAP-007 — Ordinal funcional

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-007--ordinal-funcional). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Derivación de un ordinal único y reproducible dentro de la jerarquía a partir de identificadores estables; el mismo expediente devuelve el mismo ordinal en cálculos repetidos; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-8: ver §Caché (D70-D71 + DA-8) en [docs/architecture.md](../../../architecture.md) — el ordinal se cachea con `CachePort` por identificador, pero no se trata como contador.
- D70: ver §Caché en [docs/architecture.md](../../../architecture.md) — el ordinal no entra en la prohibición de cachear contadores porque es estable por id.

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-007--ordinal-funcional` (anchor verificado vía `git grep "### Requirement: EXP-CAP-007"` en `openspec/changes/expedientes-web-migration/specs/lifecycle.md`).
- **Catálogo PRD**: fila `EXP-CAP-007 / Derivación ordinal funcional dentro de la jerarquía.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «A. Ciclo de vida»); linaje legacy `CAP-046 (parte no E2E)`, evidencia `VR-hist`, disposición `PRESERVAR`. La consolidación del legacy `CAP-046`/`CAP-047` en una sola capacidad E2E (`EXP-CAP-042`) está documentada en el epígrafe «Corrección del linaje legacy» del mismo ledger.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - Hipótesis — no verificada desde este entorno: el ordinal funcional se deriva como parte del entry de alta (`src/classes/ExpedienteOperaciones.cls:19`) o de `RegistrarExpEntidades` (`:811`), reutilizando `IDExpedienteCalculado` y `ESTADOCalculadoTexto` que el walkthrough de `data/staging/expedientes/docs/capabilities/CAP-001-alta.md` (§3 «Mapa de implementación») menciona pero no ancla por línea. La rama E2E (`EXP-CAP-042`) absorbe la parte que opera sobre `OrdinalE2E` (test `Test_Expediente_OrdinalE2E_DefaultIsNull`, ya verificado runtime).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-006-jerarquia-acuerdos-lotes-basados](CAP-006-jerarquia-acuerdos-lotes-basados.md) | Next: [capabilities-index](index.md)
