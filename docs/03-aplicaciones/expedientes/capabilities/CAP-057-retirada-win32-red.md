[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-057 — Retirada Win32 red

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-057--retirada-win32-red). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Sustitución de la resolución de red Win32 por un puerto de infraestructura con health/readiness; el diagnóstico de red se obtiene del puerto, nunca del mecanismo legacy; preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-057--retirada-win32-red` (anchor verificado vía `git grep "### Requirement: EXP-CAP-057"` en `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md`).
- **Catálogo PRD**: fila `EXP-CAP-057 / Resolución de red Win32.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «J. Mecanismos legacy retirados o reemplazados»); linaje `CAP-054`, evidencia `—`, disposición `RETIRAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - Hipótesis — no verificada desde este entorno: el linaje `CAP-054` declarado por el catálogo PRD no tiene walkthrough propio en `data/staging/expedientes/docs/capabilities/`; la sustitución por puerto de infraestructura y health/readiness se espera como adapter de red en `app/src/modules/expedientes/adapters/`.
  - Retirada condicionada: sólo se aplica cuando la sustitución tenga prueba y aceptación, per §«Dependencias y orden de especificación» punto 5 del [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [capabilities-index](index.md) | Next: [CAP-058-retirada-procesos-win32](CAP-058-retirada-procesos-win32.md)
