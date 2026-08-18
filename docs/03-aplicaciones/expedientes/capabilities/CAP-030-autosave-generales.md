[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)

# CAP-030 — Autosave generales y fechas

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/write-resilience.md#requirement-exp-cap-030-autosave-generales-y-fechas). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Autosave en cambios de cabecera y fechas; idempotente; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-11: ver §Decisiones D-/DA- en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep${cap}`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [capabilities-index](index.md) | Next: [DOCS](../../../DOCS.md)
