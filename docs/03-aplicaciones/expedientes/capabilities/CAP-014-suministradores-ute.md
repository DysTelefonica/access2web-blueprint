[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-014 — Suministradores y UTE

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-014--suministradores-y-ute). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Asignación de contratista, subcontratista y UTE sobre un árbol acíclico con la regla «el árbol manda»; denegación sin efecto parcial ante datos inválidos, permiso ausente o dependencia no resoluble; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-014--suministradores-y-ute` (anchor verificado vía `git grep "### Requirement: EXP-CAP-014"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-014 / Suministradores, contratistas, subcontratistas, UTE y jerarquía «el árbol manda».` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje `CAP-019`, evidencia `VS`, disposición `MODERNIZAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/modules/modExpedienteHelper.bas` — helper por dominio de suministradores, mapeado a adaptadores FastAPI por agregado en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)».
  - `TbExpedientesSuministradores` — sub-tabla del agregado, inventariada en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Inventario real Dysflow».
  - Hipótesis — no verificada desde este entorno: el entry específico que aplica la regla «el árbol manda» no aparece en los walkthroughs de `data/staging/expedientes/docs/capabilities/`; se espera que aterrice en `src/classes/ExpedienteOperaciones.cls`.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-013-entidades-juridicas](CAP-013-entidades-juridicas.md) | Next: [CAP-015-comerciales](CAP-015-comerciales.md)
