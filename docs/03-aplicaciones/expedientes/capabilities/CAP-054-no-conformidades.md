[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-054 — No Conformidades

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-054--no-conformidades). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Consulta del estado de No Conformidades por código de expediente o por código S4H; la respuesta devuelve el estado NC asociado al expediente; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/integrations.md#requirement-exp-cap-054--no-conformidades` (anchor verificado vía `git grep "### Requirement: EXP-CAP-054"` en `openspec/changes/expedientes-web-migration/specs/integrations.md`).
- **Catálogo PRD**: fila `EXP-CAP-054 / Enlace con No Conformidades y su estado asociado.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «I. Integraciones»); linaje `CAP-034`, evidencia `VR-hist/INT`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `CodS4H` — columna de `TbExpedientes` que sirve de clave de enlace con el estado NC, marcada en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Hallazgos críticos del esquema».
  - Hipótesis — no verificada desde este entorno: el linaje `CAP-034` declarado por el catálogo PRD no tiene walkthrough propio en `data/staging/expedientes/docs/capabilities/`; la consulta NC se espera como adapter driven de integraciones, con el estado vacío definido por contrato.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-053-gestion-riesgos](CAP-053-gestion-riesgos.md) | Next: [CAP-055-correo](CAP-055-correo.md)
