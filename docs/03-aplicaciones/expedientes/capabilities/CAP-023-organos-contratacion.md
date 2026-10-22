[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-023 — Órganos de contratación

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-023--órganos-de-contratación). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Gestión del catálogo de órganos de contratación con relaciones master-detail y claves estables; alta y actualización de un órgano con su descripción; denegación sin efecto parcial ante datos inválidos, permiso ausente o dependencia no resoluble; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-023--órganos-de-contratación` (anchor verificado vía `git grep "### Requirement: EXP-CAP-023"` en `openspec/changes/expedientes-web-migration/specs/catalogs.md`).
- **Catálogo PRD**: fila `EXP-CAP-023 / Órganos de contratación.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «C. Diez catálogos»); linaje `CAP-014`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `data/staging/expedientes/docs/capabilities/CAP-014-organo.md` — walkthrough legacy del catálogo de órganos (mismo linaje `CAP-014` que declara el catálogo PRD); el contenido de sus secciones §0..§7 se resume en la ficha enlazada en el índice.
  - `IDOrganoContratacion` — columna de `TbExpedientes` sin FK numérica, marcada ⚠️ en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Hallazgos críticos del esquema» (D94).
  - Hipótesis — no verificada desde este entorno: las anclas `ruta:línea` del walkthrough de órganos no se citan porque el fichero no está disponible para lectura línea a línea desde este entorno; se espera que aterricen en `src/classes/ExpedienteOperaciones.cls` como validaciones de la cabecera.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-015-comerciales](CAP-015-comerciales.md) | Next: [CAP-025-bandeja-paginada](CAP-025-bandeja-paginada.md)
