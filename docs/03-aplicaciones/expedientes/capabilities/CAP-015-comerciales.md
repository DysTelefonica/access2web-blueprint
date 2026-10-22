[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-015 — Comerciales

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-015--comerciales). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Gestión del catálogo de comerciales con relaciones master-detail y claves estables; alta y actualización de una entrada maestra con su descripción; denegación sin efecto parcial ante datos inválidos, permiso ausente o dependencia no resoluble; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-015--comerciales` (anchor verificado vía `git grep "### Requirement: EXP-CAP-015"` en `openspec/changes/expedientes-web-migration/specs/catalogs.md`).
- **Catálogo PRD**: fila `EXP-CAP-015 / Comerciales.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «C. Diez catálogos»); linaje `CAP-006`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `data/staging/expedientes/docs/capabilities/CAP-006-comercial.md:9` — sección «§0 Identidad» del walkthrough legacy de la capacidad (mismo linaje `CAP-006` que declara el catálogo PRD).
  - `data/staging/expedientes/docs/capabilities/CAP-006-comercial.md:30` — sección «§2 Contrato de comportamiento», ancla de regresión del walkthrough.
  - `data/staging/expedientes/docs/capabilities/CAP-006-comercial.md:60` — sección «§3 Especificacion tecnica».
  - `TbExpedientesComerciales` — sub-tabla master-detail del agregado, inventariada en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Inventario real Dysflow».

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-014-suministradores-ute](CAP-014-suministradores-ute.md) | Next: [CAP-023-organos-contratacion](CAP-023-organos-contratacion.md)
