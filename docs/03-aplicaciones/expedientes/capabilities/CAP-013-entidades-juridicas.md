[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-013 — Entidades y jurídicas

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-013--entidades-y-jurídicas). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Mantenimiento de las relaciones entre expediente, entidades y personas jurídicas, incluida la cadena de entidades como estructura acíclica; validación de datos inválidos, permiso ausente o dependencia no resoluble; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-013--entidades-y-jurídicas` (anchor verificado vía `git grep "### Requirement: EXP-CAP-013"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-013 / Relaciones con entidades, jurídicas y cadena de entidades.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje `CAP-045 (jurídicas), CAP-052`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/modules/modExpedienteEntidadesHelper.bas` — helper por dominio mapeado a adaptadores FastAPI por agregado en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)».
  - `TbExpedientesConEntidades` y `TbExpedientesJuridicas` (417 filas) — sub-tablas que comparten `IDExpediente` con la cabecera, inventariadas en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Inventario real Dysflow».
  - Hipótesis — no verificada desde este entorno: el entry específico que persiste la cadena de entidades no aparece en los walkthroughs de `data/staging/expedientes/docs/capabilities/`; se espera que aterrice en `src/classes/ExpedienteOperaciones.cls` siguiendo el patrón de los demás `Registrar<Vertical>`.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [capabilities-index](index.md) | Next: [CAP-014-suministradores-ute](CAP-014-suministradores-ute.md)
