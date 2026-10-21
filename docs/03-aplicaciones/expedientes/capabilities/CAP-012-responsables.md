[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)

# CAP-012 — Responsables

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-012--responsables). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Definición y mantenimiento de los responsables del expediente por rol (gestor, jefe de proyecto, responsable de calidad, etc.) y de la cadena de jefaturas; las relaciones PADRE/HIJO según el organigrama vigente; denegación si el rol o la cadena no son resolubles; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- DA-12: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md) — mapping legacy → profiles en tabla inmutable.

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-012--responsables` (anchor verificado vía `git grep "### Requirement: EXP-CAP-012"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-012 / Responsables por rol y jefaturas.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje legacy `CAP-044 + inventario`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - `src/classes/USUARIO.cls` — clase de dominio responsable de la sesión y rol (citada como capa de dominio en la tabla de mapeo hexagonal de [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)»).
  - `src/modules/constructor.bas` — factory `getUsuario` (135 callers referenciados en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)»).
  - Hipótesis — no verificada desde este entorno: el entry específico `RegistrarResponsable(...)` no aparece en la lista de §3 de `data/staging/.../CAP-001-alta.md`; forma parte de los verticales R05 (responsabilidad) del plan `tasks.md` y se espera que aterrice en `src/classes/ExpedienteOperaciones.cls` siguiendo el patrón de los demás `Registrar<Vertical>` (R01..R07).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-011-anualidades](CAP-011-anualidades.md) | Next: [capabilities-index](index.md)
