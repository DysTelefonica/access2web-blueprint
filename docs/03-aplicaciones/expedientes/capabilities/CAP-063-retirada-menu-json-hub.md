[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-063 — Retirada menú JSON-hub

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-063--retirada-menú-json-hub). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Retirada del menú JSON-hub y de la navegación acoplada a formularios; el módulo se sirve por rutas SSR autorizadas por capability, de modo que la autorización gobierna la navegación y no la existencia del enlace; preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-063--retirada-menú-json-hub` (anchor verificado vía `git grep "### Requirement: EXP-CAP-063"` en `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md`).
- **Catálogo PRD**: fila `EXP-CAP-063 / Menú JSON-hub y navegación acoplada a formularios.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «J. Mecanismos legacy retirados o reemplazados»); linaje `inventario de navegación`, evidencia `—`, disposición `RETIRAR`.
- **Pendiente (P-EXP-*)**: la issue #807 deja abierta la elección de la ruta canónica entre `docs/05-capacidades/...` y `docs/03-aplicaciones/.../capabilities/`; esta ficha sigue la convención ya adoptada por las 26+13 anteriores sin resolver la decisión.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/forms/Form_FormExpediente.cls` y `src/forms/Form_FormTareas.cls` — adaptadores driving UI cuya navegación acoplada se sustituye por rutas SSR, según el mapeo de [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)».
  - Hipótesis — no verificada desde este entorno: el PRD README no nombra el JSON-hub en la tabla de mapeo hexagonal, de modo que no hay ancla `ruta:línea` verificable; el contrato de autorización que gobierna la ruta es `CAP-043-deny-by-default` (ficha de la cadena de #807, PR 4 de 8; no enlazada aquí para no apuntar a un fichero ausente en esta rama).
  - Retirada condicionada: sólo se aplica cuando la sustitución tenga prueba y aceptación, per §«Dependencias y orden de especificación» punto 5 del [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-062-retirada-popups](CAP-062-retirada-popups.md) | Next: [capabilities-index](index.md)
