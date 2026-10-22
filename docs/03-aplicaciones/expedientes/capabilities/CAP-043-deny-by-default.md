[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-043 — Deny-by-default

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-043--deny-by-default). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Autorización de cada capacidad contra `effective_permissions` del principal actual; todo permiso ausente, desconocido o no resoluble deniega la operación sin efecto parcial y deja la decisión registrada; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-043--deny-by-default` (anchor verificado vía `git grep "### Requirement: EXP-CAP-043"` en `openspec/changes/expedientes-web-migration/specs/access-control.md`).
- **Catálogo PRD**: fila `EXP-CAP-043 / Autorización deny-by-default por capacidad.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «G. Identidad, autorización y auditoría»); linaje `CAP-041`, evidencia `VS`, disposición `MODERNIZAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/classes/USUARIO.cls` — clase de dominio con el estado de sesión y rol que el gate de autorización consulta, citada en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)».
  - `src/modules/constructor.bas` — factory `getUsuario` como punto de resolución del principal.
  - Hipótesis — no verificada desde este entorno: el linaje `CAP-041` declarado por el catálogo PRD no tiene walkthrough propio en `data/staging/expedientes/docs/capabilities/`; el gate se espera como cross-cutting de `app/src/modules/expedientes/` siguiendo la tabla de purzas de DA-1.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-040-destino-usuario](CAP-040-destino-usuario.md) | Next: [CAP-044-principal-lanzadera](CAP-044-principal-lanzadera.md)
