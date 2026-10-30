[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-044 — Principal Lanzadera

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-044--principal-lanzadera). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Consumo del usuario, la aplicación 19 y los permisos desde Lanzadera para construir el principal actual; denegación sin efecto parcial si Lanzadera no está disponible, el permiso falta o la dependencia no resuelve; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- D88-D91 + DA-2..DA-5: ver §Auth en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/access-control.md#requirement-exp-cap-044--principal-lanzadera` (anchor verificado vía `git grep "### Requirement: EXP-CAP-044"` en `openspec/changes/expedientes-web-migration/specs/access-control.md`).
- **Catálogo PRD**: fila `EXP-CAP-044 / Login y principal actual consumidos desde Lanzadera.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «G. Identidad, autorización y auditoría»); linaje `CAP-042`, evidencia `VS`, disposición `MODERNIZAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/classes/USUARIO.cls` — clase de dominio responsable de la sesión, citada en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)».
  - `src/modules/constructor.bas` — factory `getUsuario` (135 callers referenciados), punto único de composición del principal.
  - Hipótesis — no verificada desde este entorno: el linaje `CAP-042` declarado por el catálogo PRD no tiene walkthrough propio en `data/staging/expedientes/docs/capabilities/`. La autenticación Argon2id (D88-D91 + DA-2..DA-5) vive en Lanzadera y se describe aquí sólo como contrato consumido: Expedientes no la reimplementa.

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-043-deny-by-default](CAP-043-deny-by-default.md) | Next: [CAP-045-estado-sesion](CAP-045-estado-sesion.md)
