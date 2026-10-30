[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-010 — Anexos y referencias

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-010--anexos-y-referencias). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Gestión de anexos y referencias documentales con eliminación y retención controladas; referencia opaca al documento (sin rutas locales como contrato); denegación si la autorización o la retención lo impiden; preservar invariantes, auditar actor/fecha y aplicar deny-by-default; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- D55: ver §Decisiones arquitectónicas D-<n> cross-cutting en [docs/architecture.md](../../../architecture.md) — sin telemetría en anexos (SSID/ubicación/coordenadas).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/related-data.md#requirement-exp-cap-010--anexos-y-referencias` (anchor verificado vía `git grep "### Requirement: EXP-CAP-010"` en `openspec/changes/expedientes-web-migration/specs/related-data.md`).
- **Catálogo PRD**: fila `EXP-CAP-010 / Anexos y referencias documentales con eliminación y retención controladas.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «B. Datos relacionados»); linaje legacy `CAP-018`, evidencia `VS`, disposición `MODERNIZAR` (el mecanismo legacy `fso.FileExists` no se conserva; el storage abstracto se delega al adapter H06/CAP-056).
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - `src/classes/ExpedienteOperaciones.cls:1785` — entry `Public Function RegistrarAnexo(...)` (registro del anexo), referenciado en la lista de entries de §3 «Mapa de implementación» de `data/staging/expedientes/docs/capabilities/CAP-001-alta.md`.
  - `src/classes/ExpedienteOperaciones.cls:1102-1113` — validación 11 dentro de `MotivoNoOK`: «cada ruta (no ID numérico) debe existir como archivo (`fso.FileExists`)» (regla BR-02-11, evidenciado en `data/staging/.../CAP-002-edicion.md` §2).
  - Tabla afectada en escritura: `TbExpedientesAnexos` (FK = `IDExpediente`); referencia documental externalizada al adapter de storage de `CAP-056` (PR-6).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-009-modificados-historial](CAP-009-modificados-historial.md) | Next: [CAP-011-anualidades](CAP-011-anualidades.md)
