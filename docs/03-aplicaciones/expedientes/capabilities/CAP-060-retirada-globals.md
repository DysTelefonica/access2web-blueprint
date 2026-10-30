[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-060 — Retirada globals

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-060--retirada-globals). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Retirada de los globals y singletons mutables de sesión y entorno; el estado de usuario y de entorno pasa a request y sesión con dependencias explícitas; preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-12: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md#requirement-exp-cap-060--retirada-globals` (anchor verificado vía `git grep "### Requirement: EXP-CAP-060"` en `openspec/changes/expedientes-web-migration/specs/uat-cutover-legacy-retirement.md`).
- **Catálogo PRD**: fila `EXP-CAP-060 / Globals y singletons mutables de sesión/entorno.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «J. Mecanismos legacy retirados o reemplazados»); linaje `CAP-037, CAP-053`, evidencia `—`, disposición `RETIRAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `Variables Globales.bas` (`m_TestingMode`, `m_TestOnlyBackendActivo`, `m_TestOnlyBackendURL`) — singleton mutable legacy citado en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Mecanismo de sandbox/testing en staging»; se sustituye por estado por request y dependencias explícitas.
  - `src/classes/USUARIO.cls` y `src/classes/Entorno.cls` — clases de dominio que concentran hoy ese estado, citadas en [docs/03-aplicaciones/expedientes/README.md](../README.md) §«Descubrimiento de staging (2026-08-05, codegraph-vba)»; el contrato de estado aislado es `CAP-045-estado-sesion` (ficha de la cadena de #807, PR 4 de 8; no enlazada aquí para no apuntar a un fichero ausente en esta rama).
  - Retirada condicionada: sólo se aplica cuando la sustitución tenga prueba y aceptación, per §«Dependencias y orden de especificación» punto 5 del [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-059-retirada-ole-activex](CAP-059-retirada-ole-activex.md) | Next: [CAP-061-retirada-selector-backend](CAP-061-retirada-selector-backend.md)
