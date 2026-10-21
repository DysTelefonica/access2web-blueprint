[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../../DOCS.md)

# CAP-001 — Alta transaccional

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-001--alta-transaccional). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Creación de la cabecera del expediente (`TbExpedientes`) con sus datos derivados (`TbExpedientesConEntidades`), vinculados y `TbUltimoCambio` en una sola transacción DAO.Workspace; origen manual o autorizado por HPS; preservar invariantes, auditar actor/fecha y aplicar deny-by-default.

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

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-001--alta-transaccional` (anchor verificado vía `git grep "### Requirement: EXP-CAP-001"` en `openspec/changes/expedientes-web-migration/specs/lifecycle.md`).
- **Catálogo PRD**: fila `EXP-CAP-001 / Alta transaccional del expediente, incluida el alta iniciada por sistemas autorizados.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «A. Ciclo de vida», línea 53); linaje legacy `CAP-001`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy, mismo que ancla el PRD en [docs/03-aplicaciones/expedientes/README.md](../README.md)):
  - `src/classes/ExpedienteOperaciones.cls:19` — entry `Public Function Registrar(p_DTO, p_ExpedienteAlInicio, …)` (rama transaccional con `p_ExpedienteAlInicio Is Nothing` → alta).
  - `src/classes/ExpedienteOperaciones.cls:566` — wrapper `Public Function RegistrarAlta(p_DTO, …)` que invoca `Registrar` con `p_ExpedienteAlInicio = Nothing`.
  - `src/classes/ExpedienteOperaciones.cls:1123` — `Private Function MotivoNoOKAlta(p_DTO, p_Error)` con las 16 validaciones (Ámbito, APLICAESTADO, AGEDYS, Titulo, CodExp, CodExpLargo, Nemotecnico, IDOrganoContratacion, NPedido, CodProyecto, TIpo, …).
  - `src/forms/Form_FormExpedienteAlta.cls:52` — handler `ComandoRegistrar_Click` que invoca `m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error`.
  - Evidencia completa en `data/staging/expedientes/docs/capabilities/CAP-001-alta.md` (§2 «Contrato de comportamiento», §3 «Mapa de implementación», §7 «Registro de confianza»).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [capabilities-index](index.md) | Next: [CAP-002-edicion-concurrente](CAP-002-edicion-concurrente.md)
