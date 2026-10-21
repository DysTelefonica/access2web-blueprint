[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)

# CAP-002 — Edición concurrente

Capacidad de Expedientes. Resumen a nivel repo derivado del [spec fuente](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-002--edición-concurrente). El spec contiene los escenarios G/W/T completos; este doc los resume y los vincula a las decisiones operativas vigentes.

## Escenario clave

Actualización de los datos derivados (`TbExpedientesConEntidades`) preservando la cabecera (`TbExpedientes`); `MotivoNoOK` valida 11 grupos (titular, tipología, fechas, AGEDYS, codExp, anexos) con dirty-detection implícito (consulta de unicidad solo si el campo cambió); transacción DAO + INSERT en `TbUltimoCambio` con `FechaCambio = Now()` y `IDUsuarioCambio = g_UsuarioConectado`; revertir archivo.

## Cómo se aplica a access2web-blueprint

Esta capacidad vive en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Su contrato atraviesa los slices de discovery (D48) y los hubs D51..D54 (access-control, runtime, integrations, uat-cutover-legacy-retirement) según aplique.

- DA-1: ver §Patrón hexagonal en [docs/architecture.md](../../../architecture.md).
- DA-11: ver §Decisiones de diseño del MVP en [docs/architecture.md](../../../architecture.md).
- D70: ver §Caché en [docs/architecture.md](../../../architecture.md) — no cachear contadores ni métricas de edición concurrente.

## Core invariants

- El comportamiento de esta capacidad debe preservar los invariantes definidos en el spec fuente.
- Cualquier cambio que afecte a esta capacidad requiere actualizar el spec antes que el código.
- Las decisiones D-/DA- vigentes en [docs/architecture.md](../../../architecture.md) son vinculantes.

## Contributor checklist

- [ ] Si el PR añade lógica de esta capacidad, el spec fuente sigue siendo la fuente de verdad.
- [ ] Si el PR toca la decisión referenciada arriba, actualiza el spec vía PR al `openspec/changes/expedientes-web-migration/`.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Trazabilidad al código fuente

- **Spec fuente**: `openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-002--edición-concurrente` (anchor verificado vía `git grep "### Requirement: EXP-CAP-002"` en `openspec/changes/expedientes-web-migration/specs/lifecycle.md`).
- **Catálogo PRD**: fila `EXP-CAP-002 / Edición del agregado con invariantes y concurrencia explícitas.` en [docs/03-aplicaciones/expedientes/capability-ledger.md](../capability-ledger.md) (sección «A. Ciclo de vida»); linaje legacy `CAP-002`, evidencia `VS`, disposición `PRESERVAR`.
- **Repo Expedientes** (commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy):
  - `src/classes/ExpedienteOperaciones.cls:811` — entry `Public Function RegistrarExpEntidades(Optional ByRef p_Error, Optional p_db, Optional p_Ambito, Optional p_ForceCacheFailureForTest)` que delega en `ExpedienteEntidadOperaciones.Registrar`.
  - `src/classes/ExpedienteOperaciones.cls:855-1122` — `Private Function MotivoNoOK(p_DTO, p_ObjExpedienteAlInicio, p_Error)` con los 11 grupos de validación y la consulta condicional de unicidad.
  - `src/classes/ExpedienteOperaciones.cls:19` — entry `Public Function Registrar(...)` (rama edición: `p_ExpedienteAlInicio Is Not Nothing`).
  - `src/forms/Form_FormExpediente.cls:133` — handler `ComandoActualizarCompleto_Click` que invoca `m_ExpOp.RegistrarExpEntidades p_Error:=m_Error` (línea 148).
  - Evidencia completa en `data/staging/expedientes/docs/capabilities/CAP-002-edicion.md` (§2 «Contrato de comportamento», §3 «Mapa de implementación»).

## Lista de comprobación final

- [ ] El anchor al spec fuente resuelve (verificable con `git grep`).
- [ ] Las cross-refs a [docs/architecture.md](../../../architecture.md) siguen vigentes.
- [ ] Tono castellano peninsular formal con usted per [skills/documentation-alan-style/SKILL.md](../../../../skills/documentation-alan-style/SKILL.md).

## Navigation

Previous: [CAP-001-alta-transaccional](CAP-001-alta-transaccional.md) | Next: [CAP-003-baja-condicionada](CAP-003-baja-condicionada.md)
