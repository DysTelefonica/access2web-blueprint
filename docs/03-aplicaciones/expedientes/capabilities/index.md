[← Back to Expedientes README](../README.md)

# Expedientes — capacidades documentadas (índice)

Este índice enlaza las capacidades del Expedientes que ya están documentadas en `data/staging/expedientes/docs/capabilities/` (single source of truth del walkthrough de discovery). Cada CAP tiene un fichero estructurado con secciones §0 Identidad, §1 Intención, §2 Contrato de comportamiento, §3 Modelo, §4 Walkthrough, §5 Hallazgos, §6 Decisiones, §7 Anexos.

`data/staging/` es material de sólo lectura per `AGENTS.md` §«Alcance». No se modifica aquí; se referencia desde este índice.

## Capacidades ya documentadas

17 CAPs en `data/staging/expedientes/docs/capabilities/` a 2026-08-18:

| ID | Título corto | Ruta en staging | Dominio |
|---|---|---|---|
| CAP-001 | Alta de expediente | `data/staging/expedientes/docs/capabilities/CAP-001-alta.md` | Expediente (escritura) |
| CAP-002 | Edición de expediente | `data/staging/expedientes/docs/capabilities/CAP-002-edicion.md` | Expediente (escritura) |
| CAP-003 | Eliminación condicionada | `data/staging/expedientes/docs/capabilities/CAP-003-eliminacion.md` | Expediente (escritura) |
| CAP-004 | Cambio de tipo | [`CAP-004-cambio-tipo.md`](CAP-004-cambio-tipo.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-004--cambio-de-tipo)) | Expediente (escritura) |
| CAP-005 | Estado y garantía | [`CAP-005-estado-garantia.md`](CAP-005-estado-garantia.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/lifecycle.md#requirement-exp-cap-005--estado-y-garantía)) | Expediente (escritura) |
| CAP-006 | Comercial | `data/staging/expedientes/docs/capabilities/CAP-006-comercial.md` | Catálogos (master-detail) |
| CAP-007 | CPV | `data/staging/expedientes/docs/capabilities/CAP-007-cpv.md` | Catálogos (master-detail) |
| CAP-008 | Ejército | `data/staging/expedientes/docs/capabilities/CAP-008-ejercito.md` | Catálogos (master-detail) |
| CAP-009 | Suministrador | `data/staging/expedientes/docs/capabilities/CAP-009-suministrador.md` | Suministradores |
| CAP-010 | Lugar | `data/staging/expedientes/docs/capabilities/CAP-010-lugar.md` | Catálogos (master-detail) |
| CAP-011 | PECAL | `data/staging/expedientes/docs/capabilities/CAP-011-pecal.md` | Catálogos (master-detail) |
| CAP-012 | RAC | `data/staging/expedientes/docs/capabilities/CAP-012-rac.md` | Catálogos (master-detail) |
| CAP-013 | Grado | `data/staging/expedientes/docs/capabilities/CAP-013-grado.md` | Catálogos (master-detail) |
| CAP-014 | Órgano | `data/staging/expedientes/docs/capabilities/CAP-014-organo.md` | Catálogos (master-detail) |
| CAP-015 | Oficina | `data/staging/expedientes/docs/capabilities/CAP-015-oficina.md` | Catálogos (master-detail) |
| CAP-016 | CPV | [`CAP-016-cpv.md`](CAP-016-cpv.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-016--cpv)) | Catálogos (master-detail) |
| CAP-017 | Ejércitos | [`CAP-017-ejercitos.md`](CAP-017-ejercitos.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-017--ejércitos)) | Catálogos (master-detail) |
| CAP-018 | Suministradores | [`CAP-018-suministradores.md`](CAP-018-suministradores.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-018--suministradores)) | Suministradores |
| CAP-019 | Lugares de ejecución | [`CAP-019-lugares-ejecucion.md`](CAP-019-lugares-ejecucion.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-019--lugares-de-ejecución)) | Catálogos (master-detail) |
| CAP-020 | PECAL | [`CAP-020-pecal.md`](CAP-020-pecal.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-020--pecal)) | Catálogos (master-detail) |
| CAP-021 | RAC | [`CAP-021-rac.md`](CAP-021-rac.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-021--rac)) | Catálogos (master-detail) |
| CAP-022 | Grados de clasificación | [`CAP-022-grados-clasificacion.md`](CAP-022-grados-clasificacion.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-022--grados-de-clasificación)) | Catálogos (master-detail) |
| CAP-023 | Búsqueda avanzada | `data/staging/expedientes/docs/capabilities/CAP-023-busqueda-avanzada.md` | Consulta |
| CAP-024 | Oficinas de programa | [`CAP-024-oficinas-programa.md`](CAP-024-oficinas-programa.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/catalogs.md#requirement-exp-cap-024--oficinas-de-programa)) | Catálogos (master-detail) |
| CAP-025 | Excel | `data/staging/expedientes/docs/capabilities/CAP-025-excel.md` | Exportación |
| CAP-026 | Búsqueda avanzada (consulta) | [`CAP-026-busqueda-avanzada.md`](CAP-026-busqueda-avanzada.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/query-and-tasks.md#requirement-exp-cap-026--búsqueda-avanzada)) | Consulta |
| CAP-027 | Búsqueda técnica | [`CAP-027-busqueda-tecnica.md`](CAP-027-busqueda-tecnica.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/query-and-tasks.md#requirement-exp-cap-027--búsqueda-técnica)) | Consulta |
| CAP-028 | Exportación Excel | [`CAP-028-exportacion-excel.md`](CAP-028-exportacion-excel.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/query-and-tasks.md#requirement-exp-cap-028--exportación-excel)) | Exportación |
| CAP-029 | Tareas calculadas | [`CAP-029-tareas-calculadas.md`](CAP-029-tareas-calculadas.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/query-and-tasks.md#requirement-exp-cap-029--tareas-calculadas)) | Tareas |
| CAP-030 | Autosave generales y fechas | [`CAP-030-autosave-generales.md`](CAP-030-autosave-generales.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/write-resilience.md#requirement-exp-cap-030--autosave-generales-y-fechas)) | Escritura |
| CAP-031 | Autosave relacionados | [`CAP-031-autosave-relacionados.md`](CAP-031-autosave-relacionados.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/write-resilience.md#requirement-exp-cap-031--autosave-relacionados)) | Escritura |
| CAP-032 | Idempotencia y feedback | [`CAP-032-idempotencia-feedback.md`](CAP-032-idempotencia-feedback.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/write-resilience.md#requirement-exp-cap-032--idempotencia-y-feedback)) | Escritura |
| CAP-033 | DTO estable | [`CAP-033-dto-estable.md`](CAP-033-dto-estable.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-033--dto-estable)) | E2E |
| CAP-034 | JSON determinista | [`CAP-034-json-determinista.md`](CAP-034-json-determinista.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-034--json-determinista)) | E2E |
| CAP-035 | Batch E2E | [`CAP-035-batch-e2e.md`](CAP-035-batch-e2e.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-035--batch-e2e)) | E2E |
| CAP-036 | Hash versionado | [`CAP-036-hash-versionado.md`](CAP-036-hash-versionado.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-036--hash-versionado)) | E2E |
| CAP-037 | Paquete E2E | [`CAP-037-paquete-e2e.md`](CAP-037-paquete-e2e.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-037--paquete-e2e)) | E2E |
| CAP-038 | Trazabilidad E2E | [`CAP-038-trazabilidad-e2e.md`](CAP-038-trazabilidad-e2e.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-038--trazabilidad-e2e)) | E2E |
| CAP-039 | Selección manual | [`CAP-039-seleccion-manual.md`](CAP-039-seleccion-manual.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-039-selección-manual)) | E2E |
| CAP-040 | Destino por usuario | `data/staging/expedientes/docs/capabilities/CAP-040-tareas-bandeja.md` | Tareas (E2E) |
| CAP-041 | Sesión E2E | [`CAP-041-sesion-e2e.md`](CAP-041-sesion-e2e.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-041-sesión-e2e)) | E2E |
| CAP-042 | Ordinal E2E | [`CAP-042-ordinal-e2e.md`](CAP-042-ordinal-e2e.md) ([spec](../../../../openspec/changes/expedientes-web-migration/specs/e2e.md#requirement-exp-cap-042-ordinal-e2e)) | E2E |
| CAP-040 | Tareas (bandeja) | `data/staging/expedientes/docs/capabilities/CAP-040-tareas-bandeja.md` | Tareas |

## Capacidades pendientes de documentar

A 2026-08-18, todas las CAPs pedidas en issues abiertos están documentadas a nivel repo. Los 17 ya presentes en staging cubren la granularidad §0..§7; los 26 nuevos (CAP-004, 005, 016..022, 024, 026..039, 041, 042) son resúmenes vinculados a los Requirements del spec fuente. La lista exhaustiva está en las tablas de arriba (43 entradas totales).

## Cómo se aplica a access2web-blueprint

Este índice es el hub de las capacidades del Expedientes a nivel repo. La fuente de verdad por capacidad vive en `data/staging/expedientes/docs/capabilities/` (read-only); este índice sólo expone qué hay y dónde buscarlo. El código destino de cada capacidad vivirá en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Las decisiones operativas D86 (forma hexagonal del legacy preservada), D87 (Test_* VBA preservados), D94 (FKs conceptuales), D102 (booleans Text(2) → BOOLEAN) siguen aplicando.

## Lista de comprobación final

- [ ] Las 17 entradas referencian paths staging que existen (verificables con `ls`).
- [ ] Las 26 entradas nuevas referencian anchors a los Requirements del spec fuente (verificables con `grep`).
- [ ] El path `data/staging/` NO se ha modificado (read-only).
- [ ] El índice se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Cross-references desde DOCS.md y `docs/03-aplicaciones/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [Expedientes README](../README.md) | Next: [DOCS](../../../DOCS.md)
