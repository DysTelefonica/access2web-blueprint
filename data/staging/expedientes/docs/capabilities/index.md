<!--
Índice de capacidades — registro maestro para navegación.
Una fila por capacidad. Una IA lee ESTO primero para encontrar una capacidad y luego abre su documento.
Idioma: castellano de España. Los enums (tier/status/source/confianza) se mantienen tal cual.
-->

# Índice de capacidades — EXPEDIENTES

> Recordatorio de fuente de verdad: código + tests de Dysflow. Este índice es un mapa, no evidencia.
> Linaje: PRUEBA-002 (proposal + 56 specs) + PRUEBA-003 (refactor a 17 helpers) + cobertura acumulada de `tests/tests.vba.json`.

## Resumen ejecutivo

| Métrica | Valor |
|---|---|
| Capabilities (CAP-NNN) | 55 (numeración 001–055, con CAP-038 a CAP-055 no contiguos) |
| Dominios | 11 |
| Business rules (BR) | ~60 |
| Reglas `Verified-runtime` pre-PRUEBA-002 | 22 |
| Reglas `Verified-static` (deuda) | ~38 |
| Tests atómicos pre-cambio | 78 |
| **Tests target post-PRUEBA-002** | **~128** (78 + 38 nuevos + ~12 soporte) |
| Chained PRs PRUEBA-003 | 14 (REFAC-1a/1b/1c, 2a/2b, 3a/3b, 4a/4b1/4b2/4c1/4c2, 5) |
| Helpers introducidos en PRUEBA-003 | 17 (14 `.bas` + 3 `.cls`) |
| LOC estimados refactor + tests | ~3,750 |
| Budget por PR (encadenado) | 400 líneas |

## §1 Capabilities por dominio

### Dominio 1 — Ciclo de vida del expediente (PR-C)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-001 | Alta de expediente | critical | active | hybrid | mixta | 0/5 | [CAP-001-alta.md](./CAP-001-alta.md) |
| CAP-002 | Edición de expediente | critical | active | hybrid | mixta | 0/2 | [CAP-002-edicion.md](./CAP-002-edicion.md) |
| CAP-003 | Eliminación de expediente | standard | active | hybrid | mixta | 0/2 | [CAP-003-eliminacion.md](./CAP-003-eliminacion.md) |
| CAP-004 | Cambio de tipo | standard | active | hybrid | mixta | 0/1 | [CAP-004-cambio-tipo.md](./CAP-004-cambio-tipo.md) |
| CAP-005 | Estado calculado | standard | active | hybrid | mixta | 0/1 | [CAP-005-estado-calculado.md](./CAP-005-estado-calculado.md) |
| CAP-045 | Basados / lotes | standard | active | hybrid | mixta | 0/1 | [CAP-045-basados-lotes.md](./CAP-045-basados-lotes.md) |
| CAP-046 | Derivación ordinal (E2E) | standard | active | hybrid | Verified-runtime | 1/0 | [CAP-046-derivacion-ordinal.md](./CAP-046-derivacion-ordinal.md) |

### Dominio 2 — Gestión de entidades (PR-D1 / PR-D2)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-006 | Comercial | standard | active | hybrid | mixta | 0/2 | [CAP-006-comercial.md](./CAP-006-comercial.md) |
| CAP-007 | CPV | standard | active | hybrid | mixta | 0/1 | [CAP-007-cpv.md](./CAP-007-cpv.md) |
| CAP-008 | Ejército | standard | active | hybrid | mixta | 0/1 | [CAP-008-ejercito.md](./CAP-008-ejercito.md) |
| CAP-009 | Suministrador | standard | active | hybrid | mixta | 0/1 | [CAP-009-suministrador.md](./CAP-009-suministrador.md) |
| CAP-010 | Lugar de ejecución | standard | active | hybrid | mixta | 1/0 | [CAP-010-lugar.md](./CAP-010-lugar.md) |
| CAP-011 | PECAL | standard | active | hybrid | mixta | 0/1 | [CAP-011-pecal.md](./CAP-011-pecal.md) |
| CAP-012 | RAC | standard | active | hybrid | mixta | 1/1 | [CAP-012-rac.md](./CAP-012-rac.md) |
| CAP-013 | Grado de clasificación | standard | active | hybrid | mixta | 0/1 | [CAP-013-grado.md](./CAP-013-grado.md) |
| CAP-014 | Órgano de contratación | standard | active | hybrid | mixta | 0/1 | [CAP-014-organo.md](./CAP-014-organo.md) |
| CAP-015 | Oficina programa | standard | active | hybrid | mixta | 0/1 | [CAP-015-oficina.md](./CAP-015-oficina.md) |

### Dominio 3 — E2E batch platform (PR-B)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-026 | E2E management service | critical | active | sdd | Verified-runtime | 5/0 | [CAP-026-management.md](./CAP-026-management.md) |
| CAP-027 | E2E hash service | critical | active | sdd | Verified-runtime | 4/0 | [CAP-027-hash.md](./CAP-027-hash.md) |
| CAP-028 | E2E export service | critical | active | sdd | Verified-runtime | 5/0 | [CAP-028-export.md](./CAP-028-export.md) |
| CAP-029 | JSON exporter | standard | active | sdd | Verified-runtime | 10/0 | [CAP-029-json.md](./CAP-029-json.md) |
| CAP-030 | E2E traceability schema | standard | active | sdd | Verified-runtime | 3/0 | [CAP-030-traceability-schema.md](./CAP-030-traceability-schema.md) |

### Dominio 4 — Autosave (PR-E)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-020 | Autosave General / Fechas | critical | active | sdd | Verified-runtime | 2/0 | [CAP-020-general-fechas.md](./CAP-020-general-fechas.md) |
| CAP-021 | Autosave otros tabs | standard | active | sdd | mixta | 0/2 | [CAP-021-otros-tabs.md](./CAP-021-otros-tabs.md) |

### Dominio 5 — Documental (PR-C)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-018 | Anexos | standard | active | hybrid | mixta | 0/2 | [CAP-018-anexos.md](./CAP-018-anexos.md) |
| CAP-019 | Suministradores árbol | standard | active | hybrid | mixta | 0/1 (deuda OLE drag&drop diferida) | [CAP-019-suministradores-arbol.md](./CAP-019-suministradores-arbol.md) |

### Dominio 6 — Consultas e informes (PR-E)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-022 | Bandeja (consultas) | standard | active | hybrid | mixta | 0/1 | [CAP-022-bandeja.md](./CAP-022-bandeja.md) |
| CAP-023 | Búsqueda avanzada | standard | active | hybrid | mixta (slice actual) | 6/2 | [CAP-023-busqueda-avanzada.md](./CAP-023-busqueda-avanzada.md) |
| CAP-024 | Búsqueda técnica | standard | active | hybrid | mixta | 0/1 | [CAP-024-busqueda-tecnica.md](./CAP-024-busqueda-tecnica.md) |
| CAP-025 | Exportación Excel | standard | active | hybrid | mixta | 0/1 | [CAP-025-excel.md](./CAP-025-excel.md) |

### Dominio 7 — Tareas (PR-F)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-040 | Bandeja de tareas pendientes | minimal | **deferred** | aspirational | **Intended / Divergent** (no backend `TbTareas`) | 0/0 | [CAP-040-tareas-bandeja.md](./CAP-040-tareas-bandeja.md) |

### Dominio 8 — Configuración (PR-B / PR-F)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-031 | Backend switching | critical | active | sdd | Verified-runtime | 7/0 | [CAP-031-backend-switching.md](./CAP-031-backend-switching.md) |
| CAP-032 | Startup infra | standard | active | sdd | Verified-runtime | 2/0 | [CAP-032-startup-infra.md](./CAP-032-startup-infra.md) |
| CAP-033 | Cache hardening | critical | active | sdd | Verified-runtime | 5/0 | [CAP-033-cache.md](./CAP-033-cache.md) |
| CAP-037 | Entorno singleton | standard | active | hybrid | mixta | 0/1 | [CAP-037-entorno.md](./CAP-037-entorno.md) |
| CAP-043 | Anualidades | standard | active | hybrid | mixta | 0/1 | [CAP-043-anualidades.md](./CAP-043-anualidades.md) |

### Dominio 9 — Seguridad (PR-F)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-041 | Autorización por roles | standard | active | hybrid | mixta | 0/3 | [CAP-041-autorizacion.md](./CAP-041-autorizacion.md) |
| CAP-042 | Login / usuario conectado | standard | active | hybrid | mixta | 0/1 | [CAP-042-login.md](./CAP-042-login.md) |

### Dominio 10 — Anti-spam / popup de progreso (PR-A)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-035 | Anti-spam contract | critical | active | sdd | mixta (PR-A agrega tests) | 0/1 | [CAP-035-anti-spam.md](./CAP-035-anti-spam.md) |
| CAP-036 | Popup `frmBusy` | standard | active | sdd | mixta (PR-A agrega tests) | 0/1 | [CAP-036-frmBusy.md](./CAP-036-frmBusy.md) |

### Dominio 11 — E2E traceability (PR-B / PR-E)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-047 | OrdinalE2E domain | standard | active | sdd | Verified-runtime | 2/0 | [CAP-047-ordinal-e2e.md](./CAP-047-ordinal-e2e.md) |
| CAP-048 | Selección manual | standard | active | sdd | Verified-runtime | 2/0 | [CAP-048-seleccion-manual.md](./CAP-048-seleccion-manual.md) |
| CAP-049 | Destino / per-user config | standard | active | sdd | Verified-runtime | 3/0 | [CAP-049-destino.md](./CAP-049-destino.md) |
| CAP-050 | Sesión de exportación | standard | active | sdd | Verified-runtime | 1/0 | [CAP-050-sesion.md](./CAP-050-sesion.md) |

### Misceláneos (PR-C / PR-E / PR-F)

| ID | Nombre | Tier | Estado | Source | Confianza | Reglas (VR/VS) | Doc |
|---|---|---|---|---|---|---|---|
| CAP-016 | Hitos | standard | active | hybrid | mixta | 0/1 | [CAP-016-hitos.md](./CAP-016-hitos.md) |
| CAP-017 | Modificados | standard | active | hybrid | mixta | 0/1 | [CAP-017-modificados.md](./CAP-017-modificados.md) |
| CAP-034 | No conformidades (NCs) | critical | active | sdd | Verified-runtime | 1/0 | [CAP-034-ncs.md](./CAP-034-ncs.md) |
| CAP-038 | DTO contract | standard | active | hybrid | mixta (cubr. por helper) | 0/1 | [CAP-038-dto.md](./CAP-038-dto.md) |
| CAP-039 | Último cambio (audit-trail) | standard | active | hybrid | mixta | 0/1 | [CAP-039-ultimo-cambio.md](./CAP-039-ultimo-cambio.md) |
| CAP-044 | Jefaturas | standard | active | hybrid | mixta | 0/1 | [CAP-044-jefaturas.md](./CAP-044-jefaturas.md) |
| CAP-051 | AGEDYS (integración externa) | minimal | deferred | sdd | Intended | 0/1 | [CAP-051-agedys.md](./CAP-051-agedys.md) |
| CAP-052 | Cadena de entidades | standard | active | hybrid | mixta | 0/1 | [CAP-052-cadena-entidades.md](./CAP-052-cadena-entidades.md) |
| CAP-053 | Estado del usuario | standard | active | hybrid | mixta | 0/1 | [CAP-053-estado-usuario.md](./CAP-053-estado-usuario.md) |
| CAP-054 | Win32 net | minimal | active | hybrid | mixta | 0/1 | [CAP-054-win32-net.md](./CAP-054-win32-net.md) |
| CAP-055 | Win32 process | minimal | active | hybrid | mixta | 0/1 | [CAP-055-win32-process.md](./CAP-055-win32-process.md) |

## §2 Lagunas de cobertura (obligaciones abiertas — parte de la entrega de este change)

> Cada fila es un test que crear con `access-vba-tdd` siguiendo el patrón de los 6 ya verdes de `slice-refac-1a`.
> Total: **38 reglas `Verified-static` que migran a `Verified-runtime` tras PRUEBA-002 + PRUEBA-003**.

| Capacidad | Regla VS | PR destino | Acción |
|---|---|---|---|
| CAP-001 | BR-01-01..05 (5) | PR-C | `Helper_ExpedienteAlta` + 5 tests atómicos |
| CAP-002 | BR-02-01..02 (2) | PR-C | `Helper_ExpedienteEdicion` + 2 tests |
| CAP-003 | BR-03-01..02 (2) | PR-C | `Helper_ExpedienteEliminacion` + 2 tests |
| CAP-004 | BR-04-01 (1) | PR-C | `Helper_ExpedienteEdicion.CambiarTipo` + 1 test |
| CAP-005 | BR-05-01 (1) | PR-C | `Helper_ExpedienteEdicion.EstadoCalculado` + 1 test |
| CAP-016 | BR-16-01 (1) | PR-E | `Helper_ExpedienteHitos` + 1 test |
| CAP-017 | BR-17-01 (1) | PR-E | `Helper_ExpedienteModificados` + 1 test |
| CAP-018 | BR-18-01..02 (2) | PR-C | `Helper_ExpedienteAnexos` + 2 tests |
| CAP-019 | BR-19-01 (1, OLE drag&drop deuda) | PR-C | `Helper_ExpedienteSuministradoresArbol` (sin OLE) + 1 test |
| CAP-006..CAP-015 | 1–2 reglas c/u (10 entidades) | PR-D1/D2 | `Helper_EntidadCRUD` + tests por entidad |
| CAP-020 | ya `Verified-runtime` (2) | — | sin acción |
| CAP-021 | BR-21-01..02 (2) | PR-E | `Helper_ExpedienteHitos`/`Helper_ExpedienteModificados` autosave + 2 tests |
| CAP-022 | BR-22-01 (1) | PR-E | `Helper_BandejaTareas.CargarTareas` + 1 test |
| **CAP-023** | **BR-23-01..02 (2)** | **PR-E** | **`Helper_ExpedienteConsultas.ConstruirWhereBusqueda` — los 6 tests `slice-refac-1a` ya cubren BR-23-01 en runtime** |
| CAP-024 | BR-24-01 (1) | PR-E | `Helper_ExpedienteConsultas.ConstruirColBusquedaTecnica` + 1 test |
| CAP-025 | BR-25-01 (1) | PR-E | `Helper_ExpedientesExportExcel` + 1 test |
| CAP-035 | BR-35-01 (1) | PR-A | `Test_AntiSpam_AntiSpamEnterExit_SetClearFlag` + 1 test |
| CAP-036 | BR-36-01 (1) | PR-A | `Test_FrmBusy_ShowHide_HonorsAntiSpamFlag` + 1 test |
| CAP-040 | BR-40-01..02 (2) | PR-F | **`Helper_BandejaTareas` + 2 tests** + **crear `TbTareas` en backend** (deuda de producto, target post-PR-F) |
| CAP-041 | BR-41-01..03 (3) | PR-F | `Helper_Autorizacion` + 3 tests |
| CAP-042 | BR-42-01 (1) | PR-F | `Helper_Login` + 1 test |
| CAP-043 | BR-43-01 (1) | PR-E | `Helper_ExpedienteAnualidades` + 1 test |
| CAP-053 | BR-53-01 (1) | PR-E | `Helper_Login` + 1 test (re-tag) |
| CAP-054..CAP-055 | BR-54-01, BR-55-01 (2) | PR-E | VBA Declare + 2 tests |
| CAP-038 | BR-38-01 (1) | PR-C | cubierto por `Test_Helper_ExpedienteAlta_*` |
| CAP-039 | BR-39-01 (1) | PR-F | `Helper_UltimoCambio` + 1 test |
| CAP-044 | BR-44-01 (1) | PR-F | `Helper_EntidadCRUD.BuscarJefaturas` + 1 test |
| CAP-045 (jurídicas) | BR-45-02 (1) | PR-F | `Helper_EntidadCRUD.BuscarJuridicas` + 1 test |
| CAP-052 | BR-52-01 (1) | PR-F | `Helper_ExpedienteEntidadesTabla` + 1 test |
| CAP-037 | BR-37-01 (1) | PR-F | `Helper_EntornoCarga` + 1 test |
| CAP-051 | BR-51-01 (1) | deferred (fuera de scope) | deuda AGEDYS — fuera de PRUEBA-002/003 |

## §3 Divergencias pendientes de revisión humana

| Capacidad | Hallazgo | Detectada | Estado |
|---|---|---|---|
| (sin hallazgos activos a 2026-06-15) | — | — | — |

> Si durante la ejecución de un test aparece una divergencia entre intención SDD y realidad del código, se documenta aquí y en §7 del doc de capacidad correspondiente.

## §4 Próximos pasos

1. **Cerrar PR-REFAC-1a** (PRUEBA-003 REFAC-1a) con la documentación CAP-023 ya en este índice + los 6 tests de `slice-refac-1a` como puente `Verified-runtime`.
2. **Iterar dominio por dominio** abriendo el doc de cada capacidad según se aterricen los PRs: cada `commit: refactor(...)` debe ir acompañado de su `commit: docs(capabilities/CAP-NNN)` con §0-§7 completo.
3. **Cuando PRUEBA-002 PR-A (docs-only anti-spam) cierre**, los docs CAP-035 y CAP-036 quedan `Verified-runtime`; idem para los otros PRs.
4. **Verificación final**: `dysflow.test_vba` full run verde con 128 tests; cada doc de capacidad muestra todos los BR en `Verified-runtime`.

## §5 Convenciones

- **Formato del doc**: `docs/capabilities/CAP-NNN-<slug>.md` siguiendo `access-vba-capability-docs/assets/capability-doc-template.md`.
- **Tiers**: `critical` (sin este doc la app no sirve) / `standard` (reglas de negocio cotidianas) / `minimal` (helpers de plataforma).
- **Estados**: `active` / `deprecated` / `broken` / `deferred`.
- **Source**: `sdd` (nacido de una spec) / `reverse-engineered` (solo de código) / `hybrid` (ambos).
- **Confianza**: `Verified-runtime` (test verde) / `Verified-static` (leído, sin test — deuda) / `Intended` / `Likely` / `Divergent` / `mixta` (combinación).
- **Reglas VR/VS**: "reglas en `Verified-runtime`" / "reglas en `Verified-static`".
