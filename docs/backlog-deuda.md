# Backlog de deuda técnica

> **Sentence that organizes the whole repo**: «Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras siete apps son consumidoras.»

Este documento centraliza la deuda técnica dispersa por el repositorio para que cualquier IA o humano pueda extraer **issues suficientemente pequeños** sin tener que parsear épicas, decisiones, bugs filed y changelogs por separado. Cada entrada tiene criterios de aceptación verificables, scope explícito y estimación de tamaño.

> **Convención**: el archivo vive en `docs/backlog-deuda.md` (sin numeración al inicio) para cumplir con `documentation-alan-style §9`. Cuando se renombre la estructura de `docs/`, este archivo debe moverse al lugar que decida el usuario.

---

## Resumen por prioridad

| ID | Título | Tipo | Prioridad | Tamaño | Scope | Origen |
|---|---|---|---|---|---|---|
| [BACKLOG-001](#backlog-001) | Renombrar `data/staging/lanzaderas/` → `lanzadera/` | refactor | high | M | root | Auditoría docs Bloque 2 |
| [BACKLOG-002](#backlog-002) | Eliminar numeración `00-`, `01-`, ..., `09-` en `docs/` | refactor | high | L | root | Auditoría docs Bloque 2 |
| [BACKLOG-003](#backlog-003) | Frontmatter YAML en docs raíz | chore | high | S | root | documentation-alan-style §4 |
| [BACKLOG-004](#backlog-004) | Traducir secciones en inglés técnico neutro a castellano peninsular formal | docs | medium | L | root | Auditoría docs Bloque 2 |
| [BACKLOG-005](#backlog-005) | Reemplazar referencias a `documentation-patterns` por `documentation-alan-style` | docs | high | S | root | Skill deprecada |
| [BACKLOG-006](#backlog-006) | Dysflow #1408 — `map_form_behavior --json` rejected | bug | high | S | dysflow | Issues filed |
| [BACKLOG-007](#backlog-007) | Dysflow #1412 — `verify_form_bindings` RESULT_CONTRACT_VIOLATION | bug | critical | S | dysflow | Issues filed (CRÍTICO: habría detectado BACKLOG-016) |
| [BACKLOG-008](#backlog-008) | P10 — Periodos definitivos de retención por cumplimiento normativo | decision | medium | M | platform | `08-decisiones-y-preguntas-abiertas.md` |
| [BACKLOG-009](#backlog-009) | P13 — Catálogo de funciones del responsable de aplicación | decision | medium | M | platform | idem |
| [BACKLOG-010](#backlog-010) | P14 — Catálogo definitivo de health-checks (métricas, umbrales, severidades) | decision | medium | M | platform | idem |
| [BACKLOG-011](#backlog-011) | P15 — Estrategia de migración de datos desde `.accdb` a PostgreSQL | decision | medium | L | platform | idem |
| [BACKLOG-012](#backlog-012) | P16 — Diseño detallado del ciclo UAT (workflow, visibilidad, entorno, aprobaciones, promoción) | decision | medium | M | platform | idem |
| [BACKLOG-013](#backlog-013) | P18 — UX del dashboard global de operaciones de notificación | decision | medium | M | platform | idem |
| [BACKLOG-014](#backlog-014) | P19 — Catálogo final de variantes de UI por módulo | decision | low | M | platform | idem |
| [BACKLOG-015](#backlog-015) | P20 — Contrato del proveedor de email corporativo | decision | low | S | platform | idem |
| [BACKLOG-016](#backlog-016) | D171 — Copy-paste RowSource bug en 6/7 forms G4 (Expedientes) | bug | critical | M | expedientes | `expedientes/epic.md` |
| [BACKLOG-017](#backlog-017) | D172 — `ComandoAyuda` mal etiquetado en 10/13 forms G4 (Expedientes) | bug | medium | S | expedientes | idem |
| [BACKLOG-018](#backlog-018) | D112 — Hardcode `EquipoID = 538` en Brass SLA repair | bug | medium | S | brass | `brass/epic.md` |
| [BACKLOG-019](#backlog-019) | HPS_Solicitudes — FK por email rompe data integrity | bug | high | M | hps-solicitudes | `hps-solicitudes/data-model.md` |
| [BACKLOG-020](#backlog-020) | Lanzadera — TYPO en `TbCuestionaroRespuestas` (falta «n») | chore | low | S | lanzadera | `lanzadera/data-model.md` |
| [BACKLOG-021](#backlog-021) | Lanzadera — Decisión sobre `Form_FormObtenerContraseña` (completar o eliminar) | decision | medium | S | lanzadera | `lanzadera/epic.md` CA-F8 |
| [BACKLOG-022](#backlog-022) | D95 — FK conceptual `idNCAsociada` en Condor | decision | low | S | condor | `condor/data-model.md` |
| [BACKLOG-023](#backlog-023) | D149 — OVERLAP en NavigationControl hosts (NoConformidades) | bug | medium | S | no-conformidades | `no-conformidades/epic.md` |
| [BACKLOG-024](#backlog-024) | Revisión final de las 8 épicas (estado DRAFT v0.1) | docs | high | M | root | Épicas mergeadas |
| [BACKLOG-025](#backlog-025) | Inventario Dysflow pendiente (NoConformidades, Condor, Brass) | chore | high | M | root | `docs/00-alcance-y-evidencia.md` |
| [BACKLOG-026](#backlog-026) | D156 — Lanzadera staging NO presente en `data/staging/` | chore | medium | L | root | `DOCS.md § Migración de binarios` |
| [BACKLOG-027](#backlog-027) | D84 — Brass sin `staging/` (clarificar rama baseline) | chore | medium | S | brass | idem |
| [BACKLOG-028](#backlog-028) | HPS_Solicitudes — Backend efectivo y Dysflow pendientes | chore | medium | M | hps-solicitudes | `docs/06-autorizacion-legacy-matriz.md` |
| [BACKLOG-029](#backlog-029) | Gestion_Riesgos — Segunda pasada Dysflow read-only pendiente | chore | medium | M | gestion-riesgos | `gestion-riesgos/data-model.md` |

**Total**: 29 items. Distribución por prioridad: critical 2 · high 6 · medium 13 · low 8.

---

## Detalle

### BACKLOG-001

- **Título**: Renombrar `data/staging/lanzaderas/` (plural) → `lanzadera/` (singular).
- **Tipo**: refactor.
- **Prioridad**: high.
- **Tamaño**: M (≤ 150 líneas tocadas).
- **Scope**: root.
- **Archivos afectados**: `data/staging/lanzaderas/` (rename filesystem), `.dysflow/project.json` (paths), `README.md` (línea 18), cualquier script o configuración que apunte a la ruta.
- **Criterios de aceptación**:
  - `git mv data/staging/lanzaderas data/staging/lanzadera` ejecutado.
  - `.dysflow/project.json` actualiza los paths al nombre singular.
  - `dysflow.get_capabilities({})` sigue listando Lanzadera.
  - `README.md` línea 18 muestra `data/staging/lanzadera/`.
  - `docs/03-aplicaciones/lanzadera/` (singular, ya existe) queda alineado con el staging.
- **Origen**: Auditoría docs Bloque 2 (PR #18 body).
- **Labels sugeridas**: `type:refactor`, `scope:root`, `size:M`, `priority:high`.
- **Notas**: bloqueado por BACKLOG-029 (segunda pasada Dysflow GR) si el refactor toca scripts. Coordinar con Natalia antes de renombrar.

### BACKLOG-002

- **Título**: Eliminar numeración al inicio en archivos de `docs/` (`00-`, `01-`, ..., `09-`).
- **Tipo**: refactor.
- **Prioridad**: high.
- **Tamaño**: L (150–400 líneas: afecta 15+ archivos y todas las referencias cruzadas).
- **Scope**: root.
- **Archivos afectados**: todos los archivos en `docs/` con prefijo numérico, más todos los docs que los referencian (épicas, CODEBASE-GUIDE, DOCS, README).
- **Criterios de aceptación**:
  - Cero archivos en `docs/` con prefijo `NN-`.
  - Las secciones temáticas (`topologia-ecosistema/`, `aplicaciones/`, `integraciones-y-operacion/`, `capacidades/`, `migracion/`) se mantienen como subcarpetas.
  - Los nombres pasan a `lowercase-with-hyphens.md` sin número inicial (ej. `00-alcance-y-evidencia.md` → `alcance-y-evidencia.md`).
  - Todos los cross-refs actualizados.
  - `documentation-alan-style §9` se cumple (grep `\d{2}-` en `docs/` devuelve 0).
- **Origen**: `documentation-alan-style §9` prohíbe explícitamente `01-readme.md`.
- **Labels sugeridas**: `type:refactor`, `scope:docs`, `size:L`, `priority:high`.
- **Notas**: refactor mayor. Requiere `git mv` con renombre y actualización en lote de todos los enlaces. Considerar hacerlo en commit único con PR separado.

### BACKLOG-003

- **Título**: Frontmatter YAML en los docs raíz (README, DOCS, AGENTS, CODEBASE-GUIDE, CONTRIBUTING, CHANGELOG).
- **Tipo**: chore.
- **Prioridad**: high.
- **Tamaño**: S (≤ 50 líneas).
- **Scope**: root.
- **Archivos afectados**: 6 archivos raíz.
- **Criterios de aceptación**:
  - Cada doc raíz tiene bloque `---` al inicio con: `description`, `globs`, `alwaysApply` (donde aplique).
  - `AGENTS.md` mantiene el formato actual (OpenCode requiere frontmatter para scope filtering).
  - `description` ≤ 100 caracteres por archivo.
- **Origen**: `documentation-alan-style §4` prescribe YAML por defecto.
- **Labels sugeridas**: `type:chore`, `scope:root`, `size:S`, `priority:high`.

### BACKLOG-004

- **Título**: Traducir al Castellano peninsular formal las secciones de docs raíz que quedaron en inglés técnico neutro.
- **Tipo**: docs.
- **Prioridad**: medium.
- **Tamaño**: L (afecta 4 docs raíz).
- **Scope**: root.
- **Archivos afectados**: README (`## Foundation`, `## Smoke test`, `## Herramientas`), DOCS (sección intro, "Scope del blueprint", "Walkthrough patterns"), CODEBASE-GUIDE (sección intro, "90-second mental model" lead, "Ownership", "Walkthrough patterns", "Workflow de contribución").
- **Criterios de aceptación**:
  - `Select-String` con patrón `\bthe\b|\byou\b|\bwe\b|\bour\b|\bscope\b` en headings y lead paragraphs devuelve ≤ 5 matches (los inevitables términos técnicos universales quedan en inglés).
  - Cada H2 y primer párrafo de sección está en castellano peninsular formal con usted.
  - La skill `documentation-alan-style §4` se aplica sin excepciones.
- **Origen**: Auditoría docs Bloque 2.
- **Labels sugeridas**: `type:docs`, `scope:root`, `size:L`, `priority:medium`.

### BACKLOG-005

- **Título**: Reemplazar referencias a `documentation-patterns` por `documentation-alan-style`.
- **Tipo**: docs.
- **Prioridad**: high.
- **Tamaño**: S.
- **Scope**: root.
- **Archivos afectados**: `DOCS.md` (sección "Skills de documentación"), `CODEBASE-GUIDE.md` (varias referencias en Quick map, Ownership, Workflow).
- **Criterios de aceptación**:
  - `grep -r 'documentation-patterns' docs/ DOCS.md CODEBASE-GUIDE.md` devuelve 0 matches fuera de menciones históricas explícitas.
  - Las menciones a la skill actual apuntan a `documentation-alan-style`.
- **Origen**: `documentation-patterns` está marcada como deprecada; `documentation-alan-style` la absorbe.
- **Labels sugeridas**: `type:docs`, `scope:root`, `size:S`, `priority:high`.

### BACKLOG-006

- **Título**: Dysflow #1408 — `map_form_behavior --json` rejected por codegraph-vba CLI.
- **Tipo**: bug (third-party).
- **Prioridad**: high.
- **Tamaño**: S.
- **Scope**: dysflow.
- **Criterios de aceptación**:
  - Issue [DysTelefonica/dysflow#1408](https://github.com/DysTelefonica/dysflow/issues/1408) cerrado.
  - `map_form_behavior` acepta `--json` y devuelve output JSON válido.
  - `walkthrough.json` schemas vuelven a registrar `codegraph_evidence_status: "ok"` en lugar de `"empty_bug_1408"`.
- **Origen**: `DOCS.md § Bugs dysflow filed`.
- **Labels sugeridas**: `type:bug`, `scope:dysflow`, `size:S`, `priority:high`, `external:blocked-by-third-party`.
- **Workaround actual**: usar `autoFetchCodeGraph: false` (aceptado en método v4).

### BACKLOG-007

- **Título**: Dysflow #1412 — `verify_form_bindings` RESULT_CONTRACT_VIOLATION.
- **Tipo**: bug (third-party).
- **Prioridad**: **critical**.
- **Tamaño**: S.
- **Scope**: dysflow.
- **Criterios de aceptación**:
  - Issue [DysTelefonica/dysflow#1412](https://github.com/DysTelefonica/dysflow/issues/1412) cerrado.
  - `verify_form_bindings` devuelve estructura JSON válida con `status: "ok"` o `status: "fail"` explícito.
  - **Bug D171 (BACKLOG-016) se detecta automáticamente** al re-correr walkthrough de Expedientes G4.
- **Origen**: `DOCS.md § Bugs dysflow filed`. CRÍTICO porque la tool rota habría detectado D171 sistémicamente.
- **Labels sugeridas**: `type:bug`, `scope:dysflow`, `size:S`, `priority:critical`.
- **Workaround actual**: skip con `status:"skipped_tool_broken"`. Aceptado en método v4, pero deja bugs reales sin detectar.

### BACKLOG-008

- **Título**: P10 — Periodos definitivos de retención por cumplimiento normativo o política de IT corporativa.
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Decisión D29 (PROVISIONAL: 90 días hot + 1 año total) se actualiza con valores definitivos.
  - Las decisiones D28 (mecanismo) y P10 quedan APROBADO.
  - Las tablas `TbLog*` legacy se retiran en favor de logs estructurados (D27).
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P10` + `09-arquitectura-objetivo-y-principios.md`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:medium`.
- **Bloquea**: fases SDD posteriores (compliance).

### BACKLOG-009

- **Título**: P13 — Catálogo de funciones exactas del «responsable de aplicación» por módulo.
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Para cada uno de los 8 módulos, tabla con: funciones de admin de aplicación, capacidades delegadas, alcance de permisos.
  - D22 (capacidades por módulo) se materializa con valores concretos.
  - D45 (capabilities declaradas) integra el catálogo por módulo.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P13`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:medium`.
- **Bloquea**: diseño de capabilities UI para módulos restantes.

### BACKLOG-010

- **Título**: P14 — Catálogo definitivo de health-checks (métricas, umbrales, severidades).
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Tabla con health-checks: nombre, métrica, umbral, severidad, módulo responsable, frecuencia.
  - CLI administrativo (D30) expone los checks bajo demanda y programada (D32).
  - Destinatarios de anomalías (D33) configurables por aplicación.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P14`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:medium`.
- **Bloquea**: operativa de plataforma post-lotes.

### BACKLOG-011

- **Título**: P15 — Estrategia de migración de datos desde `.accdb` a PostgreSQL.
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: L.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Plan de migración por módulo con: orden, dependencias, validaciones, rollback.
  - Estrategia Expand and Contract (D82) aplicada a cada migración de tabla.
  - Scripts de validación de integridad ejecutables.
  - Plan documentado en `docs/07-migracion/` (o su renombrado post-BACKLOG-002).
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P15`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:L`, `priority:medium`.
- **Bloquea**: inicio de fases SDD por módulo.

### BACKLOG-012

- **Título**: P16 — Diseño detallado del ciclo UAT (workflow, visibilidad, entorno, aprobaciones, promoción).
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Diagrama ASCII del workflow UAT.
  - Tabla de estados: nombre, transiciones permitidas, actor responsable.
  - Criterios de aceptación de cada estado.
  - D57 (Lanzadera) y D81 (coexistencia UAT+prod) materializados.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P16`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:medium`.
- **Bloquea**: governance de releases (D47-D50 ya APROBADO).

### BACKLOG-013

- **Título**: P18 — UX exacta del dashboard global de operaciones de notificación.
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Mockup HTML autocontenido (estilo `docs/design/mockups/lanzadera-shell.html`).
  - Tabla de widgets: nombre, dato, fuente, permiso requerido, acción.
  - D64 (dirección APROBADA) materializada con valores concretos.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P18`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:medium`.

### BACKLOG-014

- **Título**: P19 — Catálogo final de variantes de UI por módulo.
- **Tipo**: decision (ABIERTO).
- **Prioridad**: low.
- **Tamaño**: M.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Para cada módulo: vista única vs especializada, justificación, criterios materiales de diferencia.
  - D46 (vista única por defecto) respetada.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P19`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:M`, `priority:low`.

### BACKLOG-015

- **Título**: P20 — Contrato del proveedor de email corporativo (host, remitente, entregabilidad).
- **Tipo**: decision (ABIERTO).
- **Prioridad**: low.
- **Tamaño**: S.
- **Scope**: platform.
- **Criterios de aceptación**:
  - Variables documentadas: SMTP host, port, remitente, retry policy, headers personalizados.
  - Adaptador v1 (cola por tabla, D13) sustituye por el corporativo cuando se defina.
- **Origen**: `08-decisiones-y-preguntas-abiertas.md P20`.
- **Labels sugeridas**: `type:decision`, `scope:platform`, `size:S`, `priority:low`.

### BACKLOG-016

- **Título**: D171 — Copy-paste RowSource bug en 6/7 forms G4 (Expedientes).
- **Tipo**: bug.
- **Prioridad**: **critical**.
- **Tamaño**: M.
- **Scope**: expedientes.
- **Archivos afectados**: `Form_FormSuministradoresGestion`, `Form_FormLugarEjecucionGestion`, `Form_FormOficinasProgramaGestion`, `Form_FormOrganoContratacionGestion`, `Form_FormEjercitosGestion`, `Form_FormUsuariosGestion`.
- **Criterios de aceptación**:
  - Cada `ListaFiltrados.RowSource` apunta a la tabla correcta.
  - Validación visual con el usuario (Natalia) de que las columnas mostradas son las esperadas.
  - `verify_form_bindings` (BACKLOG-007 cerrado) detecta bug 0 en re-walkthrough.
- **Origen**: `docs/03-aplicaciones/expedientes/epic.md` § 2.2c.
- **Labels sugeridas**: `type:bug`, `scope:expedientes`, `size:M`, `priority:critical`.
- **Notas**: bug silencioso en producción — los usuarios llevan años conviviendo con columnas desplazadas. Duda si está causando daño real (PO-6 del epic).

### BACKLOG-017

- **Título**: D172 — `ComandoAyuda` mal etiquetado en 10/13 forms G4 (Expedientes).
- **Tipo**: bug.
- **Prioridad**: medium.
- **Tamaño**: S.
- **Scope**: expedientes.
- **Criterios de aceptación**:
  - Cada `ComandoAyuda` apunta al `Form_*` correcto.
  - Validación manual contra el original legacy.
- **Origen**: `expedientes/epic.md § 2.2c`.
- **Labels sugeridas**: `type:bug`, `scope:expedientes`, `size:S`, `priority:medium`.

### BACKLOG-018

- **Título**: D112 — Hardcode `EquipoID = 538` en `FormMaterialAltaReparacion` (Brass).
- **Tipo**: bug.
- **Prioridad**: medium.
- **Tamaño**: S.
- **Scope**: brass.
- **Criterios de aceptación**:
  - Confirmar con Natalia si el filtro fue intencional o bug histórico.
  - Si bug: filtro lee `EquipoID` del usuario/contexto en lugar de hardcoded.
- **Origen**: `brass/epic.md § 3` hallazgo H9.
- **Labels sugeridas**: `type:bug`, `scope:brass`, `size:S`, `priority:medium`.
- **Bloquea**: validación UAT de Brass.

### BACKLOG-019

- **Título**: HPS_Solicitudes — FK por email (`TbResponsables.Correo → TbSolicitudes.emailResponsable`) rompe data integrity.
- **Tipo**: bug.
- **Prioridad**: high.
- **Tamaño**: M.
- **Scope**: hps-solicitudes.
- **Criterios de aceptación**:
  - En PostgreSQL: agregar columna `idResponsable` (FK numérica a `TbResponsables.ID`).
  - Backfill de la FK por email matching.
  - Email se mantiene como campo independiente (no FK).
  - Validación: cambio de email en `TbResponsables` ya no rompe la relación lógica.
- **Origen**: `hps-solicitudes/data-model.md § Notas`.
- **Labels sugeridas**: `type:bug`, `scope:hps-solicitudes`, `size:M`, `priority:high`.

### BACKLOG-020

- **Título**: Lanzadera — TYPO en `TbCuestionaroRespuestas` (falta «n» en Cuestionario).
- **Tipo**: chore.
- **Prioridad**: low.
- **Tamaño**: S.
- **Scope**: lanzadera.
- **Criterios de aceptación**:
  - Tabla legacy renombrada a `TbCuestionarioRespuestas` (con n).
  - Todas las queries y formularios legacy actualizados.
  - **Decisión**: el legacy retira cuestionarios (D51) — evaluar si el renombre tiene sentido o se archiva sin migración.
- **Origen**: `lanzadera/data-model.md` línea 348.
- **Labels sugeridas**: `type:chore`, `scope:lanzadera`, `size:S`, `priority:low`.
- **Notas**: probablemente quede archivado por D51 (Lanzadera retira formación/cuestionarios).

### BACKLOG-021

- **Título**: Lanzadera — Decisión sobre `Form_FormObtenerContraseña` (completar o eliminar).
- **Tipo**: decision (ABIERTO).
- **Prioridad**: medium.
- **Tamaño**: S.
- **Scope**: lanzadera.
- **Criterios de aceptación**:
  - Natalia confirma: completar walkthrough del form, o eliminarlo.
  - Si eliminar: documentar la razón en `lanzadera/epic.md`.
  - Si completar: walkthrough v4 con método v4 (post BACKLOG-007).
- **Origen**: `lanzadera/epic.md` CA-F8.
- **Labels sugeridas**: `type:decision`, `scope:lanzadera`, `size:S`, `priority:medium`.

### BACKLOG-022

- **Título**: D95 — FK conceptual `idNCAsociada` en Condor (vinculación con NoConformidades).
- **Tipo**: decision (ABIERTO).
- **Prioridad**: low.
- **Tamaño**: S.
- **Scope**: condor.
- **Criterios de aceptación**:
  - Decidir: FK numérica directa, FK lógica por texto, o eliminar vínculo.
  - D95 marcada APROBADO con la decisión.
- **Origen**: `condor/capabilities.md § F-XX`.
- **Labels sugeridas**: `type:decision`, `scope:condor`, `size:S`, `priority:low`.

### BACKLOG-023

- **Título**: D149 — OVERLAP en NavigationControl hosts (NoConformidades).
- **Tipo**: bug.
- **Prioridad**: medium.
- **Tamaño**: S.
- **Scope**: no-conformidades.
- **Criterios de aceptación**:
  - En web: nested routes / tabs eliminan el overlap de subform regions.
  - Walkthrough v3 re-ejecutado muestra 0 OVERLAP.
- **Origen**: `no-conformidades/epic.md` hallazgo H4.
- **Labels sugeridas**: `type:bug`, `scope:no-conformidades`, `size:S`, `priority:medium`.
- **Notas**: en web desaparece por arquitectura, no requiere fix legacy. Marcar como "resuelto por migración".

### BACKLOG-024

- **Título**: Revisión final de las 8 épicas (estado DRAFT v0.1).
- **Tipo**: docs.
- **Prioridad**: high.
- **Tamaño**: M.
- **Scope**: root.
- **Archivos afectados**: 8 épicas en `docs/03-aplicaciones/<app>/epic.md`.
- **Criterios de aceptación**:
  - Cada épica cambia su header de `DRAFT v0.1` a estado consolidado.
  - Cada épica revisa que su sección "Pendientes operacionales" esté actualizada.
  - Las nuevas decisiones D88-D178 están integradas.
- **Origen**: 8 épicas mergeadas con estado DRAFT.
- **Labels sugeridas**: `type:docs`, `scope:root`, `size:M`, `priority:high`.

### BACKLOG-025

- **Título**: Inventario Dysflow pendiente (NoConformidades, Condor, Brass).
- **Tipo**: chore.
- **Prioridad**: high.
- **Tamaño**: M.
- **Scope**: root.
- **Archivos afectados**: `data/staging/no-conformidades/`, `data/staging/condor/`, `data/staging/brass/` + `docs/00-alcance-y-evidencia.md`.
- **Criterios de aceptación**:
  - `dysflow.list_objects` devuelve el inventario completo de cada backend.
  - Los huecos en `docs/00-alcance-y-evidencia.md` tabla 3, 4, 7 se cierran.
- **Origen**: `docs/00-alcance-y-evidencia.md`.
- **Labels sugeridas**: `type:chore`, `scope:root`, `size:M`, `priority:high`.

### BACKLOG-026

- **Título**: D156 — Lanzadera staging NO presente en `data/staging/lanzaderas/`.
- **Tipo**: chore.
- **Prioridad**: medium.
- **Tamaño**: L.
- **Scope**: root.
- **Criterios de aceptación**:
  - Binario Lanzadera (.accdb) presente en staging.
  - `data/staging/lanzadera/frontend/Lanzadera.accdb` (post-BACKLOG-001) listo para walkthrough.
  - Walkthrough audit de queries/macros ejecutable desde staging.
- **Origen**: `DOCS.md § Migración de binarios`.
- **Labels sugeridas**: `type:chore`, `scope:root`, `size:L`, `priority:medium`.
- **Bloquea**: walkthroughs Lanzadera adicionales.

### BACKLOG-027

- **Título**: D84 — Brass sin `staging/` (clarificar rama baseline `develop` vs `release_2026-001`).
- **Tipo**: chore.
- **Prioridad**: medium.
- **Tamaño**: S.
- **Scope**: brass.
- **Criterios de aceptación**:
  - Usuario confirma rama baseline.
  - Snapshot de Brass en `data/staging/brass/` desde la rama correcta.
- **Origen**: `DOCS.md § Migración de binarios` + D84.
- **Labels sugeridas**: `type:chore`, `scope:brass`, `size:S`, `priority:medium`.

### BACKLOG-028

- **Título**: HPS_Solicitudes — Backend efectivo y Dysflow pendientes.
- **Tipo**: chore.
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: hps-solicitudes.
- **Criterios de aceptación**:
  - Backend `Solicitudes_HPS_Datos.accdb` localizado y vinculado.
  - `.dysflow/project.json` con entrada para HPS_Solicitudes (D83).
  - CodeGraph-VBA local inicializado para HPS_Solicitudes.
- **Origen**: `docs/06-autorizacion-legacy-matriz.md` + `docs/01-inventario-aplicaciones.md`.
- **Labels sugeridas**: `type:chore`, `scope:hps-solicitudes`, `size:M`, `priority:medium`.

### BACKLOG-029

- **Título**: Gestion_Riesgos — Segunda pasada Dysflow read-only pendiente.
- **Tipo**: chore.
- **Prioridad**: medium.
- **Tamaño**: M.
- **Scope**: gestion-riesgos.
- **Criterios de aceptación**:
  - `dysflow.get_capabilities` confirma read-only contra `Gestion_Riesgos_Datos.accdb`.
  - Inventario completo de tablas y columnas cosechado.
  - Perfilado de filas por tabla completado.
  - Volúmenes TBD actualizados a valores reales en `gestion-riesgos/data-model.md`.
- **Origen**: `gestion-riesgos/data-model.md` línea 5 + 65.
- **Labels sugeridas**: `type:chore`, `scope:gestion-riesgos`, `size:M`, `priority:medium`.

---

## Cómo extraer un issue desde este backlog

Una IA puede convertir cada BACKLOG-NNN en un issue de GitHub con la siguiente plantilla:

```text
Título: docs(<scope>): <BACKLOG-NNN título>
Labels: <labels sugeridas>
Body:
  ## Contexto
  <Descripción del item>
  ## Archivos
  <Archivos afectados>
  ## Criterios de aceptación
  <Lista de criterios de aceptación>
  ## Origen
  <Enlace a este doc + sección>
  ## Tamaño estimado
  <S / M / L / XL>
```

Si la IA detecta que un item tiene criterios de aceptación ambiguos o scope muy amplio, debe proponer **partirlo en varios issues** (uno por criterio de aceptación, idealmente) y reportar la partición.

## Reglas del backlog

- Cada item tiene un ID estable (`BACKLOG-NNN`). No se reasignan IDs.
- Un item se considera cerrado cuando su issue referenciado se cierra con merge del PR.
- Nuevos items se añaden al final con el siguiente ID correlativo.
- Items que se parten generan items hijos (`BACKLOG-NNNa`, `BACKLOG-NNNb`).
- El backlog se revisa en cada cierre de lote (cada vez que una épica sale de DRAFT).

[← Back to README](../README.md) · [Next: DOCS.md →](../DOCS.md)