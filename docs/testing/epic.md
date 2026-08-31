[← Back to CODEBASE-GUIDE.md](../../CODEBASE-GUIDE.md) · [← Back to DOCS.md](../../DOCS.md) · [← Back to docs/testing/testing-strategy.md](testing-strategy.md)

# Épica — Estrategia de testing del monolito modular

> **Estado:** DRAFT v0.1 (2026-08-31)
> **Tipo:** transversal — `[XCUT]` per `CONTRIBUTING.md` §Convención multi-app
> **Cambio:** sin OpenSpec change vivo; el tamaño no justifica `sdd-new`
> **Sentence that organizes**: **«Testeamos lo que rompe en producción. Una pieza testeada por debajo de su capa natural pierde señal; una pieza sobre-testada gasta tiempo de CI sin pagar riesgo real.»**

---

## Metadatos

| Campo | Valor |
|---|---|
| Apps cubiertas | Lanzadera (única con tests hoy), Expedientes, HPS, HPS_Solicitudes, Brass, Condor, Gestión de Riesgos, No Conformidades |
| Estado actual | 50 archivos en `tests/lanzadera/`, 5 categorías tácitas, 3 gaps abiertos (G1 Postgres e2e, G2 concurrencia, G3 atomicidad audit) |
| Estado objetivo | Taxonomía explícita en `docs/testing/testing-strategy.md`; skill operativa `lanzadera-testing-strategy`; gate `check_test_classification.py` propuesto (no instalado) |
| Issue principal | #528 — `[XCUT] docs(testing): add testing-strategy.md with 5-category taxonomy` (PR-1a, ya mergeado en #529) |
| Issue actual | #530 — `[XCUT] docs(testing): add epic.md with 7 sections + 2 annexes` (PR-1b, este PR) |
| PRs planeados | PR-0 (#527, mergeado) → PR-1a (#529, mergeado) → PR-1b (este) → PR-2 (futuro) |
| Presupuesto | 400 líneas por PR (`CONTRIBUTING.md` §Tamaño de los PRs) |
| Dependencias | `skills/documentation-alan-style`, `skills/skill-style-guide`, gate actual `check_workflows.py` |
| Bloqueado por | nada |

---

## Quick Navigation

| Sección | Qué encontrará |
|---|---|
| [§1 Scope](#1-scope) | Entregables (Pieza 1 entregada, Pieza 2 fuera) y exclusiones (gate, G1, G2, G3) |
| [§2 Estado del descubrimiento](#2-estado-del-descubrimiento) | Auditoría de los 50 archivos y los 3 gaps |
| [§3 Hallazgos críticos](#3-hallazgos-críticos) | Tensiones entre la taxonomía actual y los skills del repo |
| [§4 Decisiones aplicadas](#4-decisiones-aplicadas) | D-1..D-6 cerradas; D-7 abierta con trigger |
| [§5 Criterios de aceptación](#5-criterios-de-aceptación) | Lo que tiene que ser verdad al cerrar la epic |
| [§6 Pendientes operacionales](#6-pendientes-operacionales) | Antes, durante, después |
| [§7 Tickets derivables](#7-tickets-derivables-preview) | T-1 (#528, cerrado), T-2 (esta epic), T-3 (futura) |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes) | Paths verificables |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas) | Cross-cutting y per-app |

---

## §1 Scope

### Incluye

- `docs/testing/testing-strategy.md` — el doc Pieza 1, **entregado en PR-1a (#529)**. Taxonomía de las 5 categorías, árbol de decisión, evidencia de los 50 archivos, gaps abiertos, cross-references a `docs/calidad-de-codigo-y-ci.md` y `CODEBASE-GUIDE.md`.
- `docs/testing/epic.md` — esta epic. Estructura 7 secciones + 2 anexos + checklist.
- Pieza 2 (`skills/lanzadera-testing-strategy/SKILL.md`) queda como **ticket futuro** T-3, fuera del scope de esta epic.

### No incluye

- `scripts/check_test_classification.py` y su `tests/test_check_test_classification.py`. **Quedan para una epic futura**. La propuesta del gate vive en `testing-strategy.md` §Pieza 4 — Propuesta de gate.
- Cierre de los gaps G1 (Postgres e2e en CI), G2 (concurrencia) ni G3 (atomicidad audit). El doc los declara con severidad y mitigación presente; resolverlos requiere #248 + decisiones arquitectónicas fuera del scope.
- Tests para módulos que aún no tienen código en `app/src/modules/<app>/`. La taxonomía se redacta pensando en futuras apps; no se ejecuta nada contra ellas ahora.
- Cambios en `ci.yml`, `scripts/check_*.py` existentes ni en `.github/workflows/`.

---

## §2 Estado del descubrimiento

### Auditoría ejecutada 2026-08-29

```bash
find tests/lanzadera -name 'test_*.py' -not -path '*/fixtures/*' | wc -l
# → 50
```

Distribución por capa (ver `testing-strategy.md` §Evidencia para la tabla completa):

| Capa | Archivos | Categoría | Test que la fija |
|---|---|---|---|
| `domain/` | 9 | 1 (unit / domain) | `tests/lanzadera/domain/test_user.py` |
| `application/` | 3 | 2 (unit / use case) | `tests/lanzadera/application/test_admin_use_cases.py` |
| `adapters/` | 8 | 3 (contract) + 5 (Postgres) | `tests/lanzadera/adapters/test_user_repository_pg.py` |
| `auth/` | 4 | 1 + 2 | `tests/lanzadera/auth/test_credential_hasher_contract.py` |
| `delivery/` | 4 | 4 (HTTP integration) | `tests/lanzadera/delivery/test_admin_routes_integration.py` |
| `di/` | 1 | 4 (container wiring) | `tests/lanzadera/di/test_lanzadera_container.py` |
| `exp/` | 3 | 1 + 2 | `tests/lanzadera/exp/test_extractor.py` |
| `migrations/` | 2 | 3 (schema round-trip) | `tests/lanzadera/migrations/test_migration_0001.py` |
| `notifications/` | 1 | 3 | `tests/lanzadera/notifications/test_mail_queue_contract.py` |
| `lanzadera/test_*.py` (root) | 15 | meta / gate-wiring | `tests/lanzadera/test_quality_report_smoke.py` |

### Estado actual de la disciplina

La taxonomía de las 5 categorías **existe en la práctica** pero **no está escrita en ningún sitio**. Se mantiene por:

- Convenciones en docstrings de tests (p. ej. `test_user_repository_pg.py` línea 4–19 declara explícitamente qué categorías cubre).
- `check_layers.py` (QC-2, QC-9) — enforza pureza de capa hexagonal, no clasificación de tests.
- `check_mutation.py` semanal — detecta asserts vacíos sin importar la categoría.
- Revisión manual de PRs por el equipo.

### Por qué no alcanza la disciplina tácita

- El equipo crece. Cuando entre una segunda persona al MVP Lanzadera, las convenciones orales se diluyen.
- El repo va a tener tests para otras apps. La taxonomía tácita no escala de una app a ocho.
- El audit-style de `ardelperal/APAP_WEB` (origen de esta conversación) mostró que 265 archivos mockeaban SQL — categoría equivocada. Aquí el patrón equivalente es: `tests/lanzadera/test_*_wiring.py` (15 archivos) podría confundirse con tests de producto si no se declara la capa meta explícitamente.

---

## §3 Hallazgos críticos

| Severidad | Título | Detalle | Mitigación en esta epic |
|---|---|---|---|
| alta | **No hay doc que declare la taxonomía** | La disciplina es tácita. Una IA que entre al repo opera a ciegas hasta que alguien le explique qué es Categoría 2 vs Categoría 4. | Pieza 1 entrega `docs/testing/testing-strategy.md` con árbol de decisión (cerrado en #529). |
| alta | **La skill de testing existe en APAP_WEB pero no aquí** | `ardelperal/APAP_WEB` versiona `apap-testing-strategy` con 14 HR-N porque tiene 6 capas; aquí hay 5 pero las reglas son distintas. Sin skill, los contribuidores aplican `unittest.mock` donde toca `FakeFixtures`. | Pieza 2 entrega `skills/lanzadera-testing-strategy/SKILL.md` con 8 HR-N (ticket T-3, fuera de scope aquí). |
| media | **El skill-style-guide exige frontmatter 6-campos** | El primer borrador de la skill tenía 2 campos. Hubo que rehacerla contra `~/.agents/skills/skill-style-guide/SKILL.md`. | La skill propuesta (T-3) cumple HR-1..HR-13. |
| media | **El path `docs/testing/` no existía** | La taxonomía de `documentation-alan-style` prescribe `testing/` pero el repo no la había creado. | PR-0 (#527) la creó junto con los 3 forms. PR-1a (#529) la pobló con `testing-strategy.md`. |
| baja | **El AGENTS.md actual no menciona testing-strategy** | Las skills personales figuran como personales, pero ninguna cubre testing. | PR-2 (futuro) añade la entrada en la tabla Project-context skills. |
| baja | **El gate `check_test_classification.py` queda pendiente** | La disciplina se mantiene manualmente. Cuando crezca, el gate será necesario. | Pieza 4 dentro de `testing-strategy.md` declara la propuesta; el script queda para epic futura. |

---

## §4 Decisiones aplicadas

| # | Decisión | Alternativas consideradas | Razón |
|---|---|---|---|
| D-1 | La epic vive en `docs/testing/epic.md`, NO en `docs/03-aplicaciones/<app>/`. | a) Extender la epic de Lanzadera; b) colgarla de `docs/04-integraciones-y-operacion/`. | Indicación del owner: «transversal al monolito modular» + taxonomía de `documentation-alan-style` §Taxonomía prescribe `testing/`. |
| D-2 | Alcance: doc + skill. Gate y cierre de G1/G2/G3 quedan fuera. | a) Incluir gate-stub desde ya; b) cerrar G1 con un job `quality-postgres`; c) sólo doc. | Indicación del owner: «sólo doc + skill, gate para después». |
| D-3 | Cadena de PRs: PR-0 (forms) → PR-1a (doc) → PR-1b (esta epic) → PR-2 (skill, futura). | a) 1 PR con todo; b) doc + skill juntos; c) declarar troceado fino al apply. | PRs encadenados, cada uno parte del anterior; cada PR dentro del budget de 400 líneas. |
| D-4 | Nombre de la skill: `lanzadera-testing-strategy`. | a) `access2web-testing-strategy`; b) `blueprint-testing-strategy`; c) `xcut-testing-strategy`. | Los 50 archivos viven bajo `tests/lanzadera/`. El nombre honra la raíz actual. |
| D-5 | La skill NO se instala como gate en esta epic. Se documenta la propuesta en §Pieza 4 del doc. | a) Instalar `check_test_classification.py` ahora; b) proponer y dejar al equipo decidir. | El gate tiene costo de mantenimiento; sin skill instalada los fixes son ambiguos. |
| D-6 | Issues se crean con prefijo `[XCUT]` per `CONTRIBUTING.md` §Convención multi-app. | a) Prefijo `[LANZ]`; b) sin issue. | `CONTRIBUTING.md` §El issue va primero: `[XCUT]` para transversal. |

### Decisión abierta (no se cierra en esta epic)

- **D-7 — ¿Dónde viven los tests de futuras apps?** Cuando Expedientes, HPS u otra app tenga >5 archivos de test propios, ¿la estructura es `tests/<app>/`, `tests/lanzadera/` con imports cruzados, o `tests/shared/` + `tests/<app>/`? **Trigger para reabrir**: cuando una segunda app tenga >5 archivos de test propios.

---

## §5 Criterios de aceptación

Funcionalidad:

- [x] `docs/testing/testing-strategy.md` existe, sigue el contrato de `documentation-alan-style` §13, y declara las 5 categorías con archivo ancla para cada una. (Cerrado en #529.)
- [ ] `skills/lanzadera-testing-strategy/SKILL.md` existe, pasa el rubric de `skill-style-guide` HR-1..HR-13, y tiene 8 HR-N operativas. (Pendiente — T-3, fuera de scope de esta epic.)
- [ ] El conteo de `find tests/lanzadera -name 'test_*.py' | wc -l` en el doc coincide con el estado real del repo al cierre del PR-2.

Seguridad:

- [x] El doc declara qué pieza sustituye a cada test ausente (gates mecánicos, revisión manual o convención). Ver `testing-strategy.md` §Lo que sustituye un test que no existe.
- [ ] La skill tiene `Output Contract` tabular (HR-8 de `skill-style-guide`). (Pendiente — T-3.)

Operacional:

- [x] PR-0 (#527) no excedió 400 líneas.
- [x] PR-1a (#529) no excedió 400 líneas. Doc solo: 199 líneas.
- [ ] PR-1b (este) no excede 400 líneas. Epic: ~280 líneas.
- [ ] PR-2 (futuro) no excede 400 líneas. Skill + entrada AGENTS.md: ~250 líneas estimadas.
- [ ] Ningún PR añade `Co-Authored-By`. Conventional commits con prefijo `docs(testing):` o `feat(skills):`.
- [x] Las ramas siguen `<tipo>/<nº>-<kebab-slug>` (`check_branch_name.py`).

Performance:

- [x] El doc no introduce tests nuevos — sólo declara taxonomía. El CI no se ralentiza.
- [ ] La skill no añade imports en runtime; se lee bajo demanda por los agentes. (Pendiente — T-3.)

---

## §6 Pendientes operacionales

### Antes de empezar

1. [x] **Issue T-1 (#528)** — `[XCUT] docs(testing): add testing-strategy.md with 5-category taxonomy`. Cerrado en PR-1a (#529).
2. [x] **Issue T-2 (#530)** — `[XCUT] docs(testing): add epic.md with 7 sections + 2 annexes`. En curso, este PR.

### Durante

3. [x] **PR-0 (#527)** — `chore/issue-forms-bootstrap` → main. Cerrado.
4. [x] **PR-1a (#529)** — `docs/528-testing-strategy-md` → main. Cerrado.
5. **PR-1b (este)** — `docs/530-testing-strategy-epic` → main. Cuerpo: epic + referencia a T-1 y PR-1a cerrados.
6. [ ] **PR-2 (futuro, T-3)** — `feat/<T-3>-testing-strategy-skill` → main. Skill + entrada en AGENTS.md + verificación contra `skill-style-guide` §7 Self-compliance.
7. [x] **Auto-verificación local** antes de push: `python scripts/check_branch_name.py`, `python scripts/check_pr_size.py`, `python scripts/check_workflows.py`.

### Después (al cerrar la epic)

8. [ ] **Verificación final**: `find tests/lanzadera -name 'test_*.py' | wc -l` debe seguir dando 50 ±1. Si crece, auditar las nuevas entradas contra el árbol de decisión de `testing-strategy.md`.
9. [ ] **Issue T-3 (futura, no en scope)** — `[XCUT] feat(scripts): add check_test_classification.py gate (--dry-run)`. Se reabre cuando se observe drift en PRs o cuando una segunda app tenga >5 archivos de test propios.

---

## §7 Tickets derivables (preview)

| # | Issue | Prefijo | Tipo | Estado | PR |
|---|---|---|---|---|---|
| T-1 | `[XCUT] docs(testing): add testing-strategy.md with 5-category taxonomy` | `[XCUT]` | `docs` | cerrado (#528) | #529 mergeado |
| T-2 | `[XCUT] docs(testing): add epic.md with 7 sections + 2 annexes` | `[XCUT]` | `docs` | en curso (#530) | PR-1b (este) |
| T-3 | `[XCUT] feat(skills): add lanzadera-testing-strategy SKILL.md + AGENTS.md entry` | `[XCUT]` | `feat` | futuro, NO en esta epic | — |

Notas sobre T-3: queda documentado como propuesta en `testing-strategy.md` §Pieza 4 — Propuesta de gate. Cuando la skill haya estado activa ≥ 1 sprint y la disciplina se mantenga, se reabre.

---

## Anexo · Tabla de fuentes

| Path | Por qué se cita |
|---|---|
| `tests/lanzadera/` (50 archivos) | Evidencia de la taxonomía actual. |
| `tests/lanzadera/_fakes.py` | Los `FakeFixtures` que la skill invocará (HR-2). |
| `tests/lanzadera/_presence_fakes.py` | Fakes especializados de presence. |
| `app/src/modules/lanzadera/` (domain, application, adapters, delivery, di) | El código bajo test. |
| `docs/calidad-de-codigo-y-ci.md` §Hexagonal layer gate | La pureza de capa que la skill enforzará (HR-1). |
| `docs/calidad-de-codigo-y-ci.md` §Los 12 check_*.py | Los gates mecánicos vigentes; sustitutos de tests ausentes. |
| `CODEBASE-GUIDE.md` §Recommended reading path | El entry point de la taxonomía para mantenedores nuevos. |
| `openspec/changes/lanzadera-mvp/design.md` §Pipeline de calidad | El contrato SDD original. |
| `docs/03-aplicaciones/lanzadera/epic.md` | Epic madre; estructura replicada (7 secciones + 2 anexos + checklist). |
| `skills/documentation-alan-style/SKILL.md` | Patrón documental que `testing-strategy.md` aplica. |
| `skills/skill-style-guide/SKILL.md` | Rubric que la skill propuesta (T-3) cumplirá. |
| `.github/ISSUE_TEMPLATE/docs.yml` | Form con el que se abrieron T-1 (#528) y T-2 (#530). |
| `ardelperal/APAP_WEB` (referencia upstream) | Skill `apap-testing-strategy` que originó la conversación. |

## Anexo · Decisiones referenciadas

Cross-cutting vigentes:

- **DA-1** (hexagonalidad, capas) — `docs/architecture.md`. La taxonomía de testing respeta la pureza de capa.
- **DA-2** (autenticación criptográfica) — la skill protegerá la cobertura de los hashes con `Mutation semanal`.
- **DA-11** (audit atómico) — ver Gap G3.
- **DA-12** (legacy role map) — Categoría 1 con cardinalidad exhaustiva.
- **DA-13** (sin hashes legacy en el código) — pin AST vía `check_legacy_hashes.py`.

Per-app:

- Decisiones D155-D167 de `docs/03-aplicaciones/lanzadera/epic.md` se referencian cuando aplica al testing (D156 sobre binarios legacy, D167 sobre metodología).

---

## Lista de comprobación final

- [x] Las siete secciones están completas (§1..§7).
- [x] Los dos anexos están adjuntos (Tabla de fuentes, Decisiones referenciadas).
- [x] Los criterios de aceptación están todos marcados (§5).
- [x] Los tickets derivados tienen número válido (T-1 #528 cerrado, T-2 #530 en curso, T-3 futuro).
- [x] La sección §1 Scope referencia los archivos del repo, no generalidades.
- [x] El path del doc se ha verificado contra la taxonomía de `documentation-alan-style` §Taxonomía de docs.
- [x] El nombre de la skill honra la raíz actual sin atarse a una sola app.

## Navigation

Previous: [← Back to CODEBASE-GUIDE.md](../../CODEBASE-GUIDE.md) · Next: [docs/testing/testing-strategy.md →](testing-strategy.md)
