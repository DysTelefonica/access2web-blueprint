# Issue #807 — Fichas de capacidad faltantes de Expedientes (37 de 63)

## Outcome

Cubrir las **37 fichas de capacidad faltantes** que prescribe `openspec/changes/expedientes-web-migration/tasks.md` para Expedientes, en la ruta canónica ya usada por las 26 fichas existentes (`docs/03-aplicaciones/expedientes/capabilities/CAP-NNN-slug.md`). El catálogo completo EXP-CAP-001..063 queda enlazado desde `docs/03-aplicaciones/expedientes/capabilities/index.md` y la issue #807 queda cerrada mediante la cadena de 8 PRs encadenados.

## Branch & cadena

- **Rama inicial** de cada PR: `docs/807-fichas-capacidad-<n>` (1..8) sobre `origin/main @ 55b07e8` (HEAD actual de `origin/main` a 2026-10-22).
- **PR body**: `Refs #807` en los intermedios (PR-1..PR-7) y `Closes #807` en la punta (PR-8). Sección fija `## Asistencia de IA` con la frase registrada por el propietario #962.
- **Label**: `chain:partial` en los ocho al crear con `gh pr create --label chain:partial`.
- **≤ 400 líneas por PR** (contando additions + deletions, como exige el resto de las cadenas del repo).

## Distribución de las 37 fichas

| PR | Fichas | Tema | Estimación (~L) |
|----|--------|------|----------------|
| 1 | CAP-001, 002, 003, 006, 007 | Escritura / jerarquía / ordinal funcional | ~300 |
| 2 | CAP-008, 009, 010, 011, 012 | Related-data I (hitos → responsables) | ~300 |
| 3 | CAP-013, 014, 015, 023, 025 | Related-data II + Comerciales + Órganos + Bandeja | ~300 |
| 4 | CAP-040, 043, 044, 045, 046 | Destino E2E + access-control hub | ~300 |
| 5 | CAP-047, 048, 049, 050, 051, 052 | Runtime + integrations I (HPS, AGEDYS) | ~360 |
| 6 | CAP-053, 054, 055, 056 | Integrations II (Riesgos, NC, correo, documentos) | ~240 |
| 7 | CAP-057, 058, 059, 060, 061 | Retirement I (Win32 red, procesos, OLE, globals, selector) | ~300 |
| 8 | CAP-062, 063 + `index.md` | Retirement II + actualización del índice (Closes #807) | ~190 |

## Specs fuente y decisiones citadas

Las 63 capacidades tienen su requisito en `openspec/changes/expedientes-web-migration/specs/`, agrupadas así:

| Spec | Caps |
|------|------|
| `lifecycle.md` | CAP-001..007 |
| `related-data.md` | CAP-008..014 |
| `catalogs.md` | CAP-015..024 |
| `query-and-tasks.md` | CAP-025..029 |
| `write-resilience.md` | CAP-030..032 |
| `e2e.md` | CAP-033..042 |
| `access-control.md` | CAP-043..046 |
| `runtime.md` | CAP-047..050 |
| `integrations.md` | CAP-051..056 |
| `uat-cutover-legacy-retirement.md` | CAP-057..063 |

El catálogo normativo (disposición, linaje y obligaciones) vive en `docs/03-aplicaciones/expedientes/capability-ledger.md`. Las D-/DA- vinculantes siguen `docs/architecture.md`.

## Formato de cada ficha

Se copia literalmente el patrón de las 26 fichas existentes (verbatim de `CAP-004-cambio-tipo.md` y `CAP-005-estado-garantia.md`), con tres adiciones explícitas para cumplir el criterio de trazabilidad del mantenedor:

1. Frontmatter `[← Back to capabilities-index](index.md) · [← Back to DOCS](../../../DOCS.md)` (igual que las 26).
2. `# CAP-NNN — Título` copiando el título del ledger (slug limpio, sin acentos y con guiones).
3. Cuerpo estándar: párrafo ancla al spec → `## Escenario clave` → `## Cómo se aplica a access2web-blueprint` con bullets de D-/DA- relevantes → `## Core invariants` (3 viñetas literales) → `## Contributor checklist` (3 checks literales) → `## Lista de comprobación final` (3 checks literales) → `## Trazabilidad al código fuente` (NUEVO, abajo detallado) → `## Navigation` (Previous/Next).
4. **`## Trazabilidad al código fuente`** (sección añadida): tres sub-bloques copiados literalmente a las plantillas del repositorio:
   - **Spec fuente** con el anchor al Requirement de OpenSpec (verificado, `git grep` resuelve).
   - **Catálogo PRD** apuntando a la fila de `docs/03-aplicaciones/expedientes/capability-ledger.md` para esa `EXP-CAP-NNN` (verificado, mismo repo).
   - **Repo Expedientes** con `ruta:línea` y commit fijado `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` del staging legacy; marcado `Hipótesis — no verificada desde este entorno` cuando la línea no aparece en los walkthroughs de `data/staging/expedientes/docs/capabilities/` que el repo sí contiene (los únicos verificables de forma estática).
5. `## Navigation` solo enlaza ficheros que existen en la propia rama (HR-66): fichas del propio PR o ya existentes en `main`.

## Pendiente (decisión de producto, NO resuelta aquí)

La issue #807 tiene una decisión abierta:

> _Pendiente: elegir la ruta canónica — la que prescribe `tasks.md` (`docs/05-capacidades/expedientes/EXP-CAP-00X.md`) o la que ya usan 26 fichas (`docs/03-aplicaciones/expedientes/capabilities/CAP-00X-slug.md`) — y decidir si los hubs D51–D54 se quedan en `docs/05` o se mueven._

Esta tarea **no la resuelve** (per orden del propietario: «decisiones de producto abiertas (P-EXP-*) se citan, no se resuelven»). Se cita en la `## Trazabilidad` de la última ficha nueva (CAP-063) y en el cuerpo del PR-8, pero no fuerza un rename ni mueve hubs.

## Hard rules

- Cada PR es creado con `gh pr create --base main --head docs/807-fichas-capacidad-<n> --label chain:partial --title "..." --body "..."` y debe usar la rama literal (sin `HEAD`, sin variables en `git push`).
- Conventional commit sin atribución de IA: `docs(exp): CAP-001..007 fichas de capacidad (Refs #807, PR 1/8)`, etc.
- No rebase ni force. Merge posterior con `git merge origin/main` por el owner; cada PR individual.
- No push a `main`. No se borran ramas remotas. El merge lo hace el owner.
- Máximo 2 PRs en CI simultáneamente: si una PR todavía está corriendo CI, la siguiente se empuja solo cuando la anterior esté `MERGEABLE` (gate de cadena del repo per `bp-fleet-chain.sh`).
- No usar `git commit --author` ni co-author con IA (per orden #962).
- Cada ficha enlaza solo a ficheros que existan en la rama del PR (HR-66): las vecinas del propio PR o las del main pre-existente.

## Verificación local pre-PR (cada PR)

1. `cd docs/03-aplicaciones/expedientes/capabilities/ && for f in CAP-*.md; do grep -q "^## Spec fuente\|^### Spec fuente" "$f" || echo "FALTA SPEC en $f"; grep -q "^## Trazabilidad" "$f" || echo "FALTA TRAZ en $f"; done` — comprueba secciones obligatorias.
2. `git grep "^### Requirement: EXP-CAP" openspec/changes/expedientes-web-migration/specs/` — verifica que cada `Requirement:` que cita la ficha existe.
3. `git diff --stat` ≤ 400 líneas.
4. `python3 scripts/check_pr_size.py --base-ref main` (gate del repo).
5. Conventional commit message, sin co-authored-by IA.

## Limpieza

- Worktrees `claude-1073-gap-2`, `claude-1073-gap-3`, `claude-1073-gap-4` (creados en la tarea #1073) **NO se retiran ahora**: la issue #1226 sigue abierta (verificado a 2026-10-22 con `gh issue view 1226`). El owner la cerrará con el merge del PR final de la cadena #1073; solo entonces aplica `~/bin/merge-tools/agent-wt.sh remove claude-1073-gap-2`, `claude-1073-gap-3`, `claude-1073-gap-4` (en ese orden, usando `--discard` solo si no los ha empujado).

## Tasks (todo list)

- [x] Odd feature doc creada (`odd/tasks/807-fichas-capacidad-exp.md`).
- [ ] PR-1 — Lifecycle (CAP-001, 002, 003, 006, 007).
- [ ] PR-2 — Related-data I (CAP-008..012).
- [ ] PR-3 — Related-data II + cat (CAP-013, 014, 015, 023, 025).
- [ ] PR-4 — Destino E2E + access-control (CAP-040, 043..046).
- [ ] PR-5 — Runtime + integrations I (CAP-047..052).
- [ ] PR-6 — Integrations II (CAP-053..056).
- [ ] PR-7 — Retirement I (CAP-057..061).
- [ ] PR-8 — Retirement II + index + Closes #807 (CAP-062, 063).
- [ ] Limpieza de worktrees `claude-1073-gap-2/3/4` cuando #1224-#1226 estén mergeadas.

## Estado / log

- **L1** (input original del owner, 2026-10-22): «issue #807 de DysTelefonica/access2web-blueprint (BP), "[EXP] docs(exp): fichas de capacidad faltantes de Expedientes (37 de 63)"… Lee la issue entera primero… Copia exactamente el formato de las fichas ya existentes. Trazabilidad… Decisiones de producto abiertas (P-EXP-*) se citan, no se resuelven. Cadena de PRs ≤400 líneas… Subagentes gentle-ai-worker (MiniMax-M3)… Cadenas… empuja una rama por comando, máximo 2 ramas tuyas en CI a la vez.»
- **L2** (verificación previa, 2026-10-22): la issue indica 37 fichas faltantes (CAP-001, 002, 003, 006..015, 023, 025, 040, 043..063), partiendo de las 26 ya creadas en `docs/03-aplicaciones/expedientes/capabilities/`. Verificado contra `ls` del directorio. Las specs (OpenSpec) cubren las 63; el ledger normativo (`capability-ledger.md`) da los títulos canónicos por EXP-CAP. #1226 sigue OPEN, así que la limpieza queda diferida.
- **L3** (decisión de ruteo, 2026-10-22): se usa la ruta ya adoptada por las 26 existentes (`docs/03-aplicaciones/expedientes/capabilities/CAP-NNN-slug.md`) y se cita la ruta alternativa del plan en la sección de pendiente de producto. No se mueven los hubs D51..D54.
