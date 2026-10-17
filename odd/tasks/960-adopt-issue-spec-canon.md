# 960-adopt-issue-spec-canon

- Repository: DysTelefonica/access2web-blueprint · Tracking issue: #960 (feat(issue-spec): adopt the canonical byte-by-byte gate + --required-sections for the 6 issue-form sections)
- Locator: `odd/tasks/960-adopt-issue-spec-canon.md` · Engram mirror topic `odd/960-adopt-issue-spec-canon/tasks`
- Skill under change: gate canónico `scripts/check_issue_spec.py` (DysTelefonica/team-skills, asset `assets/issue-spec/`, blobs `75518d0` y tests `807cb123`)

## Specs

S1. `scripts/check_issue_spec.py` se sustituye byte a byte por el canon de `DysTelefonica/team-skills` (blob `75518d04ef73c717ccff6b848b5db06fa67274e5`, PR team-skills #646 mergeado). La copia aplica las 7 reglas del gate: 4 sobre el cuerpo del PR (referencia, resolución, no mezcla Closes+Refs, repositorio) + 3 sobre la issue consultada por la API (issue, etiquetas, fail-closed). El SHA256 del fichero en BP tras la sustitución coincide con el del canon.
S2. `.github/workflows/ci.yml` invoca el gate con `--required-sections scripts/required-sections.json` en el job `pr-event` del workflow `ci`. La bandera es opt-in en el canon; sin ella, el comportamiento de las reglas 1..6 no cambia.
S3. `scripts/required-sections.json` lista las 6 secciones obligatorias del contrato de issue, copiadas verbatim del atributo `label` de cada textarea en `.github/ISSUE_TEMPLATE/issue-canonical.yml`: "Problema y contexto", "Evidencia verificable", "Alcance y no objetivos", "Criterios de aceptación", "Plan de validación", "Dependencias y riesgos". El formato es `{"required_sections": [...]}`; el canon compara títulos Markdown H1..H6 contra estos strings, exacto y case-sensitive.
S4. `.github/ci-pattern-adoption.json` fase 0 actualiza la entrada de `scripts/check_issue_spec.py` (descripción: canon DysTelefonica/team-skills blob 75518d0, regla 7 opt-in --required-sections para las 6 secciones del contrato de issue; Refs #960) y añade `scripts/required-sections.json` como artefacto adoptado.
S5. `app/pyproject.toml` mantiene el script canónico en `extend-exclude` de ruff format (líneas largas del canon >100; mismo boundary que `skills/` y `.agents/skills/`, precedente `ardelperal/APAP_WEB#839`). La exclusión ya quedó en PR-1 (commit `b2df561`).
S6. El smoke stub de 5 tests en `tests/test_check_issue_spec.py` introducido por PR-1 sigue verde: la API expuesta por el canon (`references`, `validate_body`, `main`) no cambió. La suite canónica de 35 tests entra con PR-3 (chain 3/3, Closes #960).
S7. PR ≤ 400 líneas. Si no cabe, cadena con `chain:partial` + `Refs #960` en intermedias y `Closes #960` en la punta. Nunca `size:exception`. Sin rebase ni force-push. Una PR en CI a la vez.
S8. Commits convencionales en inglés, sin atribución de IA. La sustitución byte a byte del script, la creación del JSON de secciones, el cableado del workflow y la actualización del contrato de adopción van en commits separados, cada uno con su `Why` (Refs #960).

## Tasks

- T1 (S1) feat(issue-spec): adopt canonical 75518d0 byte-by-byte — sustituir `scripts/check_issue_spec.py` por el blob `75518d0` del canon. PR-1 (#961, ya en verde, esperando merge del usuario). Commit: `22ea523`.
- T2 (S2, S3) feat(issue-spec): wire --required-sections into pr-event mode — crear `scripts/required-sections.json` con las 6 secciones del formulario de issue y añadir `--required-sections` a la invocación del job `issue spec` en `.github/workflows/ci.yml:140`. PR-2 (este PR, chain 2/3, Refs #960). Commit: pending.
- T3 (S4) chore(ci-pattern): update phase-0 inventory with the 75518d0 canon — actualizar `.github/ci-pattern-adoption.json` fase 0: descripción del asset en `scripts/check_issue_spec.py` y nueva entrada para `scripts/required-sections.json`. PR-2 (mismo que T2). Commit: pending.
- T4 (S7) push de PR-2 (`chore/960-upgrade-75518d0` basada en `chore/960-adopt-issue-spec-canon`), esperar CI en verde, escribir «PR <n> verde, espero merge» y parar.

Estado: T1 cerrada en PR-1. T2-T4 en cola para este turno.

## Log

L1. Petición original (verbatim, del retomar tras /new por corte de cuota de MiniMax): «Tarea: DysTelefonica/access2web-blueprint #961 (adopción de issue-spec, Refs #960). Lleva una copia vieja de check_issue_spec.py (blob c2577fa). Canon issue-spec: team-skills origin/main personal/ardelperal/ci-pattern/assets/issue-spec/check_issue_spec.py (blob 75518d0; compruébalo con git hash-object) y sus tests, copiados byte a byte. El workflow pasa --required-sections con las secciones del formulario de issue del repo. Completa la cadena hasta Closes #960 con PRs de ≤400.»

L2. Verificación del canon al retomar: `git hash-object` sobre `team-skills:personal/ardelperal/ci-pattern/assets/issue-spec/check_issue_spec.py` da `75518d04ef73c717ccff6b848b5db06fa67274e5` (la rama `main` local de team-skills estaba 2 commits ahead de `origin/main`; la verificación se hace contra `origin/main`, no contra el working tree). Tests en `807cb12344841139df0173bb04453b19267fe97a` (336 líneas, suite de 35 casos).

L3. PR-1 (#961) ya estaba en verde en CI sobre la copia vieja (blob `c2577fa`, 215 inserciones / 166 deletions / 381 líneas / presupuesto 400). Cadena previa (PR-1 → #962 con 27-test suite → PR-3) abortada por reorganización tras team-skills#644/#646; ahora se reorienta: PR-1 mantiene c2577fa como entrada a la cadena, PR-2 trae el upgrade al canon con regla 7, PR-3 trae la suite canónica + preflight + Closes #960.

L4. Decisión de scope: el smoke stub de 5 tests de PR-1 sigue siendo válido (la API expuesta por el canon no cambió entre 75518d0 y la API que ya ejercita el stub); la suite canónica byte a byte de 35 tests entra con PR-3. La doc ODD del PR-1 (commit `2147cf5`) se había movido a PR-2 (commit `4b72103`); al reorganizarse, este PR-2 la recrea (35 líneas) porque el primer commit que cambia código en la nueva cadena es este, no PR-1 (que ya está pusheado y al margen de la nueva cadena).
