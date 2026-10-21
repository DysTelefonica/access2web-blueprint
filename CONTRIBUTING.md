# access2web-blueprint — Guía de contribución

## Bienvenida

Consulte esta guía antes de abrir un issue o un PR. Mantenga cada cambio centrado en una única idea revisable.

## Cómo contribuir

1. Abra un issue que describa el problema, el alcance y el resultado esperado.
2. Cree una branch específica a partir de `main`.
3. Aplique el cambio y documente las decisiones relevantes.
4. Verifique el contenido con los checks definidos por el repositorio (ver la siguiente sección «Antes de abrir una PR»).
5. Abra un PR pequeño con una descripción verificable.
6. Responda a la revisión y mantenga actualizada la branch.

## Antes de abrir una PR

La siguiente tabla recoge los checks que `ci.yml` ejecuta sobre la PR, qué verifica cada uno, cómo aprobarlo localmente y el comando que lo replica. La tabla cubre los gates `pr-gate` (revisión barata, primer corte), `quality` (tests, typecheck, cobertura, ratchets de calidad), `security` (gitleaks, pip-audit, trivy-config) y `codeql` (análisis estático). El agregador `required` exige los cuatro jobs verdes sobre la base actual (`ci.yml:497-529`).

| Check | Qué comprueba | Cómo pasarlo | Comando local |
|---|---|---|---|
| Nombre de rama (pr-gate) | La rama sigue el patrón gentle-ai `^(feat\|fix\|chore\|docs\|style\|refactor\|perf\|test\|build\|ci\|revert)/[a-z0-9._-]+$` (sin número de issue); lo enforza el ruleset del host | Nombrarla `<tipo>/<slug>` con kebab-case; las cadenas encadenadas añaden sufijo | — (lo enforza el ruleset, no un script local) |
| Issue-spec (pr-gate) | El cuerpo lleva `Closes`/`Fixes`/`Resolves`/`Refs #N` apuntando a una issue con `status:approved` y un `type:*` (Refs #855) | Poner `Closes #N` en la primera línea del cuerpo con la issue aprobada | `python scripts/check_issue_spec.py local --body-file <body>` |
| Tamaño de la PR (pr-gate) | `additions + deletions` de la API de la PR ≤ 400; la etiqueta `size:exception` la pone el mantenedor como override solo cuando partir la PR dejaría `main` roto y tras responder las 4 preguntas de `AGENTS.md` §F.7 (Refs #854) | Partir por unidad de trabajo o encadenar PRs (orden de escape); `size:exception` es la última opción | `python .agents/skills/ci-pattern/assets/pr-size/check_pr_size.py --base-ref origin/main` |
| Esquema de walkthroughs (pr-gate) | Los JSON `docs/03-aplicaciones/*/walkthrough-*.json` cumplen el contrato v2.1 (`check_walkthrough_schema.py --strict`) | Regenerar con dysflow + codegraph-vba; nunca editar a mano | `python scripts/check_walkthrough_schema.py --strict` |
| Lint (pr-gate) | `ruff check` sin findings | `ruff check .` antes de commit | `ruff check --config app/pyproject.toml .` |
| Formato (quality) | `ruff format` sin cambios en el árbol | `ruff format .` antes de commit | `ruff format --check --config app/pyproject.toml .` |
| Typecheck (quality) | `mypy --explicit-package-bases app/` verde | `mypy app/` antes de push | `mypy --explicit-package-bases app/` |
| Workflow gate (quality) | Ningún job de `.github/workflows/*.yml` usa `|| true`, `continue-on-error` ni `pull_request_target`; todos declaran `timeout-minutes` | Diseñar el job con Hard Rule 1, 13 y 15 | `python scripts/check_workflows.py` |
| Tests y cobertura (quality) | `pytest` verde con `--cov-fail-under` ≥ suelo | Correr la suite antes del push | `pytest -c app/pyproject.toml --rootdir=app -m "not integration" --cov` |
| Quality report (quality) | `quality_report.py` verde en orden `layers → complexity → DRY` | `python scripts/quality_report.py` antes del push | `python scripts/quality_report.py` |
| Security (gitleaks, pip-audit, trivy-config) | Sin secretos en el diff, sin CVEs en el grafo de dependencias, sin misconfiguraciones trivy | Auditar el diff localmente antes del push; abrir issue de incidente si aparece algo | — (solo CI) |
| CodeQL | Análisis estático sin alertas de severidad alta | — | — (solo CI) |

> Las preguntas obligatorias antes de pedir `size:exception` viven en `AGENTS.md` §F.7. «Son sobre todo tests» o «es un solo comportamiento» no bastan.

`make verify` ejecuta los gates locales en el mismo orden que el CI (Hard Rule 13): format, lint, typecheck, tests, quality report, check-workflows. Es la única forma de declarar verde una PR sin haber pasado antes por el runner.

## Propagaciones de flota

El catálogo de skills de la flota llega a este repositorio por una vía propia, y es la única entrega que no nace de una issue de trabajo: la ampara la issue paraguas #846, la misma que `fleet/registry.json` de `team-skills` declara como `propagation_issue` de este consumer.

- **Quién la dispara.** Un mantenedor (u orquestador) ejecuta, desde el clon del catálogo (`team-skills`), `scripts/propagate-team-skills.ps1 -ConsumerId DysTelefonica/access2web-blueprint`, primero en ensayo (sin `-Apply`) y luego con `-Apply`. El ensayo no modifica el consumer (`HEAD`, `status` y `reflog` quedan idénticos).
- **Qué emite.** Dos PR con Chain Context completo: el tramo 1 (`skill-fleet/access2web-blueprint`, `AGENTS.md`, `.team-skills.yaml` y sello de gobierno) y el tramo 2 (`skill-fleet/access2web-blueprint-vendored`, el árbol `.agents/skills/`). Cada PR lleva `Refs #846` y nunca `Closes`.
- **Qué no se hace.** No se edita a mano el bloque del slice de `AGENTS.md` ni se parte a mano el trabajo del propagador; no se abre un PR de propagación escrito a mano. La skill que sale del catálogo, o que pasa a vivir en `.agents/skills/`, la retira el propagador, que también retira la copia legacy de `skills/<nombre>/`.
- **Tamaño.** El tramo vendorizado puede superar las 400 líneas. La etiqueta `size:exception` la pone el mantenedor en el PR; el propagador no se la aplica.
- **Lo que queda local.** Las skills de `skills/` que no figuran en `include_skills` de `.team-skills.yaml` son de este repositorio y el propagador no las toca.

## Conventional Commits

| Tipo | Ejemplo | Uso |
|---|---|---|
| `docs` | `docs(app): documentar reglas de autorización` | Cambios de documentación |
| `fix` | `fix(app): corregir referencia de formulario` | Correcciones de comportamiento o contenido |
| `feat` | `feat(app): añadir capacidad de migración` | Nueva capacidad documentada |
| `chore` | `chore(repo): actualizar configuración de herramientas` | Mantenimiento sin cambio funcional |
| `refactor` | `refactor(app): reorganizar la especificación` | Reorganización sin cambio de alcance |

## Label system

| Label | Color | Use for |
|---|---|---|
| `🐛 bug` | `#d73a4a` | Defectos verificables |
| `📚 documentation` | `#0075ca` | Cambios documentales |
| `✨ enhancement` | `#a2eeef` | Mejoras de alcance existente |
| `🔍 research` | `#7057ff` | Investigación y evidencia |
| `⚙️ chore` | `#e4e669` | Mantenimiento técnico |

## Convención multi-app

Este repositorio gobierna la migración a web de ocho aplicaciones legadas. Cada issue y cada commit declara a qué aplicación pertenece, o se marca como transversal. La convención permite que varias IAs trabajen en paralelo sin pisarse y habilita el filtrado por aplicación.

### Etiquetas por aplicación

Aplique la etiqueta `app/<slug>` a los issues que pertenezcan a una sola aplicación. El slug coincide con el directorio canónico `docs/03-aplicaciones/<slug>/`.

| Etiqueta | Aplicación |
|---|---|
| `app/condor` | Condor |
| `app/hps` | HPS |
| `app/hps-solicitudes` | HPS Solicitudes |
| `app/brass` | Brass |
| `app/gestion-riesgos` | Gestión de Riesgos |
| `app/no-conformidades` | No Conformidades |
| `app/lanzadera` | Lanzadera |
| `app/expedientes` | Expedientes |

### Etiqueta transversal

Aplique `cross-cutting` a los issues que afecten a varias aplicaciones simultáneamente. Ejemplos: el patrón de secret manager (D93), la migración de booleanos `Text(2)` a `BOOLEAN` (D102), la regla de gitignore para `.accdb` (D92). El label `cross-cutting` coexiste con cualquier `app/<slug>` cuando una sola app origina la iniciativa.

### Prefijo en el título del issue

Anteponga al título del issue el código de la aplicación entre corchetes, en mayúsculas. Use `XCUT` para transversal.

```text
[EXP] Definir catálogos web
[LANZ] Scaffolding del MVP
[XCUT] Estandarizar booleanos a BOOLEAN
```

### Scope en commits

Cierre el placeholder `(app)` de los conventional commits con el slug de la aplicación, o con `platform` para transversal.

```text
docs(expedientes): documentar matriz de migración
feat(lanzadera): portar formulario de login
fix(platform): migrar Text(2) a BOOLEAN
```

### Codificación de prefijos

| Código | Aplicación |
|---|---|
| `[COND]` | Condor |
| `[HPS]` | HPS |
| `[HPSS]` | HPS Solicitudes |
| `[BRASS]` | Brass |
| `[GR]` | Gestión de Riesgos |
| `[NC]` | No Conformidades |
| `[LANZ]` | Lanzadera |
| `[EXP]` | Expedientes |
| `[XCUT]` | Transversal |

## Reglas específicas de este repositorio

- Mantenga los PRs pequeños y centrados en una única idea.
- Integre con `--squash`. **No borre la rama remota.** Limpie el worktree local
  con `git worktree remove <ruta>`.
- No añada `Co-Authored-By` ni atribución de IA a los commits.
- Use el prefijo `docs(app):` para las épicas.
- Mantenga los walkthrough JSONs en commits separados.

### Por qué se conserva la rama remota

La integración es `--squash`, de modo que `main` recibe un único commit por PR.
Los commits por unidad de trabajo —los que separan una entrega en pasos
revisables— existen únicamente en la rama.

Borrar la rama los deja sin referencia y el recolector de basura acaba
eliminándolos. Se perdería justo la granularidad que el trabajo por unidades
pretende crear: `main` conserva el qué, y la rama, el cómo se llegó.

El coste es una lista de ramas larga. La convención de nombre
`<tipo>/<slug>` (patrón gentle-ai, sin número de issue; lo enforza el ruleset)
la mantiene navegable; la trazabilidad con la issue vive en el cuerpo del PR,
con `Closes #N`.

[← Back to README](README.md) · [Next: CHANGELOG.md →](CHANGELOG.md)
