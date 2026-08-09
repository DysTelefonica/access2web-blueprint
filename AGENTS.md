# access2web-blueprint — Agent Skills Index

When working on this project, load the relevant skill(s) BEFORE writing any code or docs.

> **For documentation discipline (tone, naming, format, templates, anti-patterns), see the global `AGENTS.md` `<!-- DOCS_DISCIPLINE -->` block. The `documentation-alan-style` skill applies automatically to all repos via the global config.**

## Cómo usar

1. Revise la columna «trigger» para localizar las skills aplicables a la tarea actual.
2. Cargue la skill leyendo el archivo `SKILL.md` en la ruta indicada (mediante la herramienta `read` o el symlink correspondiente).
3. Aplique todos los patrones y reglas de la skill cargada.
4. Cuando varias skills apliquen a la vez, pueden combinarse.

## Skills obligatorias

These skills are **mandatory** — load them before any work in their scope:

| Skill | Obligatorio para |
|---|---|
| **`branch-pr`** | Any commit, PR creation, or merge to `main`. Conventional commits, PR pequeño y reversible. En el merge, **sin** `--delete-branch`: la rama remota se conserva; lo que se limpia es el worktree local (`git worktree remove <path>`). |
| **`estado-planificacion-update`** | Cualquier ciclo de refactor de una app del blueprint, desde el estudio inicial (`access-vba-capability-docs`, walkthrough de forms, propuesta OpenSpec) hasta la entrega (cierre de PRs, archive del change). El ciclo se actualiza con un `estado-planificacion-NEW_YYYY-MM-DD.html` cada vez que se cierra una fase, un PR o un gap. La skill trae `scripts/render_estado.py` (scaffold desde `tasks.md` + JSON config), `assets/style-base.css` y `assets/template.html`. Mantener el archivo vivo es la single source of truth visual del avance para los jefes. Vive en `~/.config/opencode/skills/estado-planificacion-update/SKILL.md`. Author: ardelperal. |

## Skills cross-cutting (provienen de otros repos)

| Skill | Trigger | Origen |
|---|---|---|
| `telefonica-brand-design` | Cualquier UI Mistica (design tokens, brand, layout). | Gentleman-Programming/mistica |
| `frontend-design` | Diseño UI distintivo (no AI defaults: cream/serif/terracotta; near-black/acid-green; broadsheet/hairline). | Gentleman-Programming |
| `dysflow-usage` | Cualquier uso de dysflow MCP (analyze_form_ui, map_form_behavior, verify_form_bindings, etc.). | Gentleman-Programming |
| `dysflow-arnes` | Hard rules de dysflow (HR-1..HR-14) y anti-patterns (AP-1..AP-12). | Gentleman-Programming |
| `maintainer-prompt-drafter` | Estructurar prompts para el mantenedor de dysflow u otro tool externo. | Gentleman-Programming |
| `access-vba-tdd` | Tests TDD en Access/VBA. | Gentleman-Programming |
| `access-form-ui-builder` | Perceive → act → verify loop para forms Access. | Gentleman-Programming |
| `deterministic-quality-harness` | Wiring o auditoría de quality gates de plataforma (CI, lint, ratchet baseline, complexity, hexagonal layer gate, secret/dependency scan). Inspirada en `unclebob/swarm-forge` y validada por APAP_WEB. Definida en `.opencode/skills/deterministic-quality-harness/SKILL.md`, symlinkeada desde `~/.config/opencode/skills/`. Author: ardelperal. | este repo |

## Refuerzo

Si una IA escribe código o documentos en este repositorio sin cargar las skills pertinentes, trabaja a ciegas. Las skills son la single source of truth para «cómo se hace X en este proyecto». Cargue primero, escriba después.

### Ciclo de refactor de una app (estado-planificacion-update)

La skill `estado-planificacion-update` cubre el ciclo completo de vida de un refactor y exige mantener el HTML vivo en cada paso. Reglas mínimas:

1. **Estudio inicial** (cargar `access-vba-capability-docs`): tras el primer walkthrough de forms, abrir `estado-planificacion-NEW_<fecha>.html` con la tabla de apps en estudio, RAG 🟢 / 🟡 / 🔴 por app, y los gaps identificados.
2. **Propuesta OpenSpec** (`/sdd-new`): cuando arranca el change, añadir al HTML la sección "Próximos pasos" con la propuesta, los PRs iniciales y el chain strategy decidido.
3. **Durante la planificación** (`sdd-spec`, `sdd-design`, `sdd-tasks`): cada cierre de fase refresca el HTML con la sección correspondiente, los gaps cerrados y los que quedan abiertos, el forecast de carga de revisión y el plan de PRs.
4. **Ejecución** (`sdd-apply`): cada PR cerrado actualiza el Gantt, la columna RAG de la tabla de PRs y el motivo del cambio. La cadena visual stack → PRs → plan debe quedar coherente.
5. **Verificación y archivo** (`sdd-verify`, `sdd-archive`): el HTML final muestra todos los PRs en verde, los gaps cerrados y el cierre del ciclo. Se commitea junto al último PR.

El archivo `estado-planificacion-*.html` es la single source of truth visual del avance para los jefes. Cuando un agente entra al repo, debe leer el HTML más reciente antes de proponer cambios.

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)
