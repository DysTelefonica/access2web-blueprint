# access2web-blueprint — Agent Skills Index

When working on this project, load the relevant skill(s) BEFORE writing any code or docs.

> **Scope reminder**: este repo es research + planning de la migración de 8 apps legacy → web. NO es código de producto. Las skills aplican tanto aquí como en los futuros repos de cada app.

## How to Use

1. Check the trigger column to find skills that match your current task
2. Load the skill by reading the SKILL.md file at the listed path (via `read` tool or symlink)
3. Follow ALL patterns and rules from the loaded skill
4. Multiple skills can apply simultaneously
5. When in doubt, prefer the more specific skill (e.g. `codebase-guide` over `documentation-patterns` for maintainer docs)

## Project-context skills (este repo)

| Skill | Trigger | Path |
|-------|---------|------|
| `documentation-patterns` | Escribir o revisar cualquier doc (epic, walkthrough, README, DOCS, CODEBASE-GUIDE). Decidir formato (tabla/lista/párrafo). | [`docs/_skills/documentation-patterns/SKILL.md`](docs/_skills/documentation-patterns/SKILL.md) |
| `docs-index` | Crear o refactorizar el `DOCS.md` raíz con Quick Navigation + Other docs. | [`docs/_skills/docs-index/SKILL.md`](docs/_skills/docs-index/SKILL.md) |
| `codebase-guide` | Crear o refactorizar el `CODEBASE-GUIDE.md` con 90-second mental model + quick map. | [`docs/_skills/codebase-guide/SKILL.md`](docs/_skills/codebase-guide/SKILL.md) |

> **Las skills viven físicamente en `C:\Proyectos\skills\skills\` y están linkeadas en opencode.** Los paths arriba son relativos al repo y funcionan vía `read` tool. Si necesitás editar las skills, editá el original en `C:\Proyectos\skills\skills\`.

## Cross-cutting skills (vienen de otros repos)

| Skill | Trigger | Origen |
|-------|---------|--------|
| `telefonica-brand-design` | Cualquier UI Mistica (design tokens, brand, layout). | Gentleman-Programming/mistica |
| `frontend-design` | Diseño UI distintivo (no AI defaults: cream/serif/terracotta; near-black/acid-green; broadsheet/hairline). | Gentleman-Programming |
| `access-vba-tdd` | Tests TDD en Access/VBA. | Gentleman-Programming |
| `access-form-ui-builder` | Perceive → act → verify loop para forms Access. | Gentleman-Programming |
| `dysflow-usage` | Cualquier uso de dysflow MCP (analyze_form_ui, map_form_behavior, etc.). | Gentleman-Programming |
| `dysflow-arnes` | Hard rules de dysflow. | Gentleman-Programming |
| `maintainer-prompt-drafter` | Estructurar prompts para el mantenedor de dysflow u otro tool externo. | Gentleman-Programming |
| `branch-pr` | Crear PR con issue-first workflow + conventional commits. | Gentleman-Programming |

## Skills workflow per app

| Si estás trabajando en... | Carga estas skills |
|---|---|
| Lanzadera (la madre, próxima a refinar) | `documentation-patterns` + `telefonica-brand-design` + `frontend-design` + `branch-pr` |
| Walkthrough de una app (G1..G5) | `documentation-patterns` + `dysflow-usage` + `dysflow-arnes` + `access-form-ui-builder` |
| Diseño de mockups UI | `telefonica-brand-design` + `frontend-design` |
| Filing de issues a dysflow (rondas 1..N) | `maintainer-prompt-drafter` + `documentation-patterns` |
| Épica de una app nueva | `documentation-patterns` + `codebase-guide` + `branch-pr` |
| Refactor de epic.md existente | `documentation-patterns` (aplicar formato consistente) |

## Reinforcement

Si una IA está escribiendo código o docs en este repo sin cargar las skills relevantes, está trabajando a ciegas. Las skills son el single source of truth para "cómo se hace X en este proyecto". Cargá primero, escribí después.