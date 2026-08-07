# access2web-blueprint — Agent Skills Index

When working on this project, load the relevant skill(s) BEFORE writing any code or docs.

## How to Use

1. Check the trigger column to find skills that match your current task
2. Load the skill by reading the SKILL.md file at the listed path (via `read` tool or symlink)
3. Follow ALL patterns and rules from the loaded skill
4. Multiple skills can apply simultaneously
5. When in doubt, prefer the more specific skill (e.g. `codebase-guide` over `documentation-patterns` for maintainer docs)

## Mandatory skills

These skills are **mandatory** — load them before any work in their scope:

| Skill | Mandatory for |
|---|---|
| **`documentation-patterns`** | **Any** doc writing or refactor (epic, walkthrough, README, DOCS, CODEBASE-GUIDE, AGENTS, walkthrough JSONs, dependency maps). Determines format (tabla/lista/párrafo), tone (español técnico neutro, sin emojis decorativos), structure (7 secciones + 2 anexos + checklist), and single-source-of-truth rules. |
| **`branch-pr`** | Any commit, PR creation, or merge to `main`. Conventional commits, PR pequeño y reversible, `--squash --delete-branch` en merge. |

## Project-context skills (este repo)

| Skill | Trigger | Path |
|---|---|---|
| `a2web-documentation-patterns` | Escribir o revisar cualquier doc. Decidir formato (tabla/lista/párrafo), tono, estructura. | `C:\Proyectos\skills\skills\documentation-patterns\SKILL.md` |
| `a2web-docs-index` | Crear o refactorizar el `DOCS.md` raíz con Quick Navigation + Other docs. | `C:\Proyectos\skills\skills\docs-index\SKILL.md` |
| `a2web-codebase-guide` | Crear o refactorizar el `CODEBASE-GUIDE.md` con 90-second mental model + quick map inverso. | `C:\Proyectos\skills\skills\codebase-guide\SKILL.md` |

> **Las skills viven físicamente en `C:\Proyectos\skills\skills\` y están linkeadas en opencode.** Los paths arriba son absolutos. Si necesitás editar las skills, editá el original en `C:\Proyectos\skills\skills\`.

## Cross-cutting skills (vienen de otros repos)

| Skill | Trigger | Origen |
|---|---|---|
| `telefonica-brand-design` | Cualquier UI Mistica (design tokens, brand, layout). | Gentleman-Programming/mistica |
| `frontend-design` | Diseño UI distintivo (no AI defaults: cream/serif/terracotta; near-black/acid-green; broadsheet/hairline). | Gentleman-Programming |
| `dysflow-usage` | Cualquier uso de dysflow MCP (analyze_form_ui, map_form_behavior, verify_form_bindings, etc.). | Gentleman-Programming |
| `dysflow-arnes` | Hard rules de dysflow (HR-1..HR-14) y anti-patterns (AP-1..AP-12). | Gentleman-Programming |
| `maintainer-prompt-drafter` | Estructurar prompts para el mantenedor de dysflow u otro tool externo. | Gentleman-Programming |
| `access-vba-tdd` | Tests TDD en Access/VBA. | Gentleman-Programming |
| `access-form-ui-builder` | Perceive → act → verify loop para forms Access. | Gentleman-Programming |

## Reinforcement

Si una IA está escribiendo código o docs en este repo sin cargar las skills relevantes, está trabajando a ciegas. Las skills son el single source of truth para "cómo se hace X en este proyecto". Cargá primero, escribí después. Si una IA escribe una doc sin haber cargado `documentation-patterns`, **la doc será rechazada en review** — el formato no cumplirá con el patrón del repo.
