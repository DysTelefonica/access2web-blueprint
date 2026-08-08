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
| **`branch-pr`** | Any commit, PR creation, or merge to `main`. Conventional commits, PR pequeño y reversible, `--squash --delete-branch` en merge. |

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

## Refuerzo

Si una IA escribe código o documentos en este repositorio sin cargar las skills pertinentes, trabaja a ciegas. Las skills son la single source of truth para «cómo se hace X en este proyecto». Cargue primero, escriba después.

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)
