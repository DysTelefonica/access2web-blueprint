[← Back to calidad-de-codigo-y-ci.md](calidad-de-codigo-y-ci.md)

# access2web-blueprint — Agent Setup

Configure el agente elegido con las skills adecuadas y verifique el acceso a las herramientas antes de trabajar.

## Quick Navigation

| Sección | Contenido |
|---|---|
| [Agentes soportados](#agentes-soportados) | Agentes compatibles y notas de configuración |
| [Cómo cargar las skills](#cómo-cargar-las-skills) | Rutas absolutas de las skills |
| [Smoke test por agente](#smoke-test-por-agente) | Comandos de verificación MCP |
| [Troubleshooting](#troubleshooting) | Errores frecuentes y correcciones |

## Agentes soportados

| Agente | Setup | Notas |
|---|---|---|
| Claude Code | Cargue `AGENTS.md` y las skills indicadas. | Use el symlink de configuración disponible. |
| OpenCode | Cargue `AGENTS.md` y las skills indicadas. | Use las herramientas MCP configuradas. |
| Codex | Cargue `AGENTS.md` y las skills indicadas. | Mantenga el contexto del worktree. |
| Gemini | Cargue `AGENTS.md` y las skills indicadas. | Verifique el acceso al proyecto antes de empezar. |

## Cómo cargar las skills

Cargue primero la skill canónica y después las skills específicas de la tarea.

1. `C:\Proyectos\skills\skills\documentation-alan-style\SKILL.md`
2. `C:\Proyectos\skills\skills\dysflow-usage\SKILL.md`
3. `C:\Proyectos\skills\skills\dysflow-arnes\SKILL.md`
4. `C:\Proyectos\skills\skills\codegraph-usage\SKILL.md`
5. `C:\Proyectos\skills\skills\access-vba-tdd\SKILL.md`

## Smoke test por agente

```bash
# Verifique el contrato de Dysflow.
dysflow.get_capabilities({})
# Consulte el grafo del proyecto.
codegraph.codegraph_explore({ query: "LeeConfiguracionLocal", projectPath: "C:/00repos/codigo/access2web-blueprint" })
```

## Troubleshooting

1. Si una skill no aparece, compruebe la ruta absoluta y el symlink correspondiente.
2. Si Dysflow no responde, compruebe `.dysflow/project.json` y el proceso del adaptador.
3. Si CodeGraph no devuelve resultados, compruebe que el proyecto tenga un índice actualizado.
4. Si el smoke test falla, revise primero el worktree y las rutas configuradas.

## Core invariants

- **Cargue primero la skill canónica y después las específicas**: el orden de carga de skills debe empezar por las que aplican al repo entero (`documentation-alan-style`, `architecture-guardrails`, `branch-pr`) y luego las específicas del trabajo (por ejemplo `dysflow-usage` para analizar binarios legacy).
- **Smoke test antes de empezar a trabajar**: ejecute los comandos de la sección §Smoke test por agente antes de abrir un issue o un PR. Si el smoke test falla, revise primero el worktree y las rutas configuradas; no escriba código hasta que pase.
- **Skills de dominio Access sólo aplican al material de entrada**: las skills `dysflow-*`, `access-*`, `vba-*` están pensadas para `data/`, `inputs/`, `.dysflow/` y los `.accdb` de la raíz. NO gobiernan `app/`, `docs/`, `scripts/`, `.github/` ni `openspec/` — ahí manda `CONTRIBUTING.md` y `AGENTS.md`.
- **Rutas de skills asumen instalación via `scripts/install-skills.sh`** (o `.ps1` en Windows). Las rutas absolutas de la sección §Cómo cargar las skills reflejan el setup actual del autor; la migración a `skills/` interno del repo es issue #167.

## Contributor checklist

- [ ] El agente (Claude, OpenCode, Codex, Gemini) carga `AGENTS.md` antes del primer comando y respeta el frontmatter YAML (`globs: *` + `alwaysApply: true`).
- [ ] Si se añade una nueva skill al catálogo, se documenta en `AGENTS.md` (tabla `Project-context skills` o `Cross-cutting skills`) y en `skills/README.md`.
- [ ] Si se cambia una ruta absoluta de skill, se actualiza también `scripts/install-skills.sh` y `.ps1` para que la instalación reproduzca el path.
- [ ] El smoke test por agente corre verde desde el worktree activo antes de mergear el PR.
- [ ] El doc mantiene castellano peninsular formal con usted y no introduce anglicismos fuera de la lista permitida en `skills/documentation-alan-style/SKILL.md` §4.

## Navigation

Previous: [calidad-de-codigo-y-ci.md](calidad-de-codigo-y-ci.md) | Next: [CODEBASE-GUIDE](../CODEBASE-GUIDE.md)

---

[← Back to README](../README.md) · [Next: AGENTS.md →](../AGENTS.md)
