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

[← Back to README](../README.md) · [Next: AGENTS.md →](../AGENTS.md)
