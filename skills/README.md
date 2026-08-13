# Skills del repositorio

Las skills que este proyecto necesita viven aquí, versionadas junto al código.
Quien clona el repositorio las tiene.

## Contrato

Sólo pueden asumirse instalados dos harness externos:

| Harness | Qué aporta |
|---|---|
| `gentle-ai` | 26 skills, entre ellas `branch-pr`, `chained-pr`, `work-unit-commits`, `issue-creation` y la suite `sdd-*` |
| `engram` | Memoria persistente entre sesiones |

Todo lo demás vive en este directorio. **Una convención obligatoria del proyecto
no puede depender de una skill instalada en el ordenador de una persona.** Si es
obligatoria, o está aquí, o está en `CONTRIBUTING.md` o en `docs/`.

## Instalación

Los agentes no leen `skills/` directamente: cada uno busca en su propio
directorio. El script copia lo de aquí al que corresponda.

```bash
# macOS / Linux
./scripts/install-skills.sh

# Windows (PowerShell)
powershell -ExecutionPolicy Bypass -File scripts/install-skills.ps1
```

Detecta los agentes presentes e instala en el destino de cada uno:

| Agente | Destino |
|---|---|
| Claude Code | `.claude/skills/` |
| OpenCode | `.opencode/skills/` |
| Codex | `.codex/skills/` |

Los destinos son locales al repositorio y están en `.gitignore`: el original es
este directorio, y las copias se regeneran. Vuelva a ejecutar el script tras un
`git pull` que toque `skills/`.

## Qué entra aquí y qué no

| Entra | No entra |
|---|---|
| Convenciones de este proyecto | Skills que ya distribuye `gentle-ai` |
| Herramientas con sus assets (plantillas, scripts) | Organización del disco de un desarrollador |
| Disciplina documental del repositorio | Skills de dominio Access, salvo para leer la entrada de migración |

Las skills de dominio Access (`dysflow-*`, `access-*`, `vba-*`) quedan fuera a
propósito. El alcance de este repositorio está descrito en `AGENTS.md`: ese
material es entrada de migración en sólo lectura, nunca producto.

## Añadir una skill

1. Cree `skills/<nombre>/SKILL.md` siguiendo el contrato LLM-first: frontmatter,
   Activation Contract, Hard Rules, Decision Gates, Execution Steps, Output
   Contract, References.
2. Los assets acompañantes van en `skills/<nombre>/assets/` o
   `skills/<nombre>/references/`.
3. Regístrela en la tabla de `AGENTS.md` con su trigger.
4. Un PR por skill. La mayoría no cabe en el presupuesto de 400 líneas junto a
   otra, de modo que se entregan encadenadas.
