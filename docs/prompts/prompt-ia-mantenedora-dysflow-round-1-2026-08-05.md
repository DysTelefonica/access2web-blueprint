Eres la IA mantenedora de dysflow MCP. Repo: `<repo path>`. Branch: `<fix-rama-sugerida>`. Versión actual observada por el consumer: `2.35.3`. TDD estricto, conventional commits, no romper consumers en fleet.

## Contexto del round

Round 1 desde el consumer `access2web-blueprint` (Microsoft Access/VBA → web). El consumer tenía dos fricciones operativas concretas al usar dysflow MCP por primera vez:

1. La skill `dysflow-usage` (y el arnés `dysflow-arnes`) no aparecen como MUST-LOAD cuando un agente toca artefactos dysflow (`*.accdb`, `*.form.txt`, `.dysflow/project.json`, `tests/*.json`).
2. Los envelopes de error de dysflow **no sugieren** al agente cargar la skill, así que el agente improvisa diagnósticos estáticos en lugar de leer el runtime.

Resultado: el consumer inventó un diagnóstico falso (D89 sobre `list_objects` fallido) y casi rompe un commit, cuando la causa real era `FRONTEND_TARGET_AMBIGUOUS` por cuatro worktrees hermanos compartiendo `projectId`. El fix del consumer (migración de config + resolve_project con recoveryToken + register_worktree) funcionó **solo después** de cargar `dysflow-usage` y `dysflow-arnes` por insistencia del user.

Rounds previos: ninguno desde este consumer.

## Lo que YA funciona (NO tocar)

Validado por el consumer en este round con `adapterVersion: "2.35.3"`:

- `get_capabilities({})` retorna `adapterVersion`, `writeExecutionPolicy`, `effectiveDryRunDefault` por tool, `humanCompilePending`, `toolsVisible` (94), `projectConfig`, `documentationBundle`. Cumple el contrato de HR-0 del arnés.
- `resolve_project({cwd})` detecta ambigüedad multi-worktree y emite `availableProjects[]` + `recoveryToken` + `recoveryInstruction`. El consumer pudo resolver tras 1 retry con el trio (`projectId` + `projectChoiceReason: "user_selected_after_ambiguous_project"` + `recoveryToken`).
- `register_worktree({cwd})` pre-calienta el cache y devuelve `projectConfig.context` coherente. Tras invocar esto, llamadas posteriores con `cwd` funcionan sin ambigüedad.
- `migrate_project_config({cwd, apply:true})` migra `allowWrites` top-level → `capabilities.allowWrites` (T18 caps-block) de forma atómica. Devuelve `outcome: "ok"`, `applied: true`, `remediation[]` con el diff explicado.
- `list_objects({cwd, outputMode:"summary"})` retorna `{itemCount: 251, keys:["classes","documentModules","forms","modules","reports"]}` en el consumer. Read-only, no necesita `apply`.
- `writeExecutionPolicy: "safe-by-default"` y `effectiveDryRunDefault` consistente con el arnés HR-8.
- El discriminador `schemaVersion: "dysflow.result/v1"` aparece en TODAS las respuestas (verificado en `get_capabilities`, `resolve_project`, `register_worktree`, `migrate_project_config`, `list_objects`).
- El bloque `discoveredProjects[]` con `active: true|false` por worktree se respeta en la resolución.

NO reintroducir la tool `compile_vba` (HR-1 cross-project, decisión intencional).
NO cambiar `humanCompilePending` semantics (HR-1).
NO relajar `FRONTEND_TARGET_AMBIGUOUS` para que auto-resuelva — la elección humana es correcta.

## Lo que falta en este round

### Bug A: la skill `dysflow-usage` no es descubrible como MUST-LOAD al tocar artefactos dysflow

#### Síntoma verificado

Consumer recibe el bloque `<available_skills>` al inicio de cada sesión. Cuando el cwd contiene `.dysflow/project.json` o cualquier `*.accdb`, las skills `dysflow-usage` y `dysflow-arnes` aparecen listadas pero el agente no las carga proactivamente. Resultado observable:

- El agente abre `.dysflow/project.json` con `Read`, lee `backends.json`, inspecciona `getdb` en código VBA estático y construye una "diagnosis" del fallo sin haber llamado nunca a `get_capabilities`.
- El agente inventa D89 (`Diagnóstico del fallo de list_objects de Dysflow`) cuando el problema real era `FRONTEND_TARGET_AMBIGUOUS` por `projectId` duplicado en 4 worktrees.
- El commit de evidencia quedó con 488+ inserciones en `docs/03-aplicaciones/no-conformidades/` que incluyen un diagnóstico falso, detectable porque la skill existe y se carga después.

#### Evidencia de repro

```text
CWD: C:\00repos\codigo\00_NO_CONFORMIDADES\00_main
Archivos presentes: .dysflow/project.json, NoConformidades.accdb, NoConformidades_Datos.accdb

Paso 1 (lo que hizo el agente — INCORRECTO):
  Read .dysflow/project.json
  Read backends.json
  Inspeccionar Variables Globales.bas para entender getdb()
  Inventar 5 hipótesis estáticas sin contactar el runtime
  
Paso 2 (lo que debió hacer — CORRECTO):
  Cargar skill dysflow-usage
  Cargar skill dysflow-arnes
  get_capabilities({})
  resolve_project({cwd})
  Trabajar sobre el runtime, no sobre el código estático
```

Métricas del consumer `access2web-blueprint` (este round):

- 1 commit erróneo creado con D89 falso.
- 5 archivos de discovery con texto que se tendrá que reescribir (security-rules.md § D89, migration-matrix.md § D89).
- Tiempo perdido estimado: ~30 minutos antes de que el user corrigiera con "lee la skill de como se usa".

#### Diagnóstico preliminar

Probablemente:

1. El bloque en `AGENTS.md` / `.atl/AGENTS.md` que referencia dysflow (visto como `<!-- user-supplement:dysflow:pointer -->`) NO incluye una directiva MUST-LOAD para `dysflow-usage` cuando hay `.dysflow/*` en el cwd.
2. La skill `dysflow-usage` en su frontmatter `description` menciona "dysflow MCP tool" pero no "LOAD FIRST before any dysflow diagnosis or call".
3. La skill `dysflow-arnes` (arnés) está listada pero su descripción no enfatiza que es el primer archivo a leer si vas a tocar dysflow.
4. No hay un test de integración que verifique que `<available_skills>` muestre la skill `dysflow-usage` con un trigger que matche `cwd contains .dysflow/`.

Confirmar o descartar con el consumer si este patrón se repite cross-fleet (11+ proyectos consumer; varios con código VBA español/portugués).

#### Riesgo

- Cross-fleet: 11+ proyectos consumer (HPS, gestion_riesgos, no_conformidades, condor, cadete, brass, expedi entes, etc.) van a repetir el mismo antipatrón cada vez que un agente nuevo entre al proyecto.
- Alert fatigue / pérdida de confianza: agentes que inventan diagnósticos estáticos ensucian commits con afirmaciones falsas que el user tiene que revisar manualmente.
- Documentación contaminada: el consumer commiteó docs de discovery que ahora contienen afirmaciones sobre Dysflow que no se pueden verificar (porque se hicieron sin contactar el runtime).

#### Tests RED sugeridos

Test 1 — la skill `dysflow-usage` debe aparecer con un trigger fuerte:

```ts
it('lists dysflow-usage as MUST-LOAD when .dysflow/project.json is in cwd', async () => {
  const skills = await client.listAvailableSkills({ cwd: '<test-repo-with-.dysflow>' });
  const dysflowUsage = skills.find(s => s.name === 'dysflow-usage');
  expect(dysflowUsage).toBeDefined();
  expect(dysflowUsage.priority).toBe('must-load');
  expect(dysflowUsage.trigger).toMatch(/dysflow|\.dysflow\/project\.json/i);
});

it('lists dysflow-arnes as MUST-LOAD when .dysflow/project.json is in cwd', async () => {
  const skills = await client.listAvailableSkills({ cwd: '<test-repo-with-.dysflow>' });
  const arnes = skills.find(s => s.name === 'dysflow-arnes');
  expect(arnes).toBeDefined();
  expect(arnes.priority).toBe('must-load');
});
```

Test 2 — el bootstrap ladder está documentado en la descripción de la skill:

```ts
it('dysflow-usage description mentions the bootstrap ladder (get_capabilities first)', async () => {
  const skill = await client.readSkill({ name: 'dysflow-usage' });
  expect(skill.description).toMatch(/get_capabilities|start here|load first/i);
});

it('dysflow-arnes description references dysflow-usage as MUST-LOAD', async () => {
  const skill = await client.readSkill({ name: 'dysflow-arnes' });
  expect(skill.description).toMatch(/dysflow-usage.*must-load|load.*dysflow-usage.*first/i);
});
```

Test 3 — el bloque en AGENTS.md del consumer incluye la directiva:

```ts
it('AGENTS.md pointer block for dysflow includes MUST-LOAD directive for dysflow-usage', async () => {
  const agentsMd = await fs.readFile('<test-repo>/AGENTS.md', 'utf8');
  // Match the <!-- user-supplement:dysflow:pointer --> block
  const pointerBlock = agentsMd.match(/<!-- user-supplement:dysflow:pointer -->[\s\S]*?<!-- \/user-supplement:dysflow:pointer -->/);
  expect(pointerBlock).not.toBeNull();
  expect(pointerBlock[0]).toMatch(/dysflow-usage.*must-load/i);
});
```

### Bug B: los envelopes de error de dysflow no sugieren cargar `dysflow-usage`

#### Síntoma verificado

Cuando el consumer recibe errores de dysflow (`MCP_INPUT_INVALID`, `FRONTEND_TARGET_AMBIGUOUS`, `CONFIG_MISSING_ACCESS_PATH`, `LACCDB_STALE_DETECTED`, etc.), el envelope NO incluye un hint que dirija al agente a la skill relevante. El agente tiene que saber por sí mismo que la respuesta es `cargar dysflow-usage` y leer la sección correspondiente.

Errores reales observados en este round:

```
1. "CONFIG_MISSING_ACCESS_PATH: Access database path is required. Define .dysflow/project.json in the repository or pass accessDbPath explicitly."
   → Sin hint: "Load dysflow-usage § Multi-worktree operation" o "Run get_capabilities first".

2. "FRONTEND_TARGET_AMBIGUOUS: Multiple sibling worktree projects are visible from this cwd."
   → Sin hint: "Load dysflow-usage § HR-11 Recover ambiguity" o "Pass projectId + projectChoiceReason + recoveryToken".

3. "MCP_INPUT_INVALID: accessDbPath is not allowed. Valid params: projectId, contextId, accessPath, backendPath, destinationRoot, projectRoot, outputMode, filter, allowExternalAccessPath, timeoutMs, cwd. Did you mean 'accessPath'?"
   → Sin hint: "Run describe_tool({name:'<tool>'}) para verificar parámetros exactos".
```

#### Evidencia de repro

```ts
it('CONFIG_MISSING_ACCESS_PATH envelope references the dysflow-usage skill', async () => {
  const result = await client.call('list_tables', { projectId: 'non-existent' });
  expect(result.ok).toBe(false);
  expect(result.error.code).toBe('CONFIG_MISSING_ACCESS_PATH');
  expect(result.error.remediation.skill).toBe('dysflow-usage');
  expect(result.error.remediation.section).toMatch(/resolve_project|get_capabilities/i);
});

it('FRONTEND_TARGET_AMBIGUOUS envelope references the dysflow-usage skill', async () => {
  const result = await client.call('resolve_project', { cwd: '<repo-with-multiple-sibling-worktrees>' });
  expect(result.outcome).toBe('ambiguous');
  expect(result.remediation.skill).toBe('dysflow-usage');
  expect(result.remediation.section).toMatch(/HR-11|ambiguity/i);
});

it('MCP_INPUT_INVALID envelope references describe_tool for parameter validation', async () => {
  const result = await client.call('list_tables', { wrongParam: 'x' });
  expect(result.ok).toBe(false);
  expect(result.error.code).toBe('MCP_INPUT_INVALID');
  expect(result.error.remediation.tool).toBe('describe_tool');
  expect(result.error.remediation.hint).toMatch(/describe_tool.*<wrongParam>/i);
});
```

#### Diagnóstico preliminar

El contrato `dysflow.result/v1` ya tiene campos `error.remediation` en algunos envelopes (visto en `migrate_project_config`), pero otros envelopes solo tienen `error.message` plano. El maintainer necesita estandarizar `error.remediation = { skill: string, section?: string, tool?: string, hint?: string, command?: string }` en todos los errores tipados.

El campo `remediation` ya está parcialmente definido (visto en `MCP_INPUT_INVALID` con `rejectedFlag`, `rejectedFlags[]`, `toolCommitFlag`). El fix es extender el shape, no inventar uno nuevo.

#### Riesgo

- Cross-fleet: cada nuevo agente de IA entra al ecosistema dysflow sin saber qué hacer con un error. Inventa diagnósticos.
- Pérdida de productividad: el consumer de access2web-blueprint perdió ~30 minutos por 1 error. Multiplicar por 11+ proyectos consumer y por N nuevos agentes que entrenan al ecosistema = alto.
- Documentación sucia: el consumer commiteó docs con afirmaciones incorrectas porque la respuesta al error fue opaca.

#### Tests RED sugeridos

(Los mismos del Bug B arriba. Cada test verifica `error.remediation.skill` o `error.remediation.tool` en envelopes tipados.)

Cross-cutting test:

```ts
it('every typed dysflow error envelope includes remediation with skill or tool hint', async () => {
  const knownErrorCodes = ['MCP_INPUT_INVALID', 'CONFIG_MISSING_ACCESS_PATH', 'FRONTEND_TARGET_AMBIGUOUS',
                            'FRONTEND_TARGET_MISSING', 'FRONTEND_PATH_NOT_BASENAME', 'INHERITED_WORKTREE_MISMATCH',
                            'PROJECT_ID_COLLISION', 'PROJECT_ID_MISMATCH', 'LACCDB_STALE_DETECTED',
                            'LIVE_PROCESS_HOLDS_LACCDB', 'PROCEDURE_NOT_FOUND', 'PROCEDURE_NOT_CALLABLE',
                            'EXPORT_OVERWRITES_SOURCE_REQUIRES_CONFIRMATION', 'DESTINATION_ROOT_REQUIRED',
                            'MCP_PROCEDURE_NOT_ALLOWED', 'MCP_WRITES_DISABLED', 'PROJECT_CONFIG_NOT_WRITE_READY'];
  for (const code of knownErrorCodes) {
    // ...trigger each error and assert remediation is present
  }
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR).
- Conventional commits con scope apropiado (`skill-discoverability`, `error-remediation-hints`, o similar).
- NO tocar las capabilities que YA funcionan (citá cuáles en "Lo que YA funciona").
- NO cambiar el discriminador `schemaVersion: "dysflow.result/v1"` ni el contrato de respuesta exitosa.
- Mantener `writeExecutionPolicy: "safe-by-default"` y `effectiveDryRunDefault` por tool.
- Si el fix del Bug B requiere un campo nuevo (`remediation.skill`), mantener retrocompatibilidad: `remediation` puede ser string (legacy) o object (nuevo). Documentar ambos.
- Si el fix del Bug A requiere actualizar múltiples `AGENTS.md` cross-fleet, hacer un PR genérico de plantilla + un script de sincronización, no editar manualmente los 11+ repos.
- NO usar `I think`, `maybe`, `probably` en el código o en los tests. Solo hechos verificados.

## Acceptance output

- PR con tests RED → GREEN (≥5 tests: 3 del Bug A, 2 del Bug B + cross-cutting).
- Changelog en `CHANGELOG.md`:
  - `Improve: dysflow-usage and dysflow-arnes skills surface as MUST-LOAD when .dysflow/ artifacts are in cwd. (#<issue>)`
  - `Improve: typed error envelopes include remediation.skill / remediation.tool hints pointing to the canonical skill. (#<issue>)`
- Version bump: **minor** (`v2.36.0`) — cambia el contrato de error envelopes (nuevo campo `remediation.skill`).
- `verify-examples-vs-runtime.ps1` agrega tests de skill discoverability + remediation hints.
- Documentación actualizada:
  - `docs/prompts/prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md` (este archivo) referenciado en CHANGELOG.
  - `dysflow-usage` SKILL.md frontmatter incluye `description` con MUST-LOAD explícito.
  - `dysflow-arnes` SKILL.md cross-reference a `dysflow-usage` como primer paso.
- Plantilla AGENTS.md del consumer (en `dysflow/install` o equivalente) actualizada para incluir el bloque MUST-LOAD.

## Quick start

```bash
git clone <repo>
cd <repo>
git checkout -b fix/skill-discoverability-and-error-hints
pnpm install
pnpm run dev  # arranca el MCP localmente
```

Test repro contra el MCP local:

```bash
# 1. Skill discoverability (Bug A)
dysflow.list_available_skills({ cwd: '<test-repo-with-.dysflow>/' })
# Esperado: dysflow-usage con priority='must-load', trigger matching .dysflow/
# Actual: dysflow-usage aparece listada pero sin priority/trigger claros

# 2. Error remediation hints (Bug B)
dysflow.resolve_project({ cwd: '<repo-with-sibling-worktrees>/' })
# Esperado: result.remediation.skill === 'dysflow-usage'
# Actual: result.remediation es undefined o string genérico

# 3. Cross-cutting
dysflow.list_available_skills({ cwd: '<test-repo-without-.dysflow>/' })
# Esperado: dysflow-usage NO aparece con priority must-load (no es relevante)
# Actual: dysflow-usage aparece siempre con priority alta
```

## Reinforcement

Mantener la regla cross-project: **"Cuando un agente toca artefactos dysflow, cargar `dysflow-usage` ANTES de inspeccionar archivos estáticos, inventar diagnósticos o modificar configs."**

Esta regla debe reforzarse desde TRES frentes complementarios:

1. **Front de skills**: `dysflow-usage` y `dysflow-arnes` deben aparecer como MUST-LOAD en `<available_skills>` cuando hay `.dysflow/*` en cwd.
2. **Front de errores**: Cada envelope de error tipado debe incluir `remediation.skill` apuntando a la skill canónica.
3. **Front de templates**: AGENTS.md del consumer debe tener un bloque `<!-- user-supplement:dysflow:pointer -->` que cargue automáticamente la skill.

Si el fix solo cubre uno de los tres frentes, el problema persistirá parcialmente. El PR debe cubrir los tres.

## Referencias cruzadas

- Skill consumer-side `dysflow-arnes`: HR-0 a HR-13 (el arnés del agente). El maintainer debe validar que sus fixes siguen siendo compatibles con HR-5 ("Runtime is source of truth").
- Skill consumer-side `dysflow-usage`: contrato de `get_capabilities`, `resolve_project`, `register_worktree`, `effectiveDryRunDefault`. El maintainer debe asegurar que los nuevos campos en error envelopes NO rompen este contrato.
- Cross-fleet impact: 11+ proyectos consumer (HPS, gestion_riesgos, no_conformidades, condor, cadete, brass, expedientes, lanzadera, apap, hps_solicitudes, etc.). La regresión cruza proyectos.
- Memoria consumer `personal/skill-discoverability-dysflow` (access2web-blueprint) y `personal/skill-discoverability-dysflow-error-hints` — registradas en este round.