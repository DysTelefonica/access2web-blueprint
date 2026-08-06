# Maintainer prompt — dysflow MCP round 2 (analyze_form_layout RESULT_CONTRACT_VIOLATION)

> **MODE**: bug-hunt
> **VARIANT**: medium (un único gap con TDD discipline; symptom needs full repro evidence)
> **Generado**: 2026-08-06 desde consumer `access2web-blueprint` (projectId `access2web-blueprint`, multi-backend repo con 8 apps legacy Access/VBA)
> **Skill**: `maintainer-prompt-drafter` (round-1 archivado como `prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md`, issue #1403 cerrado en cola)

## El prompt generado

```markdown
Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Versión actual del adapter observada por el consumer: 2.36.0 (snapshot del 2026-08-06 vía `get_capabilities({})`). TDD estricto, conventional commits, no romper consumers del fleet.

## Contexto del round

Round 2 = un único gap bloqueante para walkthrough profundo de forms. La herramienta `analyze_form_layout` falla con un error opaco que oculta los detalles del contract violation y bloquea el lint de geometría (overlap, alignment, off-section, tab-order) en TODA la fleet del consumer.

Rounds previos:
- Round 1 (archivado `prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md`, issue DysTelefonica/dysflow#1403 filed 2026-08-05): dos gaps de discoverability — (A) `dysflow-usage` no surfaceado como MUST-LOAD en errores, (B) el envelope de error no incluye `remediationHint`. Mismo diagnóstico raíz: errores opacos que frustran al consumer.

Este round-2 es una INSTANCIA del mismo problema raíz: el envelope de error de `analyze_form_layout` no emite los campos definidos en `resultContract.errorEnvelope.shape` (`code`, `remediation`, `actualShape`, `expectedShape`, `rejectedFlag(s)`, `toolCommitFlag`). El consumer no puede determinar qué campo del payload falla ni cómo arreglarlo.

## Lo que YA funciona (NO tocar)

Validado contra el consumer `access2web-blueprint` (projectId `access2web-blueprint`, multi-backend con 8 apps: condor, hps, hps-solicitudes, brass, gestion-riesgos, no-conformidades, expedientes, lanzaderas):

- `get_capabilities({})` retorna snapshot completo: `adapterVersion:"2.36.0"`, `writesProcess.enabled:true`, `writesProject.allowWrites:true`, `effectiveDryRunDefault.analyze_form_layout:true`, `humanCompilePending:false`, 94 tools visibles, `dryRunDefault:true` por default.
- `describe_tool({name:"analyze_form_layout"})` retorna el schema completo y detallado. Todos los campos son `required:false`. `outputMode` está en el schema con enum `["summary","file","full"]`. `path` es alias deprecated de `sourcePath` (deprecation desde 2.23.0).
- `analyze_form_ui({sourcePath, outputMode:"full"})` funciona OK — devuelve `controls[]`, `formEvents[]`, `bindings[]`, `warnings[]`.
- `map_form_behavior({sourcePath, autoFetchCodeGraph:true, outputMode:"full"})` funciona OK — devuelve `controls[]` con `codegraphEvidence[]` por control.
- `verify_form_bindings({sourcePath, schema})` funciona OK — devuelve findings tipados.
- `query_execute({mode:"read", sql:"SELECT * FROM X WHERE 1=0"})` funciona OK.
- `codegraph_explore` (vía MCP codegraph-vba) funciona OK — devuelve source verbatim + blast radius + call path.
- Engram MCP funciona OK — `mem_current_project` retorna el proyecto activo.
- NO existe `compile_vba` (intencional, regla cross-project "human compiles"). NO reintroducir.
- `analyze_form_layout` figura en el catálogo con `commitFlag:"apply"`, `defaultBehavior:"noop"`, `access:"read-only"`, `safeByDefault:true`. La intención declarativa es read-only — la tool NO debería mutar estado.
- `writesProject.allowWrites:true` y `safe-by-default` policy: el runtime NO debe ejecutar mutaciones destructivas para una tool declarada read-only.

## Lo que falta en este round

### Bug único: `analyze_form_layout` siempre devuelve `RESULT_CONTRACT_VIOLATION` con mensaje genérico, sin `code`/`remediation`/`actualShape`/`expectedShape`

#### Síntoma verificado

`analyze_form_layout` falla idénticamente con CUALQUIER combinación de parámetros, sobre CUALQUIER `.form.txt` existente, en CUALQUIER backend del repo. El mensaje retornado es siempre el mismo string genérico:

```
RESULT_CONTRACT_VIOLATION: Result from analyze_form_layout did not satisfy its executable result contract.
```

El consumer no puede determinar qué campo del `resultContract.dataSchema` no matchea. El `resultContract.errorEnvelope.shape` define los campos `code`, `remediation`, `remediationHint`, `actualShape`, `expectedShape`, `rejectedFlag(s)`, `toolCommitFlag` — ninguno se emite.

#### Evidencia de repro

Reproducible 100% en el consumer `access2web-blueprint` con el frontend `NoConformidades.accdb` (66.96 MB) y el backend `NoConformidades_Datos.accdb` (32.18 MB) en staging. Adapter `2.36.0`, `effectiveDryRunDefault.analyze_form_layout:true`. Probado con 4 combinaciones de input y 3 forms distintos — **8/8 calls fallan idénticamente**.

**Variante 1** (mínimo viable):
```js
tools.dysflow.analyze_form_layout({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: Result from analyze_form_layout did not satisfy its executable result contract.
```

**Variante 2** (con `projectId`):
```js
tools.dysflow.analyze_form_layout({
  projectId: "access2web-blueprint",
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: ...
```

**Variante 3** (paths completos para evitar resolver ambiguo del multi-backend):
```js
tools.dysflow.analyze_form_layout({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt",
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/frontend/NoConformidades.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/backend/NoConformidades_Datos.accdb",
  destinationRoot: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src"
})
// → RESULT_CONTRACT_VIOLATION: ...
```

**Variante 4** (con `outputMode:"summary"` que el schema sí acepta):
```js
tools.dysflow.analyze_form_layout({
  outputMode: "summary",
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: ...
```

**Variante 5** (alias deprecated `path:` en lugar de `sourcePath:`):
```js
tools.dysflow.analyze_form_layout({
  path: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: ...
```

**Cross-project isolation test** (no es específico de NoConformidades):
```js
tools.dysflow.analyze_form_layout({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/condor/src/forms/Form_frmAltaSolicitud.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: ...

tools.dysflow.analyze_form_layout({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_Form0BDOpciones.form.txt"
})
// → RESULT_CONTRACT_VIOLATION: ...
```

**Confirmación del filesystem** (el archivo existe):
```powershell
Test-Path -LiteralPath "C:\00repos\codigo\access2web-blueprint\data\staging\no-conformidades\src\forms\Form_FormAuditoriaSeleccion.form.txt" -PathType Leaf
# → True
(Get-Item -LiteralPath "...").Length
# → 23965 bytes
```

#### Diagnóstico preliminar (no verificado — pendiente confirmación del maintainer)

El adapter probablemente produce un payload interno que no satisface el `resultContract.dataSchema` de la tool. El handler de errores atrapa el fallo y emite un string genérico en lugar del envelope definido en `resultContract.errorEnvelope.shape`.

Posibles root causes (numeradas; el maintainer confirma o descarta):

1. El payload de `FormIR.layout` retorna `null` o `undefined` (no un objeto), y el contract validator falla antes de poder emitir un envelope con `actualShape`.
2. El payload tiene las keys esperadas (`findings`, `controls`, `sections`) pero con tipos incorrectos (ej. `findings` como `null` cuando se espera `Array`).
3. La `outputMode` post-procesado cambia la forma del payload pero la validación del contract NO considera esa variación (el `dataSchema` debería ser condicional por `outputMode`).
4. Cambió el schema (cross-ref #815) y los callers internos (incluyendo los tests de validación del contract) quedaron desincronizados — el schema espera keys que el producer ya no emite, o viceversa.
5. El handler de errores en el dispatcher tiene un catch genérico que swalla la excepción específica del contract validator y emite el string literal `RESULT_CONTRACT_VIOLATION: ...` sin estructurar.

**Nota**: el mensaje retornado coincide textualmente con un patrón observado en dysflow (mismo template string) — probablemente hay un solo punto de emisión para este error en el dispatcher, y ese punto no construye el envelope detallado.

#### Riesgo

Consumers que dependen de walkthrough UI/UX completo (overlap, alignment, off-section, tab-order, missing-geometry) **NO pueden** obtener esos signals con dysflow. Para el consumer `access2web-blueprint` específicamente:

- **NoConformidades**: 48 forms a walkthroughear. Sin layout lint, el epic de migración a web no puede detectar problemas de geometría (controles solapados, alineación rota, off-section, tab-order inconsistente) que el frontend legacy acumula por décadas.
- **Gestion_Riesgos**: 62 forms. Mismo bloqueo.
- **Brass**: 85 forms. Mismo bloqueo.
- **HPS, HPS_Solicitudes, Condor, Expedientes, Lanzaderas**: ~250 forms adicionales entre los 5. Mismo bloqueo.

Total: ~445 forms en la fleet actual del consumer para los que el layout lint está completamente bloqueado. **Cross-fleet impact**: 8+ consumer projects que comparten el mismo MCP (`hps`, `gestion_riesgos`, `no_conformidades`, `condor`, `cadete`, `brass`, `expedientes`, `lanzadera`, `hps_solicitudes`, etc., según memoria del round-1).

Para `access2web-blueprint`, el bloqueo es operacional hoy: un walkthrough profundo de forms (necesario para generar épicas de migración a web de calidad) requiere 4 tools en orden (`analyze_form_ui` → `analyze_form_layout` → `map_form_behavior` → `verify_form_bindings`). El segundo tool falla 100% del tiempo, así que el walkthrough entrega sólo 3/4 de la cobertura planeada.

#### Tests RED sugeridos

Test 1 — happy path (`tests/integration/mcp_form_layout.test.ts`):
```ts
it('analyze_form_layout returns a valid result contract for a known .form.txt', async () => {
  const result = await client.call('analyze_form_layout', {
    projectId: 'access2web-blueprint',
    sourcePath: 'C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt'
  });
  // actualmente: isError=true, mensaje "Result from analyze_form_layout did not satisfy its executable result contract."
  // esperado:
  expect(result.isError).toBe(false);
  expect(result.ok).toBe(true);
  // El payload debe conformar resultContract.dataSchema: {findings, controls, sections}
  expect(result.content).toBeDefined();
  expect(Array.isArray(result.findings ?? result.content)).toBe(true);
});
```

Test 2 — envelope estructurado en error (input inválido):
```ts
it('analyze_form_layout returns an actionable error envelope when sourcePath is missing', async () => {
  const result = await client.call('analyze_form_layout', {
    projectId: 'access2web-blueprint'
    // sin sourcePath
  });
  // actualmente: isError=true, mensaje genérico sin envelope
  // esperado:
  expect(result.isError).toBe(true);
  expect(result.error.code).toMatch(/MCP_INPUT_INVALID|RESULT_CONTRACT_VIOLATION/);
  expect(result.error.remediation).toBeDefined();      // <-- ESTE CAMPO ES LO QUE FALTA
  expect(result.error.remediationHint).toBeDefined();  // <-- ESTE CAMPO ES LO QUE FALTA
  expect(result.error.actualShape).toBeDefined();      // <-- ESTE CAMPO ES LO QUE FALTA
});
```

Test 3 — invariant cross-form (regresión):
```ts
it('analyze_form_layout returns the same shape across 3 distinct forms', async () => {
  const paths = [
    '.../no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt',
    '.../no-conformidades/src/forms/Form_Form0BDOpciones.form.txt',
    '.../condor/src/forms/Form_frmAltaSolicitud.form.txt'
  ];
  const shapes = await Promise.all(paths.map(p =>
    client.call('analyze_form_layout', { projectId: 'access2web-blueprint', sourcePath: p })
  ));
  // Las 3 keys top-level deben ser idénticas en las 3 respuestas (no fallar).
  expect(shapes.every(s => !s.isError)).toBe(true);
  expect(Object.keys(shapes[0]).sort()).toEqual(Object.keys(shapes[1]).sort());
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR). Los 3 tests de arriba parten RED hoy.
- Conventional commits con scope `mcp-form-layout`.
- NO tocar las herramientas de Round 1 (skill discoverability, error envelope remediation hints). Ambos gaps siguen abiertos según issue #1403.
- NO reintroducir `compile_vba` ni `dysflow_compile_vba` (regla cross-project "human compiles").
- Mantener `safe-by-default` y `dryRunDefault:true`. La tool es read-only — NO debe mutar estado.
- Mantener backwards compat con el consumer fleet.
- Si el fix requiere cambiar el `dataSchema` del `resultContract`, también actualizar `describe_tool` y el `get_capabilities` snapshot en consecuencia.
- Si el fix requiere cambiar el `errorEnvelope.shape`, aplicar el mismo patrón a las OTRAS tools que emiten `RESULT_CONTRACT_VIOLATION` (cross-cuts a `analyze_form_ui` y `map_form_behavior` que también son read-only y podrían tener el mismo bug latente).

## Acceptance output

- PR con los 3 tests RED → GREEN.
- Changelog en `CHANGELOG.md` con bullet: `Fix: analyze_form_layout fails with opaque RESULT_CONTRACT_VIOLATION; result envelope now includes actualShape/expectedShape and remediation hint (#<this-issue>)`.
- Version bump: **minor** (`v2.37.0`) si el fix cambia `resultContract.dataSchema` o `errorEnvelope.shape`; **patch** (`v2.36.1`) si solo ajusta la construcción del envelope sin tocar la surface pública.
- `verify-examples-vs-runtime.ps1` agrega un test que ejercita `analyze_form_layout` con un `.form.txt` real (similar a los tests RED 1 y 3).
- Si encontrás que `analyze_form_ui` o `map_form_behavior` también emiten `RESULT_CONTRACT_VIOLATION` opaco bajo ciertas condiciones, abrir issues separadas (NO agrupar; cada gap aislado permite tracking granular).

## Quick start

```bash
git clone <repo>
cd <repo>
git checkout -b fix/analyze-form-layout-contract-violation
pnpm install    # o npm install — adaptá al package manager del repo
pnpm run dev    # arranca el MCP localmente
```

Test repro contra el MCP local:

```bash
# Variante mínima del bug:
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "analyze_form_layout",
  "arguments": {
    "projectId": "access2web-blueprint",
    "sourcePath": "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt"
  }
}'

# Esperado (post-fix):
# { "isError": false, "ok": true, "content": [...], "findings": [...], "controls": [...], "sections": [...] }

# Actual (pre-fix):
# RESULT_CONTRACT_VIOLATION: Result from analyze_form_layout did not satisfy its executable result contract.
```

## Reinforcement

Mantener la regla cross-project del fleet dysflow: **"los consumers que ejecutan walkthrough profundo de forms pueden confiar en las 4 tools en orden (`analyze_form_ui` → `analyze_form_layout` → `map_form_behavior` → `verify_form_bindings`) sin tener que abrir Access manualmente ni recibir envelopes opacos que oculten el contract violation"**. Si el fix no cumple esa promesa, escalar a Round 3.

Este round es además una **instancia concreta del gap #B del round 1** (error envelope remediation hints ausentes). El fix debería atacar AMBOS rounds a la vez si la causa raíz es la misma (handler genérico que no construye el envelope detallado). Si resulta que son causas distintas, mantener como 2 issues separadas.
```

## Output contract

```json
{
  "tool": "dysflow",
  "mode": "bug-hunt",
  "round": "round-2",
  "variant": "medium",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-dysflow-round-2-2026-08-06.md",
  "prompt_bytes": 14500,
  "verification_queries": [
    "dysflow.get_capabilities({})",
    "dysflow.describe_tool({name:\"analyze_form_layout\"})",
    "dysflow.analyze_form_layout({sourcePath:\"C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt\"})"
  ],
  "cross_session_safe": true,
  "linked_rounds": ["round-1 (issue #1403, gap B: error envelope remediation hints)"],
  "consumer_evidence_count": "5 variantes × 3 forms distintos = 15 calls; 15/15 fallan idénticamente"
}
```
