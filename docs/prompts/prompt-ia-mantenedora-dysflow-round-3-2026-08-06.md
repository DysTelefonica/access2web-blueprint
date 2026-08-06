# Maintainer prompt — dysflow MCP round 3 (codegraph-vba CLI --json flag rejected)

> **MODE**: bug-hunt
> **VARIANT**: short (gap chico bien acotado; el fix vive en dysflow↔codegraph-vba integration seam)
> **Generado**: 2026-08-06 desde consumer `access2web-blueprint` (projectId `access2web-blueprint`, multi-backend repo con 8 apps legacy Access/VBA)
> **Skill**: `maintainer-prompt-drafter`

## El prompt generado

```markdown
Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Versión actual del adapter observada por el consumer: 2.36.0 (snapshot del 2026-08-06 vía `get_capabilities({})`). TDD estricto, conventional commits, no romper consumers del fleet.

## Contexto del round

Round 3 = un único gap chico en la integración dysflow→codegraph-vba CLI. La tool `map_form_behavior({autoFetchCodeGraph:true})` invoca el CLI `codegraph-vba.cmd explore --json ...` con un flag `--json` que el CLI codegraph-vba fork (vba-aware) **rechaza**. Resultado: `codegraphEvidence[]` siempre viene vacío (con warning). El MCP `codegraph_explore` directo funciona OK; el problema es solo en la cadena dysflow→codegraph-vba CLI subprocess.

Rounds previos (siguen abiertos, NO tocar):
- Round 1 (archivado `prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md`, issue DysTelefonica/dysflow#1403): skill discoverability + error envelope remediation hints.
- Round 2 (archivado `prompt-ia-mantenedora-dysflow-round-2-2026-08-06.md`, issue DysTelefonica/dysflow#1407): `analyze_form_layout` returns opaque `RESULT_CONTRACT_VIOLATION`; result envelope missing `actualShape`/`expectedShape`/`remediation`.

Este round-3 es OTRO facet del mismo problema raíz que rounds 1 y 2: el adapter dysflow tiene un integration seam con codegraph-vba CLI mal negociado. A diferencia de round 2, este gap **NO es bloqueante** (dysflow tiene fallback a form-declared events con warning), pero pierde valor de `codegraphEvidence[]` para todos los consumers que dependen de auto-fetch.

## Lo que YA funciona (NO tocar)

Validado contra el consumer `access2web-blueprint`:

- `get_capabilities({})` retorna snapshot completo (mismo del round-2).
- `describe_tool({name:"map_form_behavior"})` declara `autoFetchCodeGraph:boolean, outputMode:"summary"|"file"|"full"`. El fallback a form-declared events cuando codegraph fetch falla funciona (NO rompe la tool, solo emite warning).
- `analyze_form_ui`, `verify_form_bindings`, `query_execute` funcionan OK (cubierto en round-2).
- `codegraph_explore` vía MCP codegraph-vba directo funciona OK y devuelve source verbatim + blast radius + call path.
- El `codegraphIndexPath` se setea correctamente: `C:\00repos\codigo\access2web-blueprint\.codegraph-vba`.
- `map_form_behavior` SIN `autoFetchCodeGraph:true` funciona OK y devuelve controls + form-declared events.
- `safe-by-default` policy, `dryRunDefault:true`, TDD gates (`validate_manifest`, `run_vba`, `test_vba`).

## Lo que falta en este round

### Bug único: `map_form_behavior({autoFetchCodeGraph:true})` falla al invocar `codegraph-vba.cmd explore --json` — flag `--json` no soportado por el CLI codegraph-vba fork

#### Síntoma verificado

`map_form_behavior` con `autoFetchCodeGraph:true` siempre emite el siguiente warning textual y devuelve `codegraphEvidence:[]` para todos los controles:

```
CodeGraph-VBA lookup failed for form "FormAuditoriaSeleccion":
  Command failed: C:\\WINDOWS\\system32\\cmd.exe /d /s /c "\"codegraph-vba.cmd\" \"explore\" \"--json\" \"--path\" \"C:\\00repos\\codigo\\access2web-blueprint\" \"--max-files\" \"16\" \"form:FormAuditoriaSeleccion controls:EncabezadoDelFormulario,lblTitulo,Detalle,Auditoria,Etiqueta86,PieDelFormulario,ComandoSeleccionar,cmdSalir\""
  error: unknown option '--json'
. Falling back to .form.txt-declared events only.
```

La tool NO falla (correcto, hay fallback), pero `codegraphEvidence[]` queda vacío. El consumer pierde el cross-reference de handlers con codegraph que es justo lo que `autoFetchCodeGraph:true` promete.

#### Evidencia de repro

```js
tools.dysflow.map_form_behavior({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt",
  autoFetchCodeGraph: true,
  outputMode: "summary"
})
```

Retorna:
```json
{
  "formName": "FormAuditoriaSeleccion",
  "controls": [
    {"name":"EncabezadoDelFormulario", ..., "codegraphEvidence": []},
    {"name":"lblTitulo", ..., "codegraphEvidence": []},
    {"name":"Detalle", ..., "codegraphEvidence": []},
    {"name":"Auditoria", ..., "codegraphEvidence": []},
    {"name":"Etiqueta86", ..., "codegraphEvidence": []},
    {"name":"PieDelFormulario", ..., "codegraphEvidence": []},
    {"name":"ComandoSeleccionar", ..., "codegraphEvidence": []},
    {"name":"cmdSalir", ..., "codegraphEvidence": []}
  ],
  "warnings": [
    "No CodeGraph-VBA evidence was supplied.",
    "CodeGraph-VBA lookup failed for form \"FormAuditoriaSeleccion\": ... error: unknown option '--json'. Falling back to .form.txt-declared events only."
  ],
  "codegraphIndexPath": "C:\\00repos\\codigo\\access2web-blueprint\\.codegraph-vba"
}
```

Cross-check: el MCP `codegraph_explore` directo sobre el mismo query SÍ devuelve símbolos + blast radius. El bug es específico del integration seam dysflow→codegraph-vba CLI.

#### Diagnóstico preliminar (no verificado — pendiente confirmación del maintainer)

dysflow adapter pasa `--json` al CLI codegraph-vba; el CLI codegraph-vba fork (vba-aware) rechaza ese flag con `error: unknown option '--json'`.

Posibles root causes (numeradas; el maintainer confirma o descarta):

1. dysflow está hardcodeado a pasar `--json` asumiendo que codegraph-vba upstream lo acepta, pero el fork vba-aware del consumer (`C:\00repos\codigo\access2web-blueprint\.codegraph-vba`) NO lo implementó.
2. El CLI codegraph-vba fork acepta otro flag para output JSON (ej. `--format=json`, `--output=json`, o sin flag y stdout es texto).
3. El CLI codegraph-vba fork espera stdin JSON en lugar de flag CLI.
4. La negociación de capabilities entre dysflow y codegraph-vba es estática (no detecta qué flags acepta el fork).

#### Riesgo

Consumers que dependen de `codegraphEvidence[]` para walkthrough profundo (cross-reference de handlers con tablas, blast radius, call path) NO pueden auto-obtenerlo. Deben llamar `codegraph_explore` directo y mergear manualmente. Para el consumer `access2web-blueprint` específicamente, esto afecta a 8 apps × ~50 forms = ~400 walks; cada uno pierde `codegraphEvidence[]` por auto-fetch.

NO es bloqueante (fallback funciona), pero es fricción operacional alta.

#### Tests RED sugeridos

Test 1 — happy path (`tests/integration/mcp_form_behavior_autofetch.test.ts`):
```ts
it('map_form_behavior autoFetchCodeGraph returns non-empty codegraphEvidence when CodeGraph index exists', async () => {
  const result = await client.call('map_form_behavior', {
    projectId: 'access2web-blueprint',
    sourcePath: 'C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt',
    autoFetchCodeGraph: true,
    outputMode: 'summary'
  });
  // actualmente: warnings contiene "unknown option '--json'", codegraphEvidence[] siempre []
  // esperado:
  expect(result.isError).toBe(false);
  expect(result.warnings ?? []).not.toContain(/unknown option '--json'/);
  // al menos UN control debe tener codegraphEvidence no-vacía (cross-ref real con codegraph)
  const anyEvidence = result.controls.some(c => (c.codegraphEvidence ?? []).length > 0);
  expect(anyEvidence).toBe(true);
});
```

Test 2 — regression: el fallback sigue funcionando cuando codegraph no está disponible:
```ts
it('map_form_behavior still returns form-declared events when codegraph CLI is unavailable', async () => {
  // simular codegraph-vba.cmd no en PATH
  const result = await client.call('map_form_behavior', {
    projectId: 'access2web-blueprint',
    sourcePath: '.../Form_FormAuditoriaSeleccion.form.txt',
    autoFetchCodeGraph: true,
    outputMode: 'summary'
  });
  expect(result.isError).toBe(false); // fallback NO rompe la tool
  expect(result.controls.length).toBeGreaterThan(0); // controles siguen presentes
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR). Los 2 tests de arriba parten RED hoy.
- Conventional commits con scope `mcp-form-behavior-autofetch`.
- NO tocar las herramientas de Round 1/2 (skill discoverability, error envelope remediation, analyze_form_layout).
- Mantener el fallback actual (si codegraph falla, emitir warning y devolver form-declared events). El fix debe MEJORAR el happy path, NO romper el fallback.
- Mantener `safe-by-default` y `dryRunDefault:true`.
- Si el fix requiere detectar capabilities del fork codegraph-vba (qué flags acepta), documentar el contrato claramente en la skill `dysflow-usage`.

## Acceptance output

- PR con los 2 tests RED → GREEN.
- Changelog en `CHANGELOG.md` con bullet: `Fix: map_form_behavior autoFetchCodeGraph falls back to form-declared events due to codegraph-vba CLI rejecting --json flag; adapter now negotiates compatible output format (#<this-issue>)`.
- Version bump: **patch** (`v2.36.1`).
- `verify-examples-vs-runtime.ps1` agrega un test que ejercita `map_form_behavior` con `autoFetchCodeGraph:true` contra un índice `.codegraph-vba` real.

## Quick start

```bash
git clone <repo>
cd <repo>
git checkout -b fix/map-form-behavior-autofetch-codegraph-cli
pnpm install
pnpm run dev
```

Test repro contra el MCP local:

```bash
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "map_form_behavior",
  "arguments": {
    "projectId": "access2web-blueprint",
    "sourcePath": "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt",
    "autoFetchCodeGraph": true,
    "outputMode": "summary"
  }
}'

# Esperado (post-fix):
# { "warnings": [], "controls": [{..., "codegraphEvidence": [{"handler": "...", "callPath": [...], "tables": [...]}]}] }

# Actual (pre-fix):
# { "warnings": ["...error: unknown option '--json'..."], "controls": [{..., "codegraphEvidence": []}, ...] }
```
```

## Output contract

```json
{
  "tool": "dysflow",
  "mode": "bug-hunt",
  "round": "round-3",
  "variant": "short",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-dysflow-round-3-2026-08-06.md",
  "prompt_bytes": 8800,
  "verification_queries": [
    "dysflow.get_capabilities({})",
    "dysflow.map_form_behavior({sourcePath:\"...\", autoFetchCodeGraph:true, outputMode:\"summary\"})"
  ],
  "cross_session_safe": true,
  "linked_rounds": [
    "round-1 (issue #1403, gap B: error envelope remediation hints)",
    "round-2 (issue #1407, analyze_form_layout RESULT_CONTRACT_VIOLATION opaco)"
  ],
  "consumer_evidence_count": "1 form × 1 tool variant = 1 call; falla 100% del tiempo con warning explícito"
}
```
