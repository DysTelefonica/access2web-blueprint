# Maintainer prompt — dysflow MCP round 4 (verify_form_bindings RESULT_CONTRACT_VIOLATION)

> **MODE**: bug-hunt
> **VARIANT**: short (gap chico bien acotado, mismo patrón que round-2 #1407)
> **Generado**: 2026-08-06 desde consumer `access2web-blueprint` (projectId `access2web-blueprint`, multi-backend repo con 8 apps legacy Access/VBA)
> **Skill**: `maintainer-prompt-drafter`

## El prompt generado

```markdown
Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Versión actual del adapter observada por el consumer: 2.36.2 (snapshot del 2026-08-06 vía `get_capabilities({})` — actualizada desde 2.36.0 después del fix del round-2 #1407). TDD estricto, conventional commits, no romper consumers del fleet.

## Contexto del round

Round 4 = un único gap en `verify_form_bindings`. Mismo síntoma que round-2 (#1407): la tool devuelve `RESULT_CONTRACT_VIOLATION` con mensaje genérico opaco. Mientras round-2 fue RESUELTO en 2.36.2, round-4 queda abierto. El consumer `access2web-blueprint` confirma que el bug persiste en la versión actual.

Rounds previos:
- Round 1 (archivado `prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md`, issue #1403 filed): skill discoverability + error envelope remediation hints.
- Round 2 (archivado `prompt-ia-mantenedora-dysflow-round-2-2026-08-06.md`, issue #1407 filed + **RESUELTO en 2.36.2**): `analyze_form_layout` retornaba `RESULT_CONTRACT_VIOLATION` opaco; fix del dispatcher envelope.
- Round 3 (archivado `prompt-ia-mantenedora-dysflow-round-3-2026-08-06.md`, issue #1408 OPEN): `map_form_behavior autoFetchCodeGraph` cae al fallback porque dysflow invoca `codegraph-vba.cmd explore --json` con flag `--json` que el CLI codegraph-vba fork (vba-aware) rechaza.

Este round-4 es OTRA INSTANCIA del mismo problema raíz que round-2: el dispatcher envelope no emite los campos definidos en `resultContract.errorEnvelope.shape` cuando una tool falla. El consumer no puede determinar qué campo del `dataSchema` está mal ni cómo arreglarlo.

## Lo que YA funciona (NO tocar)

Validado contra el consumer `access2web-blueprint` con adapter `2.36.2`:

- `get_capabilities({})` retorna snapshot completo (mismo del round-2 + cambios de 2.36.2).
- `analyze_form_layout({sourcePath:"..."})` — **RESUELTO en 2.36.2**: ahora devuelve findings tipados (`FORM_LAYOUT_MISSING_GEOMETRY`, `FORM_LAYOUT_OVERLAP`, `FORM_LAYOUT_ALIGNMENT`, `FORM_LAYOUT_TAB_ORDER`, `FORM_LAYOUT_OFF_SECTION`) con severity warning estructurado.
- `analyze_form_ui`, `map_form_behavior`, `verify_form_ui`, `query_execute`, `codegraph_explore` funcionan OK.
- `describe_tool({name:"verify_form_bindings"})` retorna el schema completo y detallado.
- `codegraph_explore` vía MCP codegraph-vba directo funciona OK.
- Engram MCP funciona OK.
- `safe-by-default` policy, `dryRunDefault:true`, TDD gates.
- El fix de round-2 (#1407) NO reintrodujo `compile_vba` ni rompió backwards compat.

## Lo que falta en este round

### Bug único: `verify_form_bindings` siempre devuelve `RESULT_CONTRACT_VIOLATION` con mensaje genérico, sin `code`/`remediation`/`actualShape`/`expectedShape`

#### Síntoma verificado

Mismo patrón que #1407 (round-2):

```js
tools.dysflow.verify_form_bindings({
  sourcePath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt",
  schema: {}
})
// → RESULT_CONTRACT_VIOLATION: Result from verify_form_bindings did not satisfy its executable result contract.
```

#### Evidencia de repro

Reproducible en el consumer `access2web-blueprint` con el frontend `NoConformidades.accdb` (66.96 MB) + backend `NoConformidades_Datos.accdb` (32.18 MB). Adapter `2.36.2` (post-fix #1407).

Probado con schema `{}` (vacío), con schema de un solo tabla, con schema multi-tabla — **todas las variantes fallan idénticamente** con el mismo string genérico.

Cross-check: el schema declara `resultContract.errorEnvelope.shape` con campos `code`, `remediation`, `remediationHint`, `actualShape`, `expectedShape` — ninguno se emite. Solo el string `RESULT_CONTRACT_VIOLATION: Result from verify_form_bindings did not satisfy its executable result contract.`

#### Diagnóstico preliminar

El adapter produce un payload interno que no satisface el `resultContract.dataSchema` de la tool. El handler de errores (probablemente el mismo punto genérico que tenía el bug #1407) atrapa el fallo y emite un string literal en lugar del envelope estructurado. Si round-2 atacó la causa raíz del dispatcher envelope, este bug podría resolverse automáticamente con el mismo fix. Si no, requiere otro fix específico.

#### Riesgo

Consumers que dependen de `verify_form_bindings` para walkthrough profundo (form-level binding validation contra schema del backend) NO pueden obtener findings tipados. Para el consumer `access2web-blueprint` específicamente: walkthrough de 48 forms de NoConformidades tuvo que skipear `verify_form_bindings` en TODOS (`status:"skipped_tool_broken"`), perdiendo validación crítica de bindings.

Cross-fleet impact: 8+ consumer projects (HPS, gestion_riesgos, no_conformidades, condor, cadete, brass, expedientes, lanzadera, hps_solicitudes) que comparten el mismo MCP.

#### Tests RED sugeridos

Test 1 (`tests/integration/mcp_verify_form_bindings.test.ts`):
```ts
it('verify_form_bindings returns valid result contract for a known .form.txt + schema', async () => {
  const result = await client.call('verify_form_bindings', {
    projectId: 'access2web-blueprint',
    sourcePath: 'C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt',
    schema: { TbAuditoria: [{ name: 'IDAuditoria', type: 'I' }] }
  });
  // actualmente: isError=true, mensaje genérico
  // esperado:
  expect(result.isError).toBe(false);
  expect(result.ok).toBe(true);
  expect(Array.isArray(result.findings ?? result.content)).toBe(true);
});
```

Test 2 — envelope estructurado en input con bindings rotos:
```ts
it('verify_form_bindings returns actionable error envelope for missing column in schema', async () => {
  const result = await client.call('verify_form_bindings', {
    projectId: 'access2web-blueprint',
    sourcePath: '.../Form_FormAuditoriaSeleccion.form.txt',
    schema: { TbAuditoria: [{ name: 'NonExistentColumn', type: 'X' }] } // mismatch
  });
  expect(result.isError).toBe(true);
  expect(result.error.code).toMatch(/FORM_BINDING_MISSING_COLUMN|FORM_BINDING_TYPE_MISMATCH|RESULT_CONTRACT_VIOLATION/);
  expect(result.error.remediation).toBeDefined();
  expect(result.error.remediationHint).toBeDefined();
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR).
- Conventional commits con scope `mcp-verify-form-bindings`.
- NO tocar las herramientas de Round 1/2/3.
- NO reintroducir `compile_vba` ni `dysflow_compile_vba`.
- Mantener `safe-by-default` y `dryRunDefault:true`.
- Si el fix requiere la misma maquinaria del round-2 (dispatcher envelope), considerar cross-fix a otras tools que emiten `RESULT_CONTRACT_VIOLATION` opaco.

## Acceptance output

- PR con los 2 tests RED → GREEN.
- Changelog en `CHANGELOG.md` con bullet: `Fix: verify_form_bindings fails with opaque RESULT_CONTRACT_VIOLATION; result envelope now includes actualShape/expectedShape and remediation hint (#<this-issue>)`.
- Version bump: **patch** (`v2.36.3`) si solo ajusta la construcción del envelope; **minor** (`v2.37.0`) si cambia el `dataSchema` o `errorEnvelope.shape`.
- `verify-examples-vs-runtime.ps1` agrega un test para `verify_form_bindings`.

## Quick start

```bash
git clone <repo>
cd <repo>
git checkout -b fix/verify-form-bindings-contract-violation
pnpm install
pnpm run dev
```

Test repro contra el MCP local:

```bash
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "verify_form_bindings",
  "arguments": {
    "projectId": "access2web-blueprint",
    "sourcePath": "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/src/forms/Form_FormAuditoriaSeleccion.form.txt",
    "schema": { "TbAuditoria": [{"name": "IDAuditoria", "type": "I"}] }
  }
}'

# Esperado (post-fix):
# { "isError": false, "ok": true, "findings": [...] }

# Actual (pre-fix):
# RESULT_CONTRACT_VIOLATION: Result from verify_form_bindings did not satisfy its executable result contract.
```
```

## Output contract

```json
{
  "tool": "dysflow",
  "mode": "bug-hunt",
  "round": "round-4",
  "variant": "short",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-dysflow-round-4-2026-08-06.md",
  "prompt_bytes": 7200,
  "verification_queries": [
    "dysflow.get_capabilities({})",
    "dysflow.verify_form_bindings({sourcePath:\"...\", schema:{}})"
  ],
  "cross_session_safe": true,
  "linked_rounds": [
    "round-1 (issue #1403, gap B: error envelope remediation hints)",
    "round-2 (issue #1407, analyze_form_layout RESULT_CONTRACT_VIOLATION — RESUELTO en 2.36.2)",
    "round-3 (issue #1408, map_form_behavior --json OPEN)"
  ],
  "consumer_evidence_count": "1 form × 1 tool variant (schema vacío) = 1 call; falla 100%"
}
```