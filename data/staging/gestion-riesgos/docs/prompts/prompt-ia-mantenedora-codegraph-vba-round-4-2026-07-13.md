Eres la IA mantenedora de `@aroman22/codegraph-vba`. Repo: `https://github.com/ardelperal/codegraph-vba`. Branch sugerida: `fix/stub-false-post-extraction-resolver`. Versión actual: **1.6.3**.

## Contexto del round

Round 4 = un único gap bloqueante para que un consumer (gestion_riesgos, projectId `00-gestion-riesgos-staging`) pueda usar `codegraph_explore` blast radius con certidumbre sobre calls cross-form. Hoy el VBA extractor emite `1776 edges.kind='calls' con metadata.synthesizedBy='vba-name-resolution' AND metadata.stub=false`: el parser SABE el qualified name (e.g. `Form_FormRiesgosGestion.ActualizarEtiquetaUltimoCambio`, `JsonConverter.ParseJson`) pero igual crea un synthetic function node en vez de linkear al existente. Eso significa: blast radius duplica el caller, call-graph tiene nodos fantasma, y cualquier consumer que filtra "stubs only" para detectar "missing callees" recibe señal ruidosa mezclada con llamadas reales.

Rounds previos en codegraph-vba (NO repetir huecos cerrados):

- **Issue #108 round-3 (recién filed)**: clasificar `unresolved_refs.reference_kind` por forma sintáctica. Round-4 NO toca `unresolved_refs` (esos son los refs no resueltos que NO emitieron edge). Round-4 SÍ toca los `edges` que SÍ emitieron pero al synthetic stub equivocado.
- **Issue #40 (cerrado)**: paren-form pollutes with stubs; statement-form drops cross-module. Cubre el parser/qualifier detection. Round-4 va un paso después: el nombre ya está bien formado como qualified, lo que falta es el repoint post-extracción.
- **Issue #89 (cerrado)**: infer local var type from same-file factory return type. Round-4 reusa el mismo principio (resolver tipos) pero a nivel cross-module / cross-form.

**Lo que YA funciona (NO tocar)** — validado contra `gestion_riesgos` con `@aroman22/codegraph-vba@1.6.3`:

- Schema `edges`:
  ```sql
  CREATE TABLE edges (
      id INTEGER PRIMARY KEY,
      source TEXT, target TEXT,
      kind TEXT,          -- 'calls', 'contains', 'raises-event', 'implemented-by', ...
      metadata TEXT,      -- JSON object: synthesizer info
      line INTEGER, col INTEGER, provenance TEXT
  )
  ```
- VBA extractor emite `kind='calls'` con `metadata.synthesizedBy='vba-name-resolution'` cuando la CALL_RE matchea pero NO resuelve a un nodo existente (e.g. cross-module, cross-form, late-bound runtime como `fso.GetFile`).
- En el `metadata` viene `{synthesizedBy, stub, receiverType, member}`. `stub:false` significa "el parser sabe el nombre qualified pero no lo encontró como node id"; `stub:true` significa "el parser no tiene forma de saber si existe".
- Otros edges (kind=`contains`, `raises-event`, `property-of`, `implemented-by`) operativos, intactos.
- `nodes` table con ids hasheados `function:<hash>`, `class:<hash>`, etc.
- Fixture suite VBA: 167 verde en 1.6.3 (mantener).
- Round-3 issue #108 archivado: documenta el `unresolved_refs.reference_kind` y pide clasificación por forma. Independiente de este round.
- `codegraph_explore` call-path traversal ya funciona sobre `edges.kind='calls'` (verificado en HPS, gestion_riesgos, otros 14 worktrees).

## Lo que falta en este round

### Bug único: post-extraction resolver deja `stub:false` edges sin repointing

#### Síntoma verificado

Consumer corre contra `.codegraph-vba/codegraph.db` (26 MB, codebase `gestion_riesgos`):

```sql
SELECT
  CASE WHEN json_extract(metadata, '$.stub') = 1 THEN 'stub:true (real MISSING callee)'
       ELSE 'stub:false (callee EXISTE, parser no linkeó)'
  END AS bucket,
  COUNT(*) AS n
FROM edges
WHERE kind = 'calls'
  AND json_extract(metadata, '$.synthesizedBy') = 'vba-name-resolution'
GROUP BY bucket ORDER BY n DESC;
```

Devuelve literal:

```json
[
  {"bucket":"stub:true (real MISSING callee)","n":3201},
  {"bucket":"stub:false (callee EXISTE, parser no linkeó)","n":1776}
]
```

De los `1776 stub:false`, **muchos son calls a funciones que claramente existen** en otros lados del codebase. Veinte ejemplos verbatim del consumer:

```sql
SELECT
  json_extract(metadata,'$.receiverType') || '.' ||
  json_extract(metadata,'$.member') AS qualified,
  COUNT(*) AS n
FROM edges
WHERE kind='calls'
  AND json_extract(metadata,'$.synthesizedBy')='vba-name-resolution'
  AND json_extract(metadata,'$.stub')=0
GROUP BY qualified
ORDER BY n DESC LIMIT 20;
```

```json
[
  {"qualified":"Form_FormRiesgosGestion.ActualizarEtiquetaUltimoCambio","n":3},  -- ESTE FORM EXISTE, debería estar linkeado
  {"qualified":"Form_FormRiesgosGestionEdicion.FiltrarPorTipoRiesgo","n":2},  -- idem
  {"qualified":"JsonConverter.ParseJson","n":2},                                -- lib declarada, debería estar linkeado
  {"qualified":"Funciones_Generales.GetSetting","n":2},                        -- módulo existe, debería estar linkeado
  {"qualified":"Constructor.getPC","n":1},
  {"qualified":"modRiesgoEstablecerDatosHelper.SetRiesgo","n":1},              -- módulo existe (Test_/mod_ helper)
  {"qualified":"CorregirErrores.recalcular","n":1},
  {"qualified":"TbConfiguracionBackends.getBackendPath","n":1}                  -- DAO class existe
]
```

En `gestion_riesgos` los nombres `Form_FormRiesgosGestion.ActualizarEtiquetaUltimoCambio`, `Funciones_Generales.GetSetting`, `modRiesgoEstablecerDatosHelper.SetRiesgo`, etc. están todos declarados en su propio módulo/clase correspondiente — pero el post-extraction resolver no los linkea. Resultado:

- Blast radius en `codegraph_explore getCallers(X)` devuelve un set incompleto: incluye solo el stub node, no la declaración real, así que las queries que esperan "todos los callers de X" se cortan.
- `findReferences X` reporta los edges sintéticos como si X fuera un método huérfano.
- Tests downstream (como el lint #69 del consumer) que filtran `metadata.stub=true` esperando solo missing callees reales reciben señal contaminada.

#### Diagnóstico preliminar (NO verificado internamente — confirmar/descartar)

El extractor produce el synthetic stub al final de `scanCallSites` / `emitQualifiedStatementCallEdge` cuando `findFunctionNodeByName` no matchea. Pero el nombre qualified `${receiverType}.${member}` puede matchear un nodo existente con un receiver type derivado distinto al que ya está en `nodes`.

Hipótesis (es un resolver gap, no un parsing gap):

1. Existe o debería existir un "post-extraction pass" que itera `edges` filtradas por `metadata.synthesizedBy='vba-name-resolution'` y trata de repoint cada `target` (synthetic function node) al `nodes.id` real con mismo `qualifiedName`. El pass no existe o no se ejecuta en v1.6.3.
2. Alternativamente, el parsing ya tiene toda la info necesaria y solo falta emitir edges directamente al node id real en lugar de al synthetic; el synthetic se creó porque el resolver no estaba conectado al emisor.
3. La distinción `stub:true` vs `stub:false` sugiere que el maintainer YA clasifica estos stubs por separado internamente — solo faltaría el repointing pass para los `stub:false`.

#### Forma esperada tras el fix

Tras la implementación del resolver pass (o su equivalente):

```sql
-- post-fix esperado:
SELECT
  CASE WHEN json_extract(metadata, '$.stub') = 1 THEN 'stub:true (real MISSING)'
       ELSE 'stub:false'
  END AS bucket,
  COUNT(*) AS n
FROM edges
WHERE kind='calls'
  AND json_extract(metadata,'$.synthesizedBy')='vba-name-resolution'
GROUP BY bucket;
```

Resultado esperado:
- `stub:false`: **debe bajar drásticamente** (idealmente a <100, solo casos extremos tipo receiver-of-runtime-expresión que el resolver no puede inferir). En el codebase `gestion_riesgos`, debería caer de 1776 a <200.
- `stub:true`: debe mantenerse alrededor de 3200 (esos son missing callees reales).

Adicional: los synthetic function nodes correspondientes a edges repointados deben limpiarse (no dejar orphans en `nodes`).

#### Tests RED sugeridos

`__tests__/extraction-vba-stub-resolver.test.ts` (nuevo):

```ts
import { describe, it, expect } from 'vitest';
import { openTestDb } from './helpers/db';
import { extractFromSource } from '../src/extraction/vba-extractor';

describe('extraction-vba: post-extraction stub resolver (round-4)', () => {
  it('repoints cross-module qualified call when target exists elsewhere', async () => {
    const srcA = `Public Sub Helper()\n    Debug.Print "x"\nEnd Sub\n`;
    const srcB = `Public Sub Caller()\n    Helper\nEnd Sub\n`;
    // Helper in srcA is missing the import binding that earlier extraction stages rely on.
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/A.bas', srcA);
    await extractFromSource(db, 'src/modules/B.bas', srcB);
    // After resolver pass: edge from Caller → Helper must NOT have stub:true
    // OR must point to the same node id as Helper's declaration node.
    const edges = db.prepare(`
      SELECT e.target, e.metadata
      FROM edges e
      WHERE e.source IN (
        SELECT id FROM nodes WHERE name = 'Caller'
      ) AND e.kind = 'calls'
    `).all();
    expect(edges.length).toBe(1);
    const targetNode = db.prepare(
      `SELECT id FROM nodes WHERE name = 'Helper' AND kind = 'function'`
    ).get();
    expect(targetNode).toBeDefined();
    expect(edges[0].target).toBe(targetNode.id); // repointed, not a synthetic stub
    expect(JSON.parse(edges[0].metadata).stub ?? null).not.toBe(true);
  });

  it('cross-form method call resolves to .form.txt/.cls sibling', async () => {
    const formSrc = `' Form_Foo\nVERSION 1.0 CLASS\nAttribute VB_Name = "Form_Foo"\nPublic Sub Update()\n    Debug.Print "ok"\nEnd Sub\n`;
    const callerSrc = `Public Sub Caller()\n    Dim f As Object\n    Set f = New Form_Foo\n    f.Update\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/forms/Form_Foo.form.txt', formSrc);
    // Two writes (form.txt + .cls) but only one node should exist for Form_Foo.Update.
    await extractFromSource(db, 'src/modules/C.bas', callerSrc);
    const edges = db.prepare(`
      SELECT e.target FROM edges e
      WHERE e.source IN (SELECT id FROM nodes WHERE name = 'Caller')
        AND e.kind = 'calls'
    `).all();
    expect(edges.length).toBeGreaterThanOrEqual(1);
    const declaredUpdate = db.prepare(
      `SELECT id FROM nodes WHERE name = 'Update' AND qualifiedName LIKE '%Form_Foo%'`
    ).get();
    expect(declaredUpdate).toBeDefined();
    expect(edges[0].target).toBe(declaredUpdate.id);
  });

  it('keeps stub:true when target does NOT exist anywhere', async () => {
    const src = `Public Sub Caller()\n    NonExistentSub(42)\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/Z.bas', src);
    const edges = db.prepare(`
      SELECT metadata FROM edges
      WHERE source IN (SELECT id FROM nodes WHERE name = 'Caller')
        AND kind = 'calls'
    `).all();
    expect(edges.length).toBe(1);
    expect(JSON.parse(edges[0].metadata).stub).toBe(true);
  });

  it('NO regresión: 167 fixture VBA tests siguen pasando', async () => {
    const { runFixture } = await import('./fixtures/vba/run');
    const result = await runFixture();
    expect(result.passed).toBeGreaterThan(150);
  });

  it('NO regresión smoke test: stub:false count baja vs baseline 1776 en consumer', async () => {
    // Este test es opcional, puede correr solo contra el codebase del consumer:
    // `pnpm test stub-resolver-benchmark.test.ts` con env CONSUMER_DB=<path>
    // Verifica que tras el fix, SELECT stub:false cuenta <200 en una DB conocida.
  });
});
```

Smoke test re-ejecutable post-fix (cualquier codebase VBA indexado):

```sql
SELECT
  (SELECT COUNT(*) FROM edges
   WHERE kind='calls'
     AND json_extract(metadata,'$.synthesizedBy')='vba-name-resolution'
     AND json_extract(metadata,'$.stub')=0) AS stub_false,
  (SELECT COUNT(*) FROM edges
   WHERE kind='calls'
     AND json_extract(metadata,'$.synthesizedBy')='vba-name-resolution'
     AND json_extract(metadata,'$.stub')=1) AS stub_true;
-- Antes: stub_false=1776, stub_true=3201 (en gestion_riesgos).
-- Después esperado: stub_false <200.
```

#### Riesgo

- ~16 worktrees consumidores del fleet (`gestion_riesgos`, HPS, repos VBA internos Telefónica). Sus blast-radius results cambian (mejoran) tras el fix — si alguno tenía una heurística basada en el conteo de synthetic stubs como proxy de "missing", esa heurística se rompe. Documentar en CHANGELOG.
- Round-3 issue #108 sigue OPEN. Los dos rounds son independientes pero ambos tocan `edges`/`unresolved_refs` — la tabla `unresolved_refs` round-3 es complementaria (refs que ni siquiera generaron edge), este round-4 es la otra mitad (edges que emitieron a stub equivocado).
- Performance: hoy el extractor emite 4977 stubs. Si el resolver pass itera una vez tras la extracción, agregar <1% a tiempo total.
- Synthetic nodes huérfanos: tras repoint, hay que verificar que la lógica de cleanup no elimine nodos que sí tienen referencias legítimas.

#### Disciplina

- TDD: tests 1-3 antes de tocar el resolver. Test 4 confirma no regresión. Test 5 (smoke benchmark) opcional.
- Conventional commits con scope `vba-resolver` o `extraction-vba`.
- **NO** modificar el parser CALL_RE ni los call-sweep paths (cubierto por issues #40, #54, #44, #45 ya cerrados). El fix es post-extracción.
- **NO** tocar `unresolved_refs` (cubierto por issue #108 round-3).
- Mantener `stub:true` para missing callees reales — es la señal útil para el consumer.
- NO romper `vitest run vba` (167 verde en 1.6.3).
- Mantener backwards compat: las queries que filtran `metadata.synthesizedBy='vba-name-resolution'` siguen funcionando, solo cambia la distribución interna de stub:true/false.

## Acceptance output

- PR con los 5 tests RED → GREEN (4 RED→GREEN, 1 smoke benchmark opcional).
- `CHANGELOG.md`: `extractor(vba): post-extraction resolver repoints stub:false synthetic call edges to existing nodes; stub:true retained for real missing callees (#<issue>)`.
- Bump a **1.7.0** (minor — cambia superficie de blast-radius en consumidores).
- `server-instructions.md` (o equivalente) actualizado: documentar que `metadata.stub=true` significa "real missing" (no falso positivo de runtime binding).

## Quick start

```bash
git clone https://github.com/ardelperal/codegraph-vba
cd codegraph-vba
git checkout -b fix/stub-false-post-extraction-resolver
pnpm install
pnpm test  # baseline 167 verde
pnpm test __tests__/extraction-vba-stub-resolver.test.ts  # 3 verde tras el fix; rojo antes
```

Test repro contra el codebase real (`gestion_riesgos`):

```bash
node --input-type=module -e "
import {DatabaseSync} from 'node:sqlite';
const db = new DatabaseSync('.codegraph-vba/codegraph.db', {readOnly: true});
const r = db.prepare(\`
  SELECT
    CASE WHEN json_extract(metadata, '\$.stub') = 1 THEN 'stub_true' ELSE 'stub_false' END AS b,
    COUNT(*) AS n
  FROM edges WHERE kind='calls'
    AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution'
  GROUP BY b
\`).all();
console.log(JSON.stringify(r, null, 2));
"
# Antes: [{"b":"stub_true","n":3201},{"b":"stub_false","n":1776}]
# Después esperado: stub_false <200, stub_true ~3200
```

## Reinforcement

Regla cross-consumer: **"un consumer que hace blast-radius en `codegraph_explore` espera que cada callee edge apunte al node id real de la declaración, no a un synthetic stub, salvo que el target no exista en el codebase."**

Si tras el fix, en el codebase `gestion_riesgos`, `stub:false` sigue >500, o `stub:true` baja de 3000 (señal de que el resolver está sobre-repinting missing como existentes), escalar a Round 5.

**Anti-pattern explícito:** no marcar `stub:false` para edges que no resolvieron a un nodo declarado. La regla: `stub:false` solo si se confirmó la existencia del target node en algún módulo del codebase indexado.

## Output contract esperado

```json
{
  "tool": "codegraph-vba",
  "mode": "bug-hunt",
  "round": "round-4",
  "variant": "medium",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-codegraph-vba-round-4-2026-07-13.md",
  "prompt_bytes": <size>,
  "verification_queries": [
    "node -e \"const {DatabaseSync}=require('node:sqlite');const db=new DatabaseSync('.codegraph-vba/codegraph.db',{readOnly:true});console.log(JSON.stringify(db.prepare(\\\"SELECT CASE WHEN json_extract(metadata,'\\\\\\$.stub')=1 THEN 'stub_true' ELSE 'stub_false' END AS b, COUNT(*) AS n FROM edges WHERE kind='calls' AND json_extract(metadata,'\\\\\\$.synthesizedBy')='vba-name-resolution' GROUP BY b\\\").all()))\"",
    "pnpm test __tests__/extraction-vba-stub-resolver.test.ts"
  ],
  "cross_session_safe": true
}
```
