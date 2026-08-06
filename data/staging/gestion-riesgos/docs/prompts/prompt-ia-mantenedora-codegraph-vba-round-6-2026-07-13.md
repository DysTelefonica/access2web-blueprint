Eres la IA mantenedora de `@aroman22/codegraph-vba`. Repo: `https://github.com/ardelperal/codegraph-vba`. Branch sugerida: `fix/index-test-bmodules-as-nodes`. Versión actual: **1.6.3**.

## Contexto del round

Round 6 = un único gap NUEVO no cubierto por #108/#109/#110. El consumer `gestion_riesgos` midió resolución 62.8% en codegraph-vba@1.6.3. De los 635 callees sintéticos, **el top-3 de receivers más frecuentes NO son runtime APIs sino infraestructura de tests del propio proyecto**:

```
Test_Helper       : 427 edges sintéticos
Test_Fixtures     : 334 edges sintéticos
Constructor       : 364 edges sintéticos
```

**Codegraph-vba@1.6.3 NO está indexando `Test_*.bas` ni `modTestHelper.bas` / `modTestFixtures.bas` como nodos.** Por eso las llamadas a `Test_Fixtures.SeedAll`, `Test_Helper.ForceLocalBackend`, `Test_Helper.BuildJsonOk`, etc. aparecen como unresolved cuando SÍ están declaradas en el codebase (en archivos que el extractor ignora).

Rounds previos (NO repetir):
- **#108** (round-3): clasificar `unresolved_refs.reference_kind` por forma sintáctica.
- **#109** (round-4): post-extraction resolver repointa `stub:false` synthetic call edges.
- **#110** (round-5): extender round-4 al bucket `stub:true` también.
- **Este round (#X = round-6)**: **indexar `Test_*.bas` y módulos `modTestHelper`/`modTestFixtures` como nodos** para que las llamadas a sus funciones aparezcan como edges a nodos reales en lugar de synthetic stubs.

## Lo que YA funciona (NO tocar)

- VBA extractor (`src/extraction/vba/calls.ts`, `call-sweep.ts`, `declarations.ts`, etc.) está intacto y emite los synthetic stubs correctamente.
- `Test_*` SÍ se procesa parcialmente (las edges se generan). El gap es que los **nodos** correspondientes no se crean.
- Round-3/4/5 mantienen su scope (no tocarlos aquí).
- Fixture suite VBA: 167 verde en 1.6.3.

## Lo que falta en este round

### Bug único: `Test_*.bas` no produce function nodes en `nodes`

#### Síntoma verificado

Consumer corre esto contra `.codegraph-vba/codegraph.db` del codebase `gestion_riesgos` (26 MB):

```sql
SELECT
  json_extract(metadata, '$.receiverType') AS recv,
  COUNT(*) AS n
FROM edges
WHERE kind='calls'
  AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
  AND recv LIKE 'Test_%'
GROUP BY recv
ORDER BY n DESC
LIMIT 10;
```

Devuelve literal:

```
recv                  n
Test_Helper           427
Test_Fixtures         334
                     (resto)
```

Y verificar que NO existen nodos para esos symbols:

```sql
SELECT
  json_extract(metadata, '$.receiverType') || '.' ||
  json_extract(metadata, '$.member') AS q,
  COUNT(*) AS unresolved
FROM edges
WHERE kind='calls'
  AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
  AND json_extract(metadata, '$.receiverType') LIKE 'Test_%'
GROUP BY q
HAVING unresolved > 5
ORDER BY unresolved DESC;
-- resultado: Test_Helper.ForceLocalBackend (186), Test_Fixtures.SeedAll (76),
--           Test_Helper.ResetTestSession (83), Test_Helper.BuildJsonOk (60),
--           Test_Fixtures.TeardownAll (60), Test_Helper.BuildJsonFail (60),
--           Test_Fixtures.GetTestDb (169).
```

Pero `Test_Helper.bas` y `Test_Fixtures.bas` SÍ existen en `src/modules/` del codebase y declaran estas funciones con `Public Sub ...`:

```bash
$ grep -c "^Public Function\|^Public Sub" src/modules/modTestHelper.bas
53
$ grep -c "^Public Function\|^Public Sub" src/modules/modTestFixtures.bas
47
```

**El extractor produce edges (que llaman a esos symbols) pero no crea los nodes que esos symbols representan.** Como consecuencia:
- Las edges sintéticas que llaman a `Test_Helper.ForceLocalBackend`, `Test_Helper.BuildJsonOk`, `Test_Fixtures.GetTestDb`, etc. quedan con `metadata.stub=true` (porque la lookup contra `nodes` falla — `nodes` no tiene esos symbols).
- Tras round-5/#110 esos NO van a repointarse (el target nunca estuvo en nodes).
- Falso negativo en cualquier lint consumer-side que filtre `stub=true` esperando genuinos missing.

#### Consecuencia práctica

El codebase consumer `gestion_riesgos` tiene ~760 edges sintéticos hacia `Test_*` que **deberían ser reales**. El filtro del consumer tendría que excluirlos manualmente con un "ignore list" para evitar 760 falsos positivos.

#### Diagnóstico preliminar (NO verificado internamente — confirmar/descartar)

El VBA extractor probablemente tiene un check en el path de producción que filtra archivos que matchean `Test_*` (o `Test_*` en particular) para evitar ruido de tests en herramientas típicas. Ese mismo check está creando un ciego para `modTestHelper.bas` / `modTestFixtures.bas` (que NO son tests per se, son helpers invocados POR tests).

Hipótesis: una config (probablemente en `src/extraction/vba/config.ts` o equivalente) tiene `excludePatterns: [/^Test_/]`. Esa config está siendo demasiado agresiva: debería excluir SOLO el patrón `/^Test.*\\.bas$/` o `/^Test_.*\\.bas$/` que son los test atoms, no `Test_Helper.bas` / `Test_Fixtures.bas`.

#### Forma esperada tras el fix

Acceptance criteria (medible):

```sql
-- Antes (1.6.3 actual):
SELECT
  (SELECT COUNT(*) FROM nodes WHERE qualified_name LIKE 'Test_Helper.%' OR qualified_name LIKE 'Test_Fixtures.%') AS helper_nodes,
  (SELECT COUNT(*) FROM edges
   WHERE kind='calls'
     AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
     AND (json_extract(metadata, '$.receiverType') LIKE 'Test_Helper%' OR json_extract(metadata, '$.receiverType') LIKE 'Test_Fixtures%')
     AND json_extract(metadata, '$.stub')=1
  ) AS helper_stub_true;
-- esperado antes: helper_nodes=0, helper_stub_true>500

-- Después (post-fix):
-- helper_nodes > 100 (cada Public Function/Sub en los helpers indexado)
-- helper_stub_true < 50 (solo casos genuinos como llamadas via dynamic dispatch)
```

#### Tests RED sugeridos

`__tests__/extraction-vba-test-helper-index.test.ts` (nuevo):

```ts
import { describe, it, expect } from 'vitest';
import { openTestDb } from './helpers/db';
import { extractFromSource } from '../src/extraction/vba-extractor';

describe('extraction-vba: Test_*.bas + modTestHelper/Fixtures indexing (round-6)', () => {
  it('indexes Public Sub declarations from a Test_Helper-shaped module', async () => {
    const src = `Attribute VB_Name = "Test_Helper"
Option Compare Database
Option Explicit

Public Function GetTestDb(Optional ByRef p_Error As String) As String
    GetTestDb = "fake://test"
End Function

Public Sub ForceLocalBackend(ByVal p_URL As String, ByRef p_Error As String)
    ' implementation
End Sub
`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/Test_Helper.bas', src);
    const f = db.prepare(`SELECT id FROM nodes WHERE qualified_name='Test_Helper.GetTestDb'`).get();
    expect(f).toBeDefined();
    const s = db.prepare(`SELECT id FROM nodes WHERE qualified_name='Test_Helper.ForceLocalBackend'`).get();
    expect(s).toBeDefined();
  });

  it('emits calls edges with synthetic stub=false when called from a Test atom', async () => {
    const src = `Attribute VB_Name = "Test_X"
Option Compare Database
Public Sub Test_Stuff()
    Test_Helper.ForceLocalBackend "fake://", ""
End Sub
`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/Test_X.bas', src);
    const edges = db.prepare(`
      SELECT e.target, e.metadata
      FROM edges e
      JOIN nodes n ON n.id = e.source
      WHERE n.name='Test_Stuff' AND e.kind='calls'
    `).all();
    expect(edges.length).toBe(1);
    // Si el round-5/#110 también está merged, el target debe ser el nodo real.
    const target = db.prepare(`SELECT id FROM nodes WHERE qualified_name='Test_Helper.ForceLocalBackend'`).get();
    expect(target).toBeDefined();
    expect(edges[0].target).toBe(target.id);
  });

  it('NO regresion: 167 fixture VBA tests siguen pasando', async () => {
    const { runFixture } = await import('./fixtures/vba/run');
    const result = await runFixture();
    expect(result.passed).toBeGreaterThan(150);
  });

  it('smoke benchmark contra el consumer gestion_riesgos', async () => {
    // (omitido: el path real al .db). Shape del query:
    // 1. helper_nodes > 0 (Test_Helper.* y Test_Fixtures.* indexados)
    // 2. helper_stub_true < 50 (los genuinos missing restantes)
  });
});
```

#### Riesgo

- Si el exclude patterns actual es **global** y rompe otros casos, el fix debe ser surgical — solo permitir `Test_*.bas` real (los átomos TDD), no `Test_Helper.bas` / `Test_Fixtures.bas` (módulos de helpers).
- Para consumers que SÍ quieren ignorar los test atoms `Test_*` (porque atom-by-atom es muy verboso), el fix podría ofrecer una opción `--include-test-helpers` (default ON) configurable vía `.codegraph-vba/config.json`.
- Otros lenguajes del upstream (TypeScript extractor) pueden tener problemas similares (nodes para `./__tests__/*.test.ts`). Si arreglamos VBA, considere extender al extractor universal — o no, depende del scope del round-6.

#### Disciplina

- TDD: 4 tests antes de tocar el extractor.
- Conventional commits con scope `vba-extract` o `extractor-config`.
- Mantener backwards compat: `Test_*.bas` (átomos) siguen excluido; `modTest*.bas` (helpers) ahora se incluye.
- NO tocar round-3/4/5 (#108/#109/#110).
- Mantener `vitest run vba` verde (167 verde en 1.6.3).

## Acceptance output

- PR con los 4 tests RED → GREEN.
- `CHANGELOG.md`: `extractor(vba): index test helper modules (Test_Helper.bas, Test_Fixtures.bas) as nodes; test atoms (Test_*.bas) remain excluded as before. Round-6 of consumer gestion_riesgos#69 (#X)`.
- Bump a **1.7.0** (minor — feature opt-in expansion; sigue cubriendo el mismo dominio).
- Actualizar `server-instructions.md`: documentar que `Test_Helper` / `Test_Fixtures` ahora se indexan como nodos.

## Quick start

```bash
git clone https://github.com/ardelperal/codegraph-vba
cd codegraph-vba
git checkout -b fix/index-test-helper-modules
pnpm install
pnpm test  # baseline 167 verde
pnpm test __tests__/extraction-vba-test-helper-index.test.ts
```

Test repro contra codebase real (`gestion_riesgos`):

```bash
node --input-type=module -e "
import {DatabaseSync} from 'node:sqlite';
const db = new DatabaseSync('.codegraph-vba/codegraph.db', {readOnly:true});
const r = db.prepare(\`
  SELECT
    (SELECT COUNT(*) FROM nodes
     WHERE qualified_name LIKE 'Test_Helper.%' OR qualified_name LIKE 'Test_Fixtures.%') AS helper_nodes,
    (SELECT COUNT(*) FROM edges
     WHERE kind='calls'
       AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution'
       AND (json_extract(metadata, '\$.receiverType') LIKE 'Test_Helper%'
            OR json_extract(metadata, '\$.receiverType') LIKE 'Test_Fixtures%')
       AND json_extract(metadata, '\$.stub')=1) AS helper_stub_true
\`).get();
console.log(r);
"
# Antes: helper_nodes=0, helper_stub_true>500
# Después: helper_nodes>100, helper_stub_true<50
```

## Reinforcement

Regla cross-consumer: **"un codebase TDD-first debería poder confiar en que las llamadas a sus test-helpers estén indexadas, no clasificadas como unresolved."** Si un consumer tiene un manifest TDD estricto, los helpers de test son infraestructura de primera clase, no ruido.

**Anti-pattern explícito:** no filtrar archivos cuyo nombre empieza con `Test_*` como categoría global. Filtrar solo por patrón completo (`Test_*.bas` para átomos) o por convención explícita (e.g., `modTest*.bas` para helpers).

## Output contract esperado

```json
{
  "tool": "codegraph-vba",
  "mode": "bug-hunt",
  "round": "round-6",
  "variant": "medium",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-codegraph-vba-round-6-2026-07-13.md",
  "prompt_bytes": <size>,
  "verification_queries": [
    "node -e \"const {DatabaseSync}=require('node:sqlite');const db=new DatabaseSync('.codegraph-vba/codegraph.db',{readOnly:true});const r=db.prepare(\\\"SELECT (SELECT COUNT(*) FROM nodes WHERE qualified_name LIKE 'Test_Helper.%' OR qualified_name LIKE 'Test_Fixtures.%') AS helper_nodes, (SELECT COUNT(*) FROM edges WHERE kind='calls' AND json_extract(metadata,'\\\\\\$.synthesizedBy')='vba-name-resolution' AND (json_extract(metadata,'\\\\\\$.receiverType') LIKE 'Test_Helper%' OR json_extract(metadata,'\\\\\\$.receiverType') LIKE 'Test_Fixtures%') AND json_extract(metadata,'\\\\\\$.stub')=1) AS helper_stub_true\\\").get();console.log(r);\"",
    "pnpm test __tests__/extraction-vba-test-helper-index.test.ts"
  ],
  "cross_session_safe": true,
  "amplifies_rounds": ["round-3 (issue #108)", "round-4 (#109)", "round-5 (#110)"]
}
```
