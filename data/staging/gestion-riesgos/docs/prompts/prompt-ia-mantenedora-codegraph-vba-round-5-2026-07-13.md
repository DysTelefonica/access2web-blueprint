Eres la IA mantenedora de `@aroman22/codegraph-vba`. Repo: `https://github.com/ardelperal/codegraph-vba`. Branch sugerida: `fix/stub-resolver-both-flags` (sugerida para consolidar con round-4 si querés). Versión actual: **1.6.3**.

## Contexto del round

Round 5 = ampliación de scope del round-4 (`#109`): la evidencia reunida por el consumer `gestion_riesgos` indica que el post-extraction resolver no solo debe cubrir `stub=false` — también `stub=true` requiere el mismo tratamiento. **248/248 qualified names distintos extraídos de edges con `stub=true` YA EXISTEN como `nodes.qualified_name` en el codebase.**

Rounds previos en codegraph-vba (NO repetir huecos cerrados):

- **Issue #108** (round-3): clasificar `unresolved_refs.reference_kind` por forma sintáctica — independiente de este round (cubre la tabla `unresolved_refs`, no `edges`).
- **Issue #109** (round-4): post-extraction resolver que repointa `edges.kind='calls' AND metadata.synthesizedBy='vba-name-resolution' AND metadata.stub=false`. El cuerpo de #109 propone los tests RED y los acceptance criteria pero acotados al bucket `stub=false`.
- **Este round (#X)**: el mismo resolver debe actuar también sobre `metadata.stub=true`. La causa raíz es la MISMA — el extractor crea edges sintéticos sin intentar lookup contra `nodes` antes de decidir el flag `stub`.

## Lo que YA funciona (NO tocar)

- Schema `edges` y `nodes` (round-4 ya los documenta; no los repito).
- VBA extractor (`src/extraction/vba/calls.ts`, `call-sweep.ts`, `controls.ts`, `declarations.ts`, etc.) está intacto y emite los synthetic stubs correctamente.
- Round-3 issue #108 archived + este round está abierto pero independientemente (distinto target).
- Fixture suite VBA: 167 verde en 1.6.3.
- El maintainer YA distingue entre stubs y edge kinds (los kinds del schema `edges` cubren `calls`, `contains`, `raises-event`, etc.). No romper ese surface.

## Lo que falta en este round

### Bug único: post-extraction resolver cubre `stub=false` pero NO `stub=true`

#### Síntoma verificado

Consumer corre esto contra `.codegraph-vba/codegraph.db` del codebase `gestion_riesgos` (26 MB):

```sql
SELECT COUNT(*) AS n
FROM (
  SELECT DISTINCT
    json_extract(metadata, '$.receiverType') || '.' ||
    json_extract(metadata, '$.member') AS q
  FROM edges
  WHERE kind='calls'
    AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
    AND json_extract(metadata, '$.stub')=1
) synthetic
WHERE EXISTS (
  SELECT 1 FROM nodes WHERE qualified_name = synthetic.q
);
```

Devuelve literal: **`n = 248`**.

Y el total de qualified names distintos emitidos como `stub=true` también es **248**. Es decir: **el 100% de los synthetic-stub calls con `stub=true` tienen target declarado en `nodes`** — codegraph-vba nunca los enlazó.

Sample adicional sobre los 12 qualified names más frecuentes:

| `qualified` (stub=true) | n occurrences | ¿Existe en nodes? (exact match) |
|---|---|---|
| `VBA.DoEvents` | 1309 | ✅ (`function:d52a64f1...`) |
| `Collection.Add` | 874 | ✅ |
| `DAO.Execute` | 202 | ✅ |
| `fso.GetFileName` | 68 | ✅ |
| `DAO.OpenRecordset` | 51 | ✅ |
| `fso.FileExists` | 51 | ✅ |
| `ListBox.Column` | 35 | ✅ |
| `riesgo.SetPropiedad` | 33 | ✅ |
| `err.Raise` | 26 | ✅ |
| `ListBox.Selected` | 25 | ✅ |
| `correo.EnviarCorreo` | 19 | ✅ (`function:e918589ec8...`) |
| `ListBox.AddItem` | 18 | ✅ |

Cada uno tiene `nodes.qualified_name = <qualified>` y un nodo `function:` o `class:` que el codegraph-vba no logró enlazar.

#### Consecuencia práctica

El consumer `gestion_riesgos` estuvo a punto de shippear `tools/lint-missing-callees.mjs` como un workaround consumer-side para detectar missing callees. Eso producía `findings=0` porque **TODO está declarado** — solo que el codegraph no lo sabía reportar. El PR se cerró ([PR #117 rolled-back](https://github.com/DysTelefonica/GESTION_RIESGOS/pull/117)) reconociendo que el fix correcto pertenece al extractor, no al consumer.

Tras este round-5 fix, consumer-side lints que filtren por `metadata.stub=true` empezarán a tener señal real: dejarán de reportar `findings=0` (false negative) y empezarán a reportar los typos reales del dominio que quedan.

#### Diagnóstico preliminar (NO verificado internamente — confirmar/descartar)

El extractor produce el synthetic stub al final de `scanCallSites` y/o `emitQualifiedStatementCallEdge` cuando `findFunctionNodeByName` no matchea. Inicialmente marca `metadata.stub=true` como "no sé si existe".

**Hipótesis:** El parser actual NUNCA consulta `nodes` para decidir si flag con `stub=true` o `stub=false` — es un solo camino "stub=true hasta que alguien demuestre lo contrario". Round-4 (#109) introdujo el camino `stub=false` para casos donde se confirma la existencia pre-extracción, pero falta el paso post-extracción que mira `nodes.qualified_name` y repointa cualquier stub (true o false) cuyo target exista.

#### Forma esperada tras el fix

El acceptance criteria ampliado:

```sql
-- post-fix esperado:
SELECT
  (SELECT COUNT(*) FROM (
    SELECT DISTINCT
      json_extract(metadata, '$.receiverType') || '.' ||
      json_extract(metadata, '$.member') AS q
    FROM edges
    WHERE kind='calls'
      AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
    ) synthetic
    WHERE EXISTS (
      SELECT 1 FROM nodes WHERE qualified_name = synthetic.q
    )
  ) AS declared_targets,
  (SELECT COUNT(*) FROM edges
    WHERE kind='calls'
      AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
      AND json_extract(metadata, '$.stub')=1
  ) AS stub_true_count;
-- Antes: declared_targets=248, stub_true_count=3201
-- Después esperado: stub_true_count <500 (solo casos genuinos missing) y todas las
-- edges que apuntan a declared_targets están repointadas (sus edges.target = nodes.id real,
-- no = synthetic function hash).
```

Adicional: la distinción semántica `stub=true` (genuino missing) vs `stub=false` (declared but parser couldn't link pre-extraction) deja de ser útil porque tras el repointing todos los edges declaran correctamente. Mantener el campo `stub=true` SOLO para targets NO encontrados en `nodes`.

#### Tests RED sugeridos

Extender los tests del round-4 (`__tests__/extraction-vba-stub-resolver.test.ts`):

```ts
import { describe, it, expect } from 'vitest';
import { openTestDb } from './helpers/db';
import { extractFromSource } from '../src/extraction/vba-extractor';

describe('extraction-vba: post-extraction stub resolver covers BOTH flags (round-5)', () => {
  // Mismos tests que round-4 — el comportamiento esperado es IDENTICO
  // independientemente del flag stub inicial.

  it('repoints stub:true cross-module call when target exists elsewhere', async () => {
    const srcA = `Public Sub Helper()\n    Debug.Print "x"\nEnd Sub\n`;
    // Call site en modulo B sin import: parser emite synthetic stub=true.
    // Tras resolver: la edge debe repointar al function node real de A.
    const srcB = `Public Sub Caller()\n    Helper\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/A.bas', srcA);
    await extractFromSource(db, 'src/modules/B.bas', srcB);
    const edges = db.prepare(`
      SELECT e.target, e.metadata
      FROM edges e
      WHERE e.source IN (SELECT id FROM nodes WHERE name='Caller')
        AND e.kind='calls'
    `).all();
    expect(edges.length).toBe(1);
    const helper = db.prepare(
      `SELECT id FROM nodes WHERE name='Helper' AND kind='function'`
    ).get();
    expect(helper).toBeDefined();
    expect(edges[0].target).toBe(helper.id); // repointed, NO synthetic stub
    expect(JSON.parse(edges[0].metadata).stub ?? null).not.toBe(true);
  });

  it('keeps stub:true ONLY when target truly does NOT exist anywhere', async () => {
    const src = `Public Sub Caller()\n    DoesNotExistSub(42)\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/Z.bas', src);
    const edges = db.prepare(`
      SELECT metadata FROM edges
      WHERE source IN (SELECT id FROM nodes WHERE name='Caller')
        AND kind='calls'
    `).all();
    expect(edges.length).toBe(1);
    expect(JSON.parse(edges[0].metadata).stub).toBe(true);
  });

  it('NO regresion: 167 fixture VBA tests siguen pasando', async () => {
    const { runFixture } = await import('./fixtures/vba/run');
    const result = await runFixture();
    expect(result.passed).toBeGreaterThan(150);
  });

  it('integration smoke: declared_targets vs stub_true debe colapsar tras el fix', async () => {
    // Smoke benchmark contra el codebase del consumer (si se prueba contra el
    // .codegraph-vba/codegraph.db real del consumer gestion_riesgos).
    // Esperado post-fix: COUNT(declared_targets) == stub_true_count (ambos colapsan
    // cerca de 0 en codebase limpio) y los repointed edges apuntan a nodes.id reales.
    const { DatabaseSync } = await import('node:sqlite');
    // (omitido: path real al .db). Solo el shape del query importa para tests.
  });
});
```

#### Riesgo

- Si el maintainer quiere MANTENER el flag `stub=true` para `false` semantics (e.g. "realmente no estoy seguro"), entonces el fix debe repointar edge pero NO modificar el flag. Decisión del maintainer — la opción que recomiendo es: para targets que existen, repoint edge + flip `stub` a `false`; para targets que no existen, mantener `stub=true`. Documentar la elección en CHANGELOG.
- Round-4 ya cubre el caso `stub=false` pre-extracción. Este round-5 cubre el caso `stub=true` (y por extensión unifica el comportamiento).
- El consumer `gestion_riesgos` ya rolled-backeó su PR #117. No habrá conflicto en el árbol consumer. Solo este upstream PR y uno nuevo tests file.
- Performance: hoy el post-extraction resolver no existe (round-4 lo introduce). Este round-5 lo AMPLÍA para tratar ambos flags. Diff pequeño. Sin regresión.

#### Disciplina

- TDD: los 4 tests RED antes de tocar el resolver.
- Conventional commits con scope `vba-resolver` (consistente con round-4).
- NO modificar la lógica de CANCELACIÓN del flag — solo el repointing.
- Mantener backwards compat: `unresolved_refs` de round-3 + `edges.stub=false` de round-4 + `edges.stub=true` de este round-5 → todos conviven.
- Mantener el campo `metadata` JSON — solo modificamos el campo `stub` (boolean) y `target` (ref a nodes.id) cuando corresponda.
- NO romper `vitest run vba` (167 verde en 1.6.3).

## Acceptance output

- PR con los 4 tests RED → GREEN (los 2-4 de round-4 más este nuevo `declared_targets vs stub_true count`).
- `CHANGELOG.md`: `extractor(vba): post-extraction resolver now repoints BOTH stub:true AND stub=false synthetic call edges to existing nodes (round-5 amplification of #109)` — combine con el changelog de round-4 si se fusionan los PRs.
- Bump a **1.7.0** (minor — ampliando el surface del resolver).
- Smoke benchmark documentado: el query `declared_targets == 0` post-fix contra codebase consumer `gestion_riesgos` debería pasar.
- `server-instructions.md`: aclarar que `stub=true` ahora significa EXCLUSIVAMENTE "target does not exist in `nodes.qualified_name`".

## Quick start

```bash
git clone https://github.com/ardelperal/codegraph-vba
cd codegraph-vba
# Si round-4 (#109) ya está mergeado, basarse en su branch.
# Si no, partir de main y aplicar ambos rounds juntos.
git checkout -b fix/stub-resolver-both-flags
pnpm install
pnpm test  # baseline 167 verde + nuevos tests RED
pnpm test __tests__/extraction-vba-stub-resolver.test.ts
```

Test repro contra codebase real (consumer `gestion_riesgos`):

```bash
node --input-type=module -e "
import {DatabaseSync} from 'node:sqlite';
const db = new DatabaseSync('.codegraph-vba/codegraph.db', {readOnly:true});
const r = db.prepare(\`
  SELECT
    (SELECT COUNT(DISTINCT json_extract(metadata, '\$.receiverType')||'.'||json_extract(metadata, '\$.member'))
       FROM edges WHERE kind='calls'
         AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution') AS distinct_targets,
    (SELECT COUNT(*) FROM edges WHERE kind='calls'
       AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution'
       AND json_extract(metadata, '\$.stub')=1) AS stub_true,
    (SELECT COUNT(*) FROM (
       SELECT DISTINCT json_extract(metadata, '\$.receiverType')||'.'||json_extract(metadata, '\$.member') AS q
       FROM edges WHERE kind='calls'
         AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution'
         AND json_extract(metadata, '\$.stub')=1
    ) s WHERE EXISTS (SELECT 1 FROM nodes WHERE qualified_name = s.q)) AS declared_but_stub_true
\`).all();
console.log(JSON.stringify(r, null, 2));
"
# Antes:  [{"distinct_targets":248,"stub_true":3201,"declared_but_stub_true":248}]
# Después: distinct_targets cerca de 0 AND declared_but_stub_true=0
```

## Reinforcement

Regla cross-consumer: **"un consumer que use el codegraph como fuente de missing callees debe poder confiar en `metadata.stub=true` como `target genuinely does not exist`. Si el codegraph emite `stub=true` para un target que SÍ existe, el guardrail falla silenciosamente (false negative). Este round-5 cierra esa puerta."**

El consumer `gestion_riesgos` rolled-back su PR #117 (workaround) precisamente porque este gap no se arregla consumer-side. El fix correcto es en el extractor.

**Anti-pattern explícito:** no emitir `stub=true` cuando `nodes.qualified_name` ya matchea el target.

## Output contract esperado

```json
{
  "tool": "codegraph-vba",
  "mode": "bug-hunt",
  "round": "round-5",
  "variant": "medium",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-codegraph-vba-round-5-2026-07-13.md",
  "prompt_bytes": <size>,
  "verification_queries": [
    "node -e \"const {DatabaseSync}=require('node:sqlite');const db=new DatabaseSync('.codegraph-vba/codegraph.db',{readOnly:true});console.log(JSON.stringify(db.prepare(\\\"SELECT (SELECT COUNT(*) FROM edges WHERE kind='calls' AND json_extract(metadata,'\\\\\\$.synthesizedBy')='vba-name-resolution' AND json_extract(metadata,'\\\\\\$.stub')=1) AS stub_true, (SELECT COUNT(*) FROM (SELECT DISTINCT json_extract(metadata,'\\\\\\$.receiverType')||'.'||json_extract(metadata,'\\\\\\$.member') AS q FROM edges WHERE kind='calls' AND json_extract(metadata,'\\\\\\$.synthesizedBy')='vba-name-resolution' AND json_extract(metadata,'\\\\\\$.stub')=1) s WHERE EXISTS (SELECT 1 FROM nodes WHERE qualified_name = s.q)) AS declared_but_stub_true\\\").all()))\"",
    "pnpm test __tests__/extraction-vba-stub-resolver.test.ts"
  ],
  "cross_session_safe": true,
  "amplifies_round": "round-4 / issue #109"
}
```
