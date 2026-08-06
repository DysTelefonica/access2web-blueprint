Eres la IA mantenedora de `@aroman22/codegraph-vba`. Repo: `https://github.com/ardelperal/codegraph-vba`. Branch sugerida: `fix/reference-kind-classification`. Versión actual: **1.6.3** (release 2026-07-12).

## Contexto del round

Round 3 = un único gap bloqueante para que el consumer `gestion_riesgos` (projectId `00-gestion-riesgos-staging`) pueda usar `unresolved_refs` como fuente fiable de "missing callees". Hoy `reference_kind` siempre vale `"references"`, así que el consumer no puede filtrar a nivel SQL y la PR #114 (lint) fue cerrada sin merge por ruido.

Rounds previos en codegraph-vba (NO repetir huecos cerrados):

- **Issue #2 / PR (cerrado)**: parser fixes masivos (string-awareness en call-site detection, multi-variable Dim, inline SQL wrappers, comment stripping, qualified statement-form calls, Form/Section filtering, bracketed receivers, With-blocks, bang operator). NO tocaba clasificación de `reference_kind`.
- **Issue #87 (cerrado)**: tag de SQL table refs con dirección `read`/`write`. Prueba de que el maintainer ya clasifica con enum domain-specific en otros frentes.
- **Issues #92-97 (cerrados)**: parse de Dysflow VBA test manifests. Sin relación con este round.

**Lo que YA funciona (NO tocar)** — validado contra `gestion_riesgos` con `@aroman22/codegraph-vba@1.6.3`:

- Tabla `unresolved_refs` existe con schema (en `src/db/schema.sql`):
  ```sql
  CREATE TABLE unresolved_refs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      from_node_id TEXT NOT NULL,
      reference_name TEXT NOT NULL,
      reference_kind TEXT NOT NULL,    -- <-- columna a clasificar en este round
      line INTEGER NOT NULL, col INTEGER NOT NULL,
      candidates TEXT, file_path TEXT NOT NULL DEFAULT '',
      language TEXT NOT NULL DEFAULT 'unknown',
      status TEXT NOT NULL DEFAULT 'pending',
      name_tail TEXT NOT NULL DEFAULT '',
      metadata TEXT,
      FOREIGN KEY (from_node_id) REFERENCES nodes(id) ON DELETE CASCADE
  )
  ```
- Tools `codegraph_explore`, `codegraph_index`, `codegraph_query` operativas.
- VBA extractor (`src/extraction/vba/calls.ts`, `call-sweep.ts`, `controls.ts`, `declarations.ts`, `dims.ts`, etc.) emite edges `kind: 'calls'`, `'raises-event'`, `'contains'`, `'property-of'`, etc. en `edges`. `unresolved_refs` es el cajón de sastre para refs que NO resolvieron a nodo.
- Fixture suite VBA: 167 verde en 1.6.3 (mantener — Test 5 los re-corre).
- `name_tail` ya contiene el último segmento de un nombre qualificado (issue relative). Útil como input, pero NO es la clasificación.
- Schema es TEXT libre para `reference_kind` — **no agregar CHECK constraints**.

## Lo que falta en este round

### Bug único: `unresolved_refs.reference_kind` siempre emite `"references"` para VBA

#### Síntoma verificado

Consumer corre esto contra `.codegraph-vba/codegraph.db` (26 MB, 1426 rows en `unresolved_refs`):

```sql
SELECT DISTINCT reference_kind, COUNT(*) AS n FROM unresolved_refs GROUP BY reference_kind;
```

Devuelve literal:

```json
[{"reference_kind": "references", "n": 1426}]
```

El **100%** cae en el mismo string. Sin distinción entre callsite, field access, property get, bang operator. Top 30 reference_name muestra mezcla de categorías:

```json
[
  {"reference_name":"Name","n":94},            -- form property
  {"reference_name":"Controls","n":85},        -- form collection
  {"reference_name":"AllowEdits","n":74},      -- form property
  {"reference_name":"NombreIcono","n":32},     -- DAO field
  {"reference_name":"EvidenciaUTE","n":24},    -- DAO field
  {"reference_name":"hWnd","n":19},            -- form property
  {"reference_name":"IDContingencia","n":17},  -- DAO field
  {"reference_name":"HayErrorEnRiesgo","n":12},-- ¿call? ¿property? — hoy indistinguible
  {"reference_name":"InsideWidth","n":11}      -- form property
]
```

Tres categorías mezcladas:

| Categoría | Cantidad aprox. | ¿Missing callee real? |
|---|---|---|
| Form properties / collections (`Name`, `hWnd`) | ~207 | NO — ruido Access |
| Miembros cross-class del dominio (`HayErrorEnRiesgo`, `Valoracion`) | ~400-600 | A veces sí, depende del path |
| DAO field accesses (`!IDContingencia`, `NombreIcono`) | ~500-800 | NO — field, no call |

PR #114 cerrada por el reviewer del consumer (`ardelperal`) con veredicto literal:
> *"Cierro esta PR y retiro su rama. La revisión mostró que el reporte todavía no ofrece una señal fiable: mezcla referencias que no son llamadas y genera un volumen de falsos positivos que impide usarlo como guardrail."*

Conclusión: hasta que `reference_kind` no distinga, ningún consumer puede usar `unresolved_refs` como guardrail. Issue #69 sigue OPEN en `gestion_riesgos`.

#### Diagnóstico preliminar (NO verificado internamente — confirmar/descartar)

VBA extractor tiene varios call-sweep paths en `src/extraction/vba/calls.ts` (`scanCallSites`, `detectStatementCall`, `detectQualifiedStatementCall`, `detectWithMemberCall`) — todos emiten edges `kind:'calls'` cuando resuelven, **o nada cuando no**. No clasifican el unresolved.

Hipótesis (es un plumbing gap, no un parsing gap):

1. Hay un path en el extractor VBA que emite `unresolved_refs` con `reference_kind: 'references'` literal, sin propagar la forma sintáctica que produjo la ref.
2. O el populate en `src/db/queries.ts::insertUnresolvedRef` recibe un valor constante hardcoded desde el lado VBA.

El extractor YA sabe (vía CALL_RE + With-normalizer + bang-detector + qualified-statement detector) si la ref era paren call, statement call, property get/set, bang get/set, member-with. **Solo necesita propagar esa info al insert.**

#### Forma esperada de la clasificación (mínimo para que el consumer pueda filtrar; ajustá nombres si hay motivos)

| `reference_kind` | Sintaxis VBA |
|---|---|
| `call` | `Foo(...)`, `Call Foo`, statement-form `Foo arg`, `Mod.Foo args` |
| `qualified-call` | `obj.Foo(...)` donde `obj` es local var o expresión runtime |
| `property-get` | `Me.Name`, `Me.InsideWidth` lectura |
| `property-set` | `obj.Prop = value` |
| `bang-get` | `Me!SubCtl` lectura |
| `bang-set` | `obj!Field = value` |
| `dao-field-get` / `dao-field-set` | `rs!Field`, `rs.Fields("X")` lectura / escritura |
| `unqualified-ident` | Identifier suelto sin `(` después (e.g. `HayErrorEnRiesgo` en `If HayErrorEnRiesgo Then`) |
| `member-with` | `.Member` dentro de bloque `With <receiver>` |

Si separás más fino (p. ej. `event-handler-ref` para `Form_X.Click` declarations), está bien. La tabla es mínimo viable.

#### Tests RED sugeridos

`__tests__/extraction-vba-reference-kind.test.ts` (nuevo):

```ts
import { describe, it, expect } from 'vitest';
import { openTestDb } from './helpers/db';
import { extractFromSource } from '../src/extraction/vba-extractor';

describe('extraction-vba: reference_kind classification', () => {
  it('paren-form call → reference_kind = "call"', async () => {
    const src = `Public Sub Caller()\n    HelperFunction(42)\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/H.bas', src);
    const rows = db.prepare(
      `SELECT reference_kind FROM unresolved_refs WHERE reference_name = 'HelperFunction'`
    ).all();
    expect(rows.length).toBe(1);
    expect(rows[0].reference_kind).toBe('call');
  });

  it('Me.Name property read → "property-get"', async () => {
    const src = `Public Sub Foo()\n    Dim x As String\n    x = Me.Name\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/forms/Form_X.cls', src);
    const rows = db.prepare(
      `SELECT reference_kind FROM unresolved_refs WHERE reference_name = 'Name' AND reference_kind IN ('property-get','property-set')`
    ).all();
    expect(rows.length).toBeGreaterThanOrEqual(1);
    expect(rows[0].reference_kind).toBe('property-get');
  });

  it('Me!Ctl bang read → "bang-get"', async () => {
    const src = `Public Sub Foo()\n    Dim v As Variant\n    v = Me!SubCtl\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/forms/Form_X.cls', src);
    const rows = db.prepare(
      `SELECT reference_kind FROM unresolved_refs WHERE reference_name = 'SubCtl'`
    ).all();
    expect(rows.length).toBe(1);
    expect(rows[0].reference_kind).toBe('bang-get');
  });

  it('rs!Field DAO access → "dao-field-get"', async () => {
    const src = `Public Sub Foo()\n    Dim rs As DAO.Recordset\n    Set rs = CurrentDb.OpenRecordset("SELECT IDContingencia FROM TbX")\n    Dim x As Long\n    x = rs!IDContingencia\nEnd Sub\n`;
    const { db } = await openTestDb();
    await extractFromSource(db, 'src/modules/H.bas', src);
    const rows = db.prepare(
      `SELECT reference_kind FROM unresolved_refs WHERE reference_name = 'IDContingencia'`
    ).all();
    expect(rows.length).toBeGreaterThanOrEqual(1);
    expect(rows[0].reference_kind).toBe('dao-field-get');
  });

  it('NO regresión: fixture VBA suite sigue pasando (vitest run vba)', async () => {
    const { runFixture } = await import('./fixtures/vba/run');
    const result = await runFixture();
    expect(result.passed).toBeGreaterThan(150); // ~167 en 1.6.3
  });
});
```

Smoke test en cualquier consumer indexado:

```sql
SELECT reference_kind, COUNT(*) AS n
FROM unresolved_refs
GROUP BY reference_kind
HAVING n > 0 ORDER BY n DESC;
-- Antes: 1 fila con kind="references".
-- Después: ≥3 filas con kinds distintos (call + property-get + al menos uno más).
```

#### Riesgo

- ~16 worktrees consumidores del fleet (`gestion_riesgos`, HPS, repos VBA internos Telefónica) que escriben heurísticas `NOISE_PATTERNS` o `_filtered reads` necesitarán migrar a `WHERE reference_kind IN (...)`. **Mantener retro-compat**: el valor legacy `"references"` debe seguir siendo emitido por paths legados durante un release cycle. NO eliminarlo en este round.
- Performance: hoy el extractor escribe 1426 rows. Si la clasificación requiere más parsing, delta de runtime <5%.
- Tests fixture VBA (167 verde) deben seguir verdes sin modificar.

#### Disciplina

- TDD: tests 1-4 antes de tocar el extractor.
- Conventional commits con scope `vba-extract` o `extraction-vba`.
- NO eliminar el valor `"references"` como salida emitida — mantener retro-compat. Paths NUEVOS deben clasificar; los legados pueden seguir emitiendo `"references"` mientras el conteo global BAJE.
- NO cambiar schema de `unresolved_refs`. `reference_kind` ya es TEXT libre.
- NO romper `vitest run vba` (167 verde en 1.6.3). Si tu clasificación toca `calls.ts`/`call-sweep.ts`, re-corre y verifica.

## Acceptance output

- PR con los 5 tests RED → GREEN.
- `CHANGELOG.md`: `extractor(vba): classify unresolved_refs.reference_kind by syntactic shape; legacy "references" retained for back-compat (#<issue>)`.
- Bump a **1.7.0** (minor — cambia surface de clasificación).
- `server-instructions.md` (o equivalente) actualizado con tabla de kinds.

## Quick start

```bash
git clone https://github.com/ardelperal/codegraph-vba
cd codegraph-vba
git checkout -b fix/reference-kind-classification
pnpm install
pnpm test  # baseline 167 verde en 1.6.3
pnpm test __tests__/extraction-vba-reference-kind.test.ts  # 5 verde tras el fix; rojo antes
```

Test repro contra codebase real (consumer `gestion_riesgos`):

```bash
sqlite3 .codegraph-vba/codegraph.db \
  "SELECT reference_kind, COUNT(*) AS n FROM unresolved_refs GROUP BY reference_kind ORDER BY n DESC"
# Actual antes: [{"reference_kind":"references","n":1426}]
# Esperado tras fix: ≥3 filas con kinds distintos.
```

## Reinforcement

Regla cross-consumer: **"un consumer que necesita detectar 'missing callees' puede confiar en `WHERE reference_kind IN ('call','qualified-call','unqualified-ident','member-with','bang-call')` y obtener una lista con <10% de falsos positivos."** Si en el codebase `gestion_riesgos` (1426 rows) la lista filtrada supera ese umbral, escalar a Round 4.

**Anti-pattern explícito:** no emitir `"references"` desde paths nuevos. La regla: paths nuevos clasifican; el conteo de `"references"` debe BAJAR, no quedar congelado ni crecer.

## Output contract esperado

```json
{
  "tool": "codegraph-vba",
  "mode": "bug-hunt",
  "round": "round-3",
  "variant": "medium",
  "prompt_path": "docs/prompts/prompt-ia-mantenedora-codegraph-vba-round-3-2026-07-13.md",
  "prompt_bytes": <size>,
  "verification_queries": [
    "sqlite3 .codegraph-vba/codegraph.db \"SELECT DISTINCT reference_kind, COUNT(*) AS n FROM unresolved_refs GROUP BY reference_kind\"",
    "pnpm test __tests__/extraction-vba-reference-kind.test.ts"
  ],
  "cross_session_safe": true
}
```
