# Audit — codegraph-vba v1.6.3 resolution rate en gestion_riesgos (2026-07-13)

**Autor:** consumer `gestion_riesgos` (issue #69 + #109).
**Repo source:** `DysTelefonica/GESTION_RIESGOS` rama `staging` SHA `72325a1`.
**Codegraph DB:** `.codegraph-vba/codegraph.db` (26 MB) generada por `@aroman22/codegraph-vba@1.6.3`.
**Generado por:** script Node standalone (ver bloque "Reproducibilidad" abajo).

## TL;DR

**Resolution rate = 62.8%.** De los 635 distinct qualified names que el codegraph emite como synthetic stubs (callees no enlazados estáticamente), 399 tienen su target declarado en `nodes` — es decir, codegraph NO los enlazó. Es exactamente el gap descrito en `ardelperal/codegraph-vba#109 round-4` y `#110 round-5`.

Después del fix del post-extraction resolver (#110), se espera que la resolution rate suba a ~100% (~635 distinct callees repointados al nodo real, solo el subset genuino-missing queda como `stub=true`).

## Métricas raw

| Métrica | Valor |
|---|---|
| Nodes totales (`function`/`method`/`property`/`class`) | 9,600 |
| Edges totales | 27,290 |
| Edges `kind='calls'` | 11,340 |
| Synthetic calls (vba-name-resolution) | 4,977 |
| └ stub=true (codegraph dice "no sé") | 3,201 |
| └ stub=false (parser "falló al linkear") | 1,776 |
| Distinct qualified names emitidos como stub | 635 |
| └ con `stub=true` | 248 |
| └ con `stub=false` | 387 |
| └ **target EXISTE en nodes** (resoluble) | **399 (62.8%)** |
| └ └ entre los 248 stub=true | 248 (100%) |
| └ └ entre los 387 stub=false | 151 (39%) |

## Distribución de receiver types (top 15)

| Receiver | n edges sintéticos |
|---|---|
| VBA | 1,365 |
| Collection | 874 |
| Test_Helper | 427 |
| Constructor | 364 |
| Test_Fixtures | 334 |
| DAO | 255 |
| fso | 150 |
| modFormCoordinationHelper | 100 |
| ListBox | 92 |
| riesgo | 85 |
| Edicion | 50 |
| SuministradoresHelper | 34 |
| p_db | 33 |
| RiesgoMaterializacion | 31 |
| err | 27 |

## Top unresolved callees NO-runtime (los que importan para el filtro del consumer)

Estos son los que NO son late-bound runtime (`VBA.*`, `DAO.*`, `Collection.*`, `ListBox.*`, etc.) — son los que un filtro consumidor estricto consideraría "reales" tras descartar runtime noise.

| Qualified | stub | n occurrences |
|---|---|---:|
| `Test_Helper.ForceLocalBackend` | false | 186 |
| `Test_Fixtures.GetTestDb` | false | 169 |
| `Test_Fixtures.SeedAll` | false | 76 |
| `Test_Helper.ResetTestSession` | false | 83 |
| `Test_Helper.BuildJsonFail` | false | 60 |
| `Test_Helper.BuildJsonOk` | false | 60 |
| `Test_Fixtures.TeardownAll` | false | 60 |
| `Constructor.getEdicion` | false | 49 |
| `Constructor.getRiesgo` | false | 46 |
| `Constructor.getProyecto` | false | 24 |
| `riesgo.SetPropiedad` | true | 33 |
| `m_ObjEntorno.ColProyectosTotales` | (mix) | 9 |
| `CORREO.Registrar` | (mix) | 9 |
| `correo.EnviarCorreo` | true | 19 |

**Notable**: la mayoría de los "unresolved del proyecto" están en `Test_Helper` y `Test_Fixtures` (TDD infrastructure). Si round-5/#110 repointa correctamente los callables del proyecto, los residuos deberían ser: (a) runtime APIs, (b) `Test_*` helpers que el parser no indexa porque están en `.bas` test modules, (c) genuine typos.

## Findings accionables para el maintainer de codegraph-vba

### Finding 1 (round-5 #110 support): 248/248 stub=true con target en nodes

Confirmado con query reproducible. Round-5 acceptance criteria debería colapsar este número. Smoke benchmark:

```sql
-- Antes del fix (1.6.3 actual):
SELECT COUNT(*) FROM (
  SELECT DISTINCT
    json_extract(metadata, '$.receiverType') || '.' ||
    json_extract(metadata, '$.member') AS q
  FROM edges
  WHERE kind='calls'
    AND json_extract(metadata, '$.synthesizedBy')='vba-name-resolution'
    AND json_extract(metadata, '$.stub')=1
) s WHERE EXISTS (SELECT 1 FROM nodes WHERE qualified_name = s.q);
-- resultado: 248
-- Esperado post-fix: 0
```

### Finding 2: 151 de los 387 stub=false también tienen target en nodes

El resolver actual maneja `stub=false` (round-4) pero parece no repointar completamente. 151/387 = 39% de los stub=false tienen target existente pero el post-extraction resolver no los conecta. Round-4/#109 cubre stub=false pero la implementación actual parece incompleta en este consumer.

Recomendación al maintainer: round-5/#110 debe unificar ambos flags. La distinción semántica `stub=true` vs `stub=false` pierde sentido si ambos buckets quedan unrepointed.

### Finding 3: case-sensitivity handling

Notar que el receiver `riesgo` (lowercase) está bien caseado contra el nodo declarado `riesgo.SetPropiedad` (también lowercase, porque es la convención del proyecto de nombres de clase en lowercase para entities). Esto sugiere que **el lookup nodes.qualified_name contra synthetic qualified es case-sensitive** y matcha solo si el caso coincide. Si el consumer tiene case mismatches (p. ej. `Riesgo` vs `riesgo`), los counts pueden ser mayores.

Verificación rápida: SQL `LIKE` es case-insensitive para ASCII por default, pero `=` es case-sensitive. El resolver debería usar `LIKE` o normalizar el caso antes de comparar.

### Finding 4: Test_Helper y Test_Fixtures en top receivers

Los tests `Test_*.bas` usan receivers `Test_Helper.ForceLocalBackend`, `Test_Fixtures.SeedAll`, etc. Estos helpers SÍ están declarados en `modTestHelper.bas` / `modTestFixtures.bas`, pero los `Test_*.bas` no se indexan como nodos. Por eso codegraph marca las llamadas a estos como "unresolved". **El extractor's failure to index `Test_*.bas` como nodos es un gap separado** (no cubierto por round-4 ni round-5).

Recomendación: un round-6 sobre la indexación de test modules. O alternativamente, documentar como limitación aceptable: los consumer-side lints pueden filtrar receivers que empiezan con `Test_` como known-noise.

## Reproducibilidad

```bash
# Desde el repo de gestion_riesgos, con el codegraph DB actualizado:
node -e "
const {DatabaseSync}=require('node:sqlite');
const db=new DatabaseSync('.codegraph-vba/codegraph.db',{readOnly:true});
const r=db.prepare(\`
  WITH synth AS (
    SELECT DISTINCT
      json_extract(metadata, '\$.receiverType') || '.' ||
      json_extract(metadata, '\$.member') AS q,
      json_extract(metadata, '\$.stub') AS stub
    FROM edges
    WHERE kind='calls'
      AND json_extract(metadata, '\$.synthesizedBy')='vba-name-resolution'
  )
  SELECT
    COUNT(DISTINCT q) AS distinct_qualnames,
    SUM(CASE WHEN stub=1 THEN 1 ELSE 0 END) AS stub_true_count,
    SUM(CASE WHEN stub=1 AND EXISTS (SELECT 1 FROM nodes WHERE qualified_name=q) THEN 1 ELSE 0 END) AS stub_true_resolvable
  FROM synth
\`).get();
console.log(r);
"
```

## Limitaciones

- Las queries cuentan ocurrencias en metadata; un mismo callee en 5 líneas aparece 5 veces (correcto: cada call site es distinto).
- El conteo "distinct_qualnames" es por distinct (q, stub), no por distinct (q), porque semanticamente stub=true y stub=false son diferentes categorías.
- El extractor podría emitir casos donde `receiverType` es vacío (early scan failures); esos están excluidos de las listas porque no producen qualified name.
- No es exhaustivo del codebase — solo refleja el state actual del codegraph DB (que es coherente con el state al momento de sync).

Refs: gestion_riesgos#69 + ardelperal/codegraph-vba#109 + #110
