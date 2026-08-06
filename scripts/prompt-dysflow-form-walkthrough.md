# Prompt para sub-agentes que hacen walkthrough de forms con Dysflow

> **Estado:** workflow rules verificadas contra errores del incidente 2026-08-06 (sub-agentes del walkthrough de NoConformidades).
> **Aplicación:** cualquier task de sub-agente que use Dysflow para walkthrough de forms Access/VBA.

## MUST-LOAD antes de hacer nada

Cargá estas tres skills **antes de llamar la primera tool**:

| Skill | Por qué |
|---|---|
| `dysflow-usage` | Nombres canónicos de tools, write flags, códigos de error. La IA **no debe memorizar tool names** — la skill es la fuente. |
| `dysflow-arnes` | Hard rules (HR-1..HR-14) y anti-patterns (AP-1..AP-12). Incluye HR-14 y AP-12 que son específicamente sobre forms desatendidos. |
| `access-form-ui-builder` | El loop perceive → act → verify completo, con la sección "Forms desatendidos" que cubre el caso bindings vacíos. |

Si tu host tiene auto-discovery de skills, listalas con el mecanismo nativo
y verificá que las tres aparezcan. Si falta alguna, **pará y avisale al humano**
— no improvises.

## Bootstrap (paso 0, no negociable)

```text
get_capabilities({})
```

Eso te devuelve el **catálogo vivo** de tools con sus parámetros canónicos
y el `canonicalCommitFlag` de cada una. Si un nombre de tool no aparece
acá, **no existe**. La evidencia del incidente: la IA llamó `get_schema`
sin chequear y ese tool no existe — el equivalente real es
`query_execute({mode:"read", sql:"..."})` para inspeccionar schema.

Después de capabilities, por cada tool que vayas a usar por primera vez:

```text
describe_tool({name:"<tool>"})
```

Eso te da el schema exacto. La IA del incidente llamó `analyze_form_ui`
con `formName` (parámetro inventado) y `analyze_form_layout` con
`outputMode:"summary"` (modo no soportado por esa tool) — ambos errores
se hubieran evitado con un `describe_tool` previo.

## Loop por cada form

Para cada `.form.txt`:

### 1. Perceive — declarado

```text
analyze_form_ui({sourcePath:"<abs path al .form.txt>", outputMode:"full"})
```

- **Parámetro: `sourcePath`** (path absoluto al `.form.txt`). NO `formName`.
- **Output modes válidos**: `"summary" | "file" | "full"`. Default es `summary`;
  usá `"full"` cuando necesites bindings y eventos detallados por control.
- Devuelve: `controls[]` con `name`, `type`, `role`, `events[]`, `bindings[]`,
  `properties{}`, más `formEvents[]` y `warnings[]`.

### 2. Perceive — geometry

```text
analyze_form_layout({sourcePath:"<abs path>"})
```

- **SÍ acepta `outputMode`** (`"summary" | "file" | "full"`) según el schema
  actual. El prompt inicial decía lo contrario — estaba MAL.
- **KNOWN BUG** (DysTelefonica/dysflow issue #1407, round-2, 2026-08-06): la
  tool falla con `RESULT_CONTRACT_VIOLATION` opaco en TODAS las variantes
  probadas (5 variantes × 3 forms distintos = 15/15 fallan). El envelope
  de error NO emite `code`/`remediation`/`actualShape`/`expectedShape`.
  Mientras el bug esté abierto, **SKIP este step** y registralo en el
  output del form como `layout_status: "skipped_bug_1407"`.
- Devuelve findings tipados de overlap, alignment, off-section, tab-order. **Cuando esté arreglado**, leer el response completo; la tool ya
  devuelve `{findings, controls, sections}` sin filtros.

### 3. Perceive — handler REAL

```text
map_form_behavior({sourcePath:"<abs path>", autoFetchCodeGraph:true, outputMode:"full"})
```

- `autoFetchCodeGraph:true` relaja el boundary no-MCP-to-MCP y deja que
  dysflow consulte codegraph-vba internamente. Si no lo pasás, tenés que
  suministrar vos misma `codegraphEvidence[]`.
- Devuelve: `controls[]` con `codegraphEvidence[]` (handler + callPath +
  tables + effects) por control. Esa es la fuente de verdad del comportamiento.

### 4. Verify — bindings contra schema

```text
query_execute({mode:"read", sql:"SELECT * FROM <Tabla> WHERE 1=0"})  // fan out por tabla
verify_form_bindings({sourcePath:"<abs path>", schema:<aggregate>})
```

- `schema` acepta `Record<tableName, ColumnSchema[]>` (multi-tabla) o
  un payload single-table `{schema:[...], tableName:"..."}`.
- Findings tipados: `FORM_BINDING_MISSING_TABLE`,
  `FORM_BINDING_MISSING_COLUMN`, `FORM_BINDING_EMPTY`,
  `FORM_BINDING_SQL_UNPARSEABLE`, `FORM_BINDING_TYPE_MISMATCH`.
- Severity siempre `warning` (informativo, nunca gatea).

## Si los `bindings[]` vienen vacíos — es formulario desatendido

El `.form.txt` está "limpio" por diseño. NO es bug. NO leas el `.form.txt`
con `Read` a mano. Pasos:

1. Confirma que es desatendido buscando en el `.cls` hermano:
   `Me\.\w+\.(RowSource|ControlSource)\s*=` dentro de `Form_Open` /
   `Form_Load`.
2. Si confirmás, los bindings viven en el `.cls`. Usá
   `map_form_behavior({autoFetchCodeGraph:true})` — eso te da el call
   path real del handler con las tablas/queries que referencia.
3. Después validá con `verify_form_bindings({schema})` aunque ya
   sepas que el binding es por código — el tipado `FORM_BINDING_*`
   te canta columnas renombradas o tablas borradas upstream.
4. Para tracear un control puntual hasta la SQL final, usá la skill
   `vba-handler-backtrace` (no es de dysflow, es del bundle).

Anti-pattern documentado: **`Read` del `.form.txt` para extraer bindings
que `analyze_form_ui` devolvió vacíos**. Ver `dysflow-arnes` HR-14 + AP-12.

## Errores comunes del incidente del 2026-08-06 (no repetir)

| Error | Causa | Fix |
|---|---|---|
| `dysflow.get_schema` no existe | Tool inventada | Usar `query_execute({mode:"read", sql:...})` para inspeccionar schema, después pasar aggregate a `verify_form_bindings` |
| `analyze_form_ui({formName:...})` falla | Parámetro mal | El parámetro canónico es `sourcePath` (path absoluto al `.form.txt`) |
| `analyze_form_layout({outputMode:"summary"})` → result contract violation | Modo no soportado en layout | Layout no acepta `outputMode`. Sacalo. |
| `analyze_form_layout` con `projectId` falla | Layout usa path resolver distinto | Pasá solo `sourcePath` |
| Asumir que `bindings:[]` es bug | Asumir forma "web" en Access | Si el form es desatendido, los bindings se asignan en `Form_Open`/`Form_Load` del `.cls` — IR correcto |
| Leer `.form.txt` a mano cuando `bindings:[]` | Anti-pattern | `map_form_behavior({autoFetchCodeGraph:true})` + `verify_form_bindings` |

## Output esperado por form

Por cada `.form.txt` procesado, devolver:

```text
## Form_<Name>
- sourcePath: <abs path>
- declared controls: <count>, formEvents: <list>
- declared bindings: <count>  // si 0, ver "Forms desatendidos"
- geometry findings: <list por severity>
- handler evidence: <count> codegraphEvidence entries (handler + callPath)
- binding findings: <list FORM_BINDING_*>
- unattended?: <true|false> + nota si lo es
```

## Self-check antes de cerrar la sesión

- ¿Cargué las 3 skills (dysflow-usage, dysflow-arnes, access-form-ui-builder)?
- ¿Llamé `get_capabilities({})` antes de la primera tool?
- ¿Por cada tool nuevo hice `describe_tool`?
- ¿Para cada form usé el loop de 4 tools en orden (ui → layout → behavior → verify_bindings)?
- ¿Si vi `bindings:[]` confirmé si era desatendido antes de investigar más?
- ¿No leí ningún `.form.txt` con `Read` para extraer bindings?

Si la respuesta a cualquiera es NO, volvé al paso correspondiente.
