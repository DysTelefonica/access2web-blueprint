# Prompt para IA mantenedora de dysflow MCP — Round 1 (lint gaps)

Eres la IA mantenedora de dysflow MCP. Repo: `<repo path>`. Branch sugerida: `fix/lint-gaps-round-1`. Versión actual: 2.1.7 (verificado por el consumer via `dysflow.get_capabilities` 2026-07-09).

## Contexto del round

Round 1 = tres gaps del linter detectados por el consumer `gestion_riesgos` (projectId `00-gestion-riesgos-staging`) durante la implementación del cluster de issues #64/#66/#67/#68 (correo de rechazo de propuesta de publicación). Cada gap bloqueó una compilación real en Access VBE durante 30+ minutos de debugging manual.

No hay rounds previos del lint (round 1 del lint específicamente). Rounds previos del tool (no del lint) que deben respetarse:
- Round tooling capabilities: CONFIG_TOP_LEVEL_FIELDS_REMOVED (resuelto).
- Round security: correcciones + scope (resuelto).

## Lo que YA funciona (NO tocar)

Validado contra el consumer `gestion_riesgos` con 28/28 átomos verdes + 4/4 regresión:

- `lint_module` corre via `dysflow.lint_module({projectId, module, source?})`. Shape actual del reporte: `{lintResult, module, results: [{line, severity, message, ruleId}]}`.
- Reglas actuales operativas: `option-declaration`, `identifier-safety`, `declaration-order`, `arg-type-match`, `forbidden-name (F22)`, `logical-short-circuit`, `implicit-variant`, `missing-exit-handler`, `invalid-static-class-call`.
- `lint_form_code` cubre forms (.form.txt + .cls) con reglas equivalentes.
- `forbidden-name (F22)` flag identifiers que shadow VBA / Access / DAO globals: `Err, Date, Name, Form, DoCmd, etc.` (case-insensitive). **Limitación actual**: solo cubre estos globals nombrados explícitamente. NO cubre TODAS las VBA reserved words (palabras clave del lenguaje como `Enum`, `Function`, `Sub`, `If`, `Then`, `Else`, `Dim`, etc.).
- `humanCompilePending:false` se respeta; los consumers compilan manualmente.
- Capabilities block migrado (`capabilities.allowWrites`, `effectiveDryRunDefault`, etc.). NO romper el schema.
- 64 tools visibles, `writeClassToolsPermitted` correcto, `dryRunDefault:true` por default.
- Backwards compatibility con el consumer fleet (gestion_riesgos + hps + 15+ worktrees).

## Lo que falta en este round

### Bug 1: `Dim x As Type: y = expr` con `:` después de Dim no se detecta

#### Síntoma verificado

El consumer escribió en `Test_Issue64Cluster.bas` (test file nuevo para los fixes #64/#66/#67/#68):

```vba
Dim eNum As Long: eNum = err.Number
Dim eDesc As String: eDesc = err.description
```

Access reporta **`Error de sintaxis`** en el bloque `EH_Seed:` (después del handler de error). El cursor se posiciona en la línea `Dim eNum As Long: eNum = err.Number`.

#### Evidencia de repro

Disparo del consumer en Access VBE:

```
Compile VBA Project
→ Error de compilación:
   "Error de sintaxis"
   [cursor en línea 107 de Test_Issue64Cluster.bas]
```

Captura visual del popup en Access: el cursor se posiciona sobre `Dim eNum As Long: eNum = err.Number` (línea 107).

#### Diagnóstico preliminar (verificado)

En VBA, el carácter `:` se usa como separador de **statements ejecutables** dentro de una misma línea. NO se puede usar para combinar una **declaración** (`Dim`) con otra statement. El statement `Dim eNum As Long` debe ir solo en su línea; el `eNum = err.Number` debe ir en una línea separada.

El linter NO detecta este anti-pattern. La regla `declaration-order` (existente) valida el ORDEN de declaraciones pero no la FORMA sintáctica de cada una.

#### Riesgo

Consumers que escriben fixtures de test con handlers `On Error GoTo <label>` que necesitan guardar `err.Number` en variables locales caen en este anti-pattern. El tiempo de debugging es alto (~30 min en el consumer) porque:
1. El cursor apunta a la línea errónea (correcto)
2. Pero el mensaje genérico "Error de sintaxis" no dice cuál es el problema específico
3. Requiere conocimiento experto de VBA para identificar `:`-after-`Dim`

#### Tests RED sugeridos

Test 1 (caso positivo — debe pasar después del fix):

```ts
it('accepts Dim on its own line followed by assignment', async () => {
  const source = 'Sub Test()\nDim x As Long\nx = 5\nEnd Sub';
  const result = await client.call('lint_module', {
    projectId: '00-gestion-riesgos-staging',
    module: 'TestModule',
    source
  });
  const dimColonErrors = result.results.filter(r =>
    r.message.includes('colon after Dim') || r.message.includes('Dim')
  );
  expect(dimColonErrors.length).toBe(0);
});
```

Test 2 (caso negativo — debe detectar el anti-pattern):

```ts
it('flags Dim x As Type: y = expr with severity error', async () => {
  const source = 'Sub Test()\nDim eNum As Long: eNum = err.Number\nEnd Sub';
  const result = await client.call('lint_module', {
    projectId: '00-gestion-riesgos-staging',
    module: 'TestModule',
    source
  });
  const dimColonError = result.results.find(r =>
    r.message.toLowerCase().includes('colon after dim') ||
    (r.ruleId === 'dim-colon-statement' && r.severity === 'error')
  );
  expect(dimColonError).toBeDefined();
  expect(dimColonError.line).toBe(2);
});
```

### Bug 2: Identificadores que colisionan con VBA reserved words (más allá de los globals)

#### Síntoma verificado

El consumer usó `Dim eNum As Long` en `Test_Issue64Cluster.bas:107`. Access reporta **`Error de sintaxis`** porque `eNum` (case-insensitive) colisiona con la keyword reservada `Enum`.

Captura visual: cursor en la línea `Dim eNum As Long`, Access muestra `Error de sintaxis`.

#### Evidencia de repro

```vba
' Línea 107:
Dim eNum As Long
' VBA parsea como: Dim Enum As Long
' → "Error de sintaxis: se esperaba: expresión"
```

El consumer perdió ~10 min identificando que `eNum` era el problema (no `:` ni nada sintáctico obvio).

#### Diagnóstico preliminar

La regla actual `forbidden-name (F22)` flag identifiers que shadow VBA / Access / DAO globals: `Err, Date, Name, Form, DoCmd, etc.` (case-insensitive). Está hardcoded para esta lista corta.

**Lo que falta**: cover TODAS las VBA reserved words (keywords del lenguaje). Lista mínima a cubrir (case-insensitive):
- Declaraciones: `Dim`, `Const`, `Static`, `Public`, `Private`, `Friend`, `ReDim`, `Type`, `Enum`, `WithEvents`
- Control: `If`, `Then`, `Else`, `ElseIf`, `End`, `For`, `Next`, `While`, `Wend`, `Do`, `Loop`, `Select`, `Case`, `Exit`, `Goto`, `GoTo`
- Procedimiento: `Sub`, `Function`, `Property`, `Set`, `Let`, `Call`, `Return` (VBA 7+)
- Tipos: `Integer`, `Long`, `Single`, `Double`, `Currency`, `String`, `Boolean`, `Date`, `Object`, `Variant`, `Byte`, `Integer`
- Operadores: `And`, `Or`, `Not`, `Xor`, `Eqv`, `Imp`, `Mod`, `Like`
- Objetos intrínsecos: `Err`, `Me`, `Nothing`, `Null`, `Empty`, `True`, `False`
- I/O: `Open`, `Close`, `Input`, `Output`, `Print`, `Write`, `Seek`, `Get`, `Put`

No es necesario enumerar TODAS — el linter puede consultar la lista oficial de VBA. Lo crítico es que `eNum`, `fnTest`, `subFoo`, `ifX`, `forY`, `dimZ`, `constW`, `typeA`, `modB` (todos lower-case collisiones) sean flagged.

#### Riesgo

Cualquier consumer que use nombres de variables con prefijos comunes (`e*`, `f*`, `s*`, `i*`, `d*`) puede colisionar silenciosamente. El tiempo de debugging es alto porque VBA reporta error genérico "se esperaba: expresión" sin identificar el identifier problemático.

#### Tests RED sugeridos

Test 1:

```ts
it('flags Dim eNum because eNum collides with Enum (case-insensitive)', async () => {
  const source = 'Sub Test()\nDim eNum As Long\nEnd Sub';
  const result = await client.call('lint_module', {
    projectId: '00-gestion-riesgos-staging',
    module: 'TestModule',
    source
  });
  const reservedWordError = result.results.find(r =>
    (r.ruleId === 'forbidden-name' || r.ruleId === 'reserved-word') &&
    r.message.toLowerCase().includes('enum')
  );
  expect(reservedWordError).toBeDefined();
  expect(reservedWordError.line).toBe(2);
});

it('flags Sub foo() because Sub collides with statement keyword', async () => {
  const source = 'Sub Test()\nDim Sub As String\nEnd Sub';
  const result = await client.call('lint_module', {
    projectId: '00-gestion-riesgos-staging',
    module: 'TestModule',
    source
  });
  const subError = result.results.find(r =>
    r.message.toLowerCase().includes('sub')
  );
  expect(subError).toBeDefined();
});
```

Test 2 (regresión — F22 actual sigue cubriendo globals):

```ts
it('still flags Dim Err because F22 covers globals (regression)', async () => {
  const source = 'Sub Test()\nDim Err As String\nEnd Sub';
  const result = await client.call('lint_module', {
    projectId: '00-gestion-riesgos-staging',
    module: 'TestModule',
    source
  });
  expect(result.results.some(r => r.message.toLowerCase().includes('err'))).toBe(true);
});
```

### Bug 3: Nombres duplicados de Sub/Function entre módulos no se detectan

#### Síntoma verificado

El consumer tenía 2 archivos .bas:
- `Test_Fixtures.bas:148`: `Public Sub SeedAll()`
- `Test_Fixtures.bas:158`: `Public Sub TeardownAll()`

Y luego en `Test_Issue64Cluster.bas` (test nuevo) declaró:
- `Test_Issue64Cluster.bas:57`: `Private Sub SeedAll(Optional ByVal cfgError As String = "")`
- `Test_Issue64Cluster.bas:115`: `Private Sub TeardownAll()`

Access reporta **`Error de sintaxis`** en `Test_Issue64Cluster.bas` con cursor en `TeardownAll`. Mensaje genérico — el parser no sabe distinguir "duplicate procedure name" de otros errores sintácticos.

#### Evidencia de repro

```vba
' Test_Fixtures.bas (existente, público):
Public Sub SeedAll()
Public Sub TeardownAll()

' Test_Issue64Cluster.bas (nuevo, intentaba ser privado):
Private Sub SeedAll(Optional ByVal cfgError As String = "")  ' <-- nombre duplicado
Private Sub TeardownAll()                                    ' <-- nombre duplicado
```

El consumer perdió ~20 min identificando el problema (no es obvio porque `Private` debería "ocultar" el `Public`, pero VBA no lo respeta entre módulos).

#### Diagnóstico preliminar

VBA reporta "duplicate procedure name" como **error de compilación** (no warning). Esto bloquea toda la compilación del módulo.

El linter NO detecta duplicados de nombres de Sub/Function entre módulos del mismo proyecto. Cada módulo se analiza individualmente sin contexto cross-module.

Solución propuesta: agregar una regla `cross-module-duplicate-name` que:
1. Recolecte todos los nombres de Sub/Function/Property de todos los .bas/.cls del proyecto
2. Marque como `error` cualquier nombre duplicado (incluso entre `Public` y `Private`)

El consumer ya mitigó este caso renombrando a `SeedAll_Cluster` y `TeardownAll_Cluster` — pero es un band-aid, no una solución de raíz.

#### Riesgo

Cualquier consumer que escriba tests con helpers privados que coincidan con nombres de fixtures helpers se encuentra con errores de compilación silenciosos. Esto bloquea el TDD cycle (no se puede ejecutar tests hasta arreglar el nombre).

#### Tests RED sugeridos

Test 1 (caso negativo — debe detectar duplicado cross-module):

```ts
it('flags duplicate Sub name across two standard modules', async () => {
  // Setup: 2 modules both declaring Sub SeedAll
  await client.call('import_modules', { projectId, moduleNames: ['ModuleA'] });
  await client.call('import_modules', { projectId, moduleNames: ['ModuleB'] });
  
  const result = await client.call('lint_module', {
    projectId,
    module: 'ModuleB',
    source: 'Sub SeedAll()\nEnd Sub'
  });
  // The lint should scan all modules of the project, not just ModuleB
  const dupError = result.results.find(r =>
    r.ruleId === 'cross-module-duplicate-name' &&
    r.message.includes('SeedAll')
  );
  expect(dupError).toBeDefined();
});
```

Test 2 (Public + Private):

```ts
it('flags Public in ModuleA + Private in ModuleB as duplicate', async () => {
  // Even with different visibility, VBA rejects duplicate names
  const result = await client.call('lint_module', {
    projectId,
    module: 'ModuleB',
    source: 'Private Sub Shared()\nEnd Sub'  // Shared() is Public in ModuleA
  });
  expect(result.results.some(r =>
    r.ruleId === 'cross-module-duplicate-name'
  )).toBe(true);
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR).
- Conventional commits con scope `lint`. Sugerencia: `feat(lint): detect Dim-colon statement anti-pattern`, `feat(lint): cover all VBA reserved words`, `feat(lint): detect cross-module duplicate procedure names`.
- NO tocar las 9 reglas existentes del linter. Solo AGREGAR las 3 nuevas.
- NO tocar las capabilities ni el schema de respuesta de `lint_module` (mantener backwards compatibility).
- Mantener `humanCompilePending:false` y la regla "human compiles" (los consumers no esperan que el linter bloquee, solo que avise).

## Acceptance output

- 3 PRs (uno por bug, como el consumer prefiere mantener commits atómicos y revisar cada fix independientemente).
- Changelog en `<repo>/CHANGELOG.md` con 3 bullets:
  - `Lint: detect Dim x As Type: y = expr (Dim-colon anti-pattern). (#<issue>)`
  - `Lint: cover all VBA reserved words in forbidden-name rule. (#<issue>)`
  - `Lint: detect cross-module duplicate Sub/Function names. (#<issue>)`
- Version bump a 2.1.8 (patch).
- 6 tests RED → GREEN (2 por bug).
- Actualizar `<repo>/tools/lint/rules.ts` (o path equivalente) con las 3 reglas nuevas.

## Quick start

```bash
git clone <repo path>
cd <repo>
git checkout -b fix/lint-gaps-round-1
pnpm install
pnpm run dev  # arranca el MCP localmente
```

Test repro contra el MCP local:

```bash
# Repro del Bug 1 (Dim-colon)
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "lint_module",
  "arguments": {
    "projectId": "00-gestion-riesgos-staging",
    "module": "TestModule",
    "source": "Sub Test()\nDim eNum As Long: eNum = err.Number\nEnd Sub"
  }
}'
# Esperado post-fix: result.results contiene error "Dim-colon-statement" en line 2
# Actual: result.results vacío (no detecta)

# Repro del Bug 2 (eNum reserved)
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "lint_module",
  "arguments": {
    "projectId": "00-gestion-riesgos-staging",
    "module": "TestModule",
    "source": "Sub Test()\nDim eNum As Long\nEnd Sub"
  }
}'
# Esperado post-fix: result.results contiene error "reserved-word:enum" en line 2
# Actual: result.results vacío

# Repro del Bug 3 (cross-module dup)
curl -X POST http://localhost:<port>/mcp -d '{
  "tool": "lint_module",
  "arguments": {
    "projectId": "00-gestion-riesgos-staging",
    "module": "TestIssue64Cluster"
  }
}'
# Esperado post-fix: result.results contiene error "cross-module-duplicate-name:SeedAll" + "TeardownAll"
# Actual: result.results vacío
```

## Reinforcement

El linter es la primera línea de defensa contra errores sintácticos que rompen la compilación de Access. Si el lint no detecta estos 3 patrones, los consumers pierden tiempo de debugging manual que no escala (especialmente en agent loops donde el agente escribe fixtures con muchas variantes de `Dim`). El fix debe mantener el contrato: el lint es ADVERTENCIA TEMPRANA, no bloqueante. El bloque sigue siendo la compilación humana.

Si el fix no cumple esta promesa, escalar a Round 2 del lint.
