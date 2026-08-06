Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Branch: <fix-rama-sugerida>. Versión: 2.10.0 (live runtime verificado 2026-07-14 en consumer `gestion_riesgos`, con #861 OPEN reportando inestabilidades post-release).

## Contexto del round

Round 9 = un sister issue en DysTelefonica/dysflow para un **lint genérico de "missing callees" en VBA**. No es bug-hunt: es una feature pedida por el consumer `gestion_riesgos` (issue #69 alli) tras un incidente real (issue #64: rechazo flow roto por un helper-method cuyo shape no fue verificado estaticamente).

**Este lint NO es specifico de gestion_riesgos.** Es una utility CI-side aplicable a cualquier consumer con source tree VBA (.cls/.bas). El consumer-side baseline concreto esta en `gestion_riesgos` (vía #69), pero la ubicacion correcta del codigo es **dysflow mismo** (`DysTelefonica/dysflow/tools/` o equivalente) porque:

1. Es generico (no tiene estado del consumer)
2. Multi-consumer: `no_conformidades`, `condor`, `cadete`, `brass`, `hps` y el mismo `gestion_riesgos` se benefician
3. Coherente con herramientas existentes como `lint_form_code`, `lint_module`, `verify_code`
4. Otras herramientas CI-related de dysflow viven en el mismo repo

Rounds previos del maintainer-prompt series (todos en `gestion_riesgos`-side o DysTelefonica/dysflow-side):
- Round 1-4: cerrados
- Round 5 (#110 codegraph-vba): cerrado
- Round 6 (#849): import_modules re-vincula code-module -- cerrado PR #854 en v2.10.0
- Round 7 (#852): non-canonical form names + source resolver triple-prefix -- cerrado PR #853 en v2.10.0
- Round 8 (#861): v2.10.0 inestabilidad post-release -- OPEN con 5 acceptance criteria, esperar v2.10.1
- **Round 9 (este)**: tools(lint): missing-callees detector para VBA

## Lo que YA funciona (NO tocar en este round)

- `lint_module` con identifier-safety, arg-type-match, declaration-order, forbidden-name (issue #768 de dysflow). El nuevo lint es **complementario**, no solapa con `lint_module`.
- `lint_form_code` con form-control-binding, listbox-no-list-assignment, bare-function-call-with-parens, named-and-positional-args-mixing, unicode-sensitive-executable-tokens, control-property-support. El nuevo lint NO es form-specific; cubre classes/, forms/ y modules/.
- `codegraph-vba` v1.7.0 indexado de callsites (round-5 work). El nuevo lint es más simple y rapido (regex sobre source); corre como guard CI en ~5s. Para analisis profundo, codegraph-vba sigue siendo la herramienta.
- Toda la policy safe-by-default y la regla cross-project "human compiles".
- `verify_code` (verify-side del drift source-binario) -- distinto scope.

## Baseline concreto que el consumer entrego

El consumer `gestion_riesgos` escribio un baseline funcional del lint en `tools/lint-missing-callees.mjs` (17.816 bytes, Node.js ESM, sin dependencias npm mas alla de `node:fs`/`node:path`/`node:process` del stdlib). **El codigo entero esta embebido al final de este issue como "Consumer baseline code"** -- el maintainer puede adoptarlo tal cual, refactorizarlo, o usarlo solo como referencia arquitectonica.

Honest disclosure sobre el baseline:
- **Works**: exclusion lists (VBA keywords, DAO/Database/Recordset/TableDef runtime members, Access event handlers via `_Click`/`_Open`/`_Load` suffix detection), JSON output mode, exit codes 0/1/2, performance <100ms en ~270 archivos, error format `src/path:N:C missing callee: Module.Name (call|method)`.
- **Known bug**: la regex `DECL_RE` no matchea declaraciones en el consumer's test (devuelve 0 declaraciones cuando deberia detectar >500). Patron regex correcto es `/^\s*(?:Public|Private)\s+(?:Function|Sub|Property(?:\s+(?:Get|Let|Set))?)\s+([A-Za-z_][A-Za-z0-9_]*)/gm` -- V8 da false en tests unitarios. La causa exacta requiere diagnistico del engine regex de V8 (probablemente relacionada con `(?:\s+(?:Get|Let|Set))?` en multilinea). Recomendacion al maintainer: simplificar la regex a `/^\s*(?:Public|Private)\s+(?:Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)/gm` mas una segunda regex para `Property`, ambas probadas independientemente antes de integrar.
- **Missing**: cobertura de `With x / End With` (el receiver dentro de `With` requiere tracking de tipo explicito). Cobertura de `For Each` (la variable del for-each tiene miembros de la coleccion). Cobertura de `Set var = New X` (crea una "declaration implicita" del tipo X).
- **Hardcoded exclusion lists**: listas de DAO/Dictionary/Collection/MSForms son razonables pero incompletas. Si el maintainer quiere completarlas, el approach sugerido es derivarlas de `codegraph-vba`'s stub=true / declined-runtime map (ya en disco de dysflow).
- **No tests**: el baseline del consumer no incluye tests TDD para el lint mismo. El maintainer debe anadir tests antes de considerar v2.11.0 (acceptance criteria abajo).

## Acceptance output (5 criterios de aceptacion explicitos y testables desde el lado del consumer)

Estos 5 puntos son el contrato "done" para este issue. Cada uno es testeable desde fuera de dysflow (sin requerir conocimiento interno del maintainer). Hasta que los 5 pasen, **gestion_riesgos#69 sigue OPEN con link a este upstream**. Una vez los 5 pasen en una version estable, gestion_riesgos#69 se cierra con merge-closure apuntando al release.

1. **MCP tool o CLI subcommand disponible.** Dysflow expone el lint como **un nuevo tool MCP** llamado `lint_missing_callees` (recomendado para coherencia con `lint_module`, `lint_form_code`) **O** como un subcomando CLI `dysflow lint callees <source-root>`. La eleccion entre MCP vs CLI queda al maintainer; lo que NO es aceptable es solo-docs o solo-referencia a una herramiento third-party separada. Argumentos del tool/CLI: `projectId` (para resolucion canónica), `sourceRoot` (default `src`), `json` (default false para output humano). Exit codes: `0` sin missing, `1` con missing, `2` error de invocación.

2. **Performance < 5 segundos sobre ~270 archivos .cls/.bas del consumer `gestion_riesgos`**. El baseline del consumer corre en ~85ms con regex simple. Hasta con parser AST de V8 o codegraph-vba, el budget de 5 segundos debe cumplirse. Si el approach elegido falla el budget, revertir a regex simple sin perder accuracy.

3. **Output estructurado y actionable.** Cada missing callee debe reportar: ruta relativa al project root (`src/forms/Foo.cls`), linea, columna, nombre del callee, module donde se usa, kind (`call` para `Foo(...)` o `method` para `.Foo(...)`). Output humano: `src/path:LINE:COL  missing callee: Module.Name (kind)`. Output JSON: `{ok, elapsedMs, totals: {declarations, missing, unused}, missing: [...], unused: [...]}`. El output debe ser consumible directamente por un PR-bot que añade un comentario de review.

4. **Lista de exclusions documentada y validada.** Las tres listas del baseline (VBA_KEYWORDS, RUNTIME_MEMBERS por container, IMPLICIT_MEMBERS) son validas pero incompletas. El maintainer debe: (a) documentar la fuente de cada exclusion (referencia a docs VBA, DAO reference, etc.); (b) anadir tests que aseguren que `Debug.Print`, `MsgBox`, `db.OpenRecordset(...)`, `Me.lblFoo.Visible = True` se excluyen correctamente; (c) exponer la lista como configurable por consumer (env var `DYSFLOW_LINT_EXTRAS` con JSON de additional-keyword sets) para casos no cubiertos.

5. **Tests TDD en el suite CI de dysflow.** Al menos 6 tests RED inicial que pasan tras el fix: (a) declaracion publica simple matchea; (b) declaracion publica con return type matchea; (c) declaracion privada matchea; (d) call site valido no emite missing; (e) call site a identificador no declarado emite missing con location correcta; (f) call site a identificador en exclusion list (e.g. `Debug.Print`) no emite missing. Nombre de tests informativos. Tests viven en `tests/lint-missing-callees.test.mjs` (o equivalente MCP-side). Deben correr en menos de 30s total en CI.

## Disciplina

- TDD estricto (RED -> GREEN -> REFACTOR). El primer test RED debe fallar con output actual (regex sin declaraciones detectadas, ej. `tests[0]: se esperaba 5 declaraciones, obtuvo 0`). El GREEN debe ser la implementacion que matchea correctament.
- Conventional commits con scope `tools(lint)`. Mensaje del commit: `feat(tools): add lint_missing_callees for VBA callee resolution (#<issue>)`.
- NO tocar `lint_module`, `lint_form_code`, `codegraph-vba`, ni v2.10.0 fixes (#849/#852). Cross-pollinar con esos proyectos solo si la fix lo exige, con PR separados.
- NO reintroducir `compile_vba` (regla cross-project).
- Si la fix cambia el shape de error envelope (ha pasado en #861), documentar en `references/error-codes.md`.
- Cross-consumer safety: la lista de exclusiones debe ser parametrizable, no hardcoded-por-consumer. **`gestion_riesgos` no debe requerir un fork** para usar el lint.

## Acceptance output (de cara al maintainer, cuando se mergea)

- PR con **6+ tests verdes** (los 5 obligatorios + 1 regression que verifica que el resto del suite MCP sigue verde).
- Changelog en `CHANGELOG.md` con bullet: `feat(tools): add lint_missing_callees -- missing-callees detector for VBA callee resolution (#<issue>)`. Incluir ejemplo de output.
- Version bump: **minor** (`v2.11.0`) porque anade un nuevo tool public-surface (heredando el patron de `lint_form_code` en v1.20.0).
- Documentacion actualizada: nueva seccion en `docs/runbooks/` o `docs/tools/lint-missing-callees.md` con: invocacion, opciones, configuracion (`DYSFLOW_LINT_EXTRAS`), como ignorar lineas in-line (`' dysflow:lint-ignore-line`), como integrar en GitHub Actions o pre-commit hook.
- Si la fix requiere un nuevo code-path en la skill `dysflow-usage` (consumer-side), coordinar con el consumer antes de publicar el changelog.

## Quick start

```bash
git clone <repo path>
cd <repo>
git checkout -b feat/tools-lint-missing-callees
<comando install>
<comando test>
```

Test repro contra el dev con consumer staging (`gestion_riesgos@staging`):

```bash
# Setup: clonar el consumer en /tmp/consumer-gr, su .dysflow/project.json apunta a este dev build de dysflow
cd /tmp/consumer-gr
node ../../dysflow/dist/index.js lint-missing-callees src 2>&1 | head -20
# Expected: lista de missing callees en formato  src/path:LINE:COL  missing callee: ...
# Actual en v2.10.0: tool no existe.
```

## Reinforcement

El lint se origina en un **fallo de produccion real**: el issue #64 rechazo flow roto por un helper-method no verificado estaticamente. El costo de discovery para el consumer fue:
1. Code reviewed PR, merged sin test del helper method
2. Deploy a staging
3. User ejecuta flujo de rechazo
4. Runtime error 91 surfaced
5. Sesion de debug de 30 min identificando el callee faltante

El ROI de este lint es **evitar ese ciclo para todos los consumers** que comparten el patron (cualquier proyecto VBA con shapes de clases implicitos). Ademas, **alimenta la base de conocimiento** de codegraph-vba: las invocaciones mal-formadas que el lint detecta son input data para entrenar el stub-resolver de codegraph-vba en futuras versiones.

Cross-pollination sugerida: cuando el lint detecte "missing callee" en el consumer, el output JSON deberia incluir la declaracion esperada (e.g. `expected: Public Function EsXNotificado(p_Correo As CORREO) As Boolean`) para que el developer pueda pegar la declaracion correcta. Esto acorta el fix loop de horas a minutos.

## Consumer baseline code (gestión de riesgos)

El codigo siguiente es el WIP del consumer, entregado como referencia arquitectonica. **No es la implementacion final** -- tiene un bug conocido en la regex de declaraciones (devuelve 0 en tests) y cobertura incompleta (no maneja `With`/`End With`, `For Each`, `Set var = New X`). El maintainer debe tratarlo como punto de partida, no como PR.

Path local: `gestion_riesgos/tools/lint-missing-callees.mjs` (branch WIP `tools/lint-missing-callees-wip`, no mergeada a staging).

```javascript
#!/usr/bin/env node
// tools/lint-missing-callees.mjs -- issue #69
//
// CI lint: emits non-zero exit code when a VBA call site in
// src/classes/*.cls, src/forms/*.cls, or src/modules/*.bas references
// an identifier that is NOT declared anywhere in the source tree.
//
// Discovered during issue #64 (an implicit shape dependency between
// the rechazo flow and helper members slipped through review and broke
// at runtime). This lint is a static guardrail that catches the class
// of bug before any user-facing failure.
//
// What is NOT covered (by design, per issue #69 acceptance criteria):
//   - VBA keywords / built-ins (Debug.Print, MsgBox, Err.Raise, etc.).
//   - External types: DAO.Recordset, DAO.Database, MSForms.*, etc.
//   - Standard library types (String, Dictionary, Collection, etc.).
//   - Form designer-generated code (Form_*.form.txt). Only .cls/.bas.
//
// Run:
//   node tools/lint-missing-callees.mjs            # lint src/
//   node tools/lint-missing-callees.mjs --json    # JSON output for CI
//
// Exit codes:
//   0  no missing callees
//   1  one or more missing callees detected (or read errors)
//   2  invocation error (no source root, etc.)
//
// Performance budget per issue #69 AC: < 5 seconds on the current
// source tree (~270 .cls/.bas files).

import { readdirSync, readFileSync, statSync } from 'node:fs';
import { extname, join, relative, resolve, sep } from 'node:path';
import { argv, exit } from 'node:process';

// --- Source roots ---------------------------------------------------------
const PROJECT_ROOT = resolve(new URL('..', import.meta.url).pathname);
const SRC_ROOT = join(PROJECT_ROOT, 'src');

const SOURCE_DIRS = [
  join(SRC_ROOT, 'classes'),
  join(SRC_ROOT, 'forms'),
  join(SRC_ROOT, 'modules'),
];

// --- Exclusion lists (per issue #69 AC) -----------------------------------
// VBA keywords: control flow, assignment, declarations, error handling,
// I/O. Match either as a bare identifier in call position OR as a
// dotted member (e.g. "Debug.Print", "Application.StatusBar").
const VBA_KEYWORDS = new Set([
  'Debug', 'Print', 'MsgBox', 'InputBox',
  'Err', 'Raise', 'Number', 'Description', 'Source', 'Clear',
  'Set', 'Let', 'Dim', 'Const', 'Static', 'ReDim', 'ReDimPreserve',
  'If', 'Then', 'Else', 'ElseIf', 'End', 'EndIf', 'EndWith', 'With',
  'For', 'Next', 'To', 'Step', 'Each', 'While', 'Wend', 'Do', 'Loop',
  'Select', 'Case', 'Function', 'Sub', 'Property', 'Get', 'SetStatement',
  'On', 'Error', 'Resume', 'Exit', 'Return', 'GoTo', 'GoSub', 'ReturnStatement',
  'Call', 'DoCmd', 'OpenForm', 'Close', 'Run', 'RunCommand', 'RunMacro',
  'Hourglass', 'SetWarning', 'Echo', 'Beep', 'Quit', 'Save',
  'OpenReport', 'OpenQuery', 'PrintOut', 'TransferSpreadsheet',
  'Application', 'TempVars', 'StatusBar', 'Eval', 'Execute',
  'VBProject', 'VBE', 'VBA', 'CreateObject', 'GetObject',
  'New', 'Nothing', 'True', 'False', 'Empty', 'Null', 'Me',
  'And', 'Or', 'Not', 'Xor', 'Eqv', 'Imp', 'Mod',
  'NothingStatement',
]);

// Standard library / runtime classes whose members we exclude.
// When a call site references `db.OpenRecordset(...)`, we treat
// `.OpenRecordset` as resolved by DAO.Database regardless of whether
// any local code declares an `OpenRecordset` Function.
const RUNTIME_MEMBERS = new Map([
  // [ContainerName, Set<memberName>]
  ['DAO.Database', new Set(['OpenRecordset', 'Execute', 'Close', 'RecordsAffected', 'TableDefs', 'QueryDefs', 'Recordsets', 'OpenRecordset', 'OpenSchema', 'CreateQueryDef', 'CreateTableDef', 'Version', 'CollatingOrder'])],
  ['DAO.Recordset', new Set(['AddNew', 'Update', 'Close', 'Edit', 'CancelUpdate', 'FindFirst', 'FindLast', 'FindNext', 'FindPrevious', 'BOF', 'EOF', 'MoveFirst', 'MoveLast', 'MoveNext', 'MovePrevious', 'OpenRecordset', 'AbsolutePosition', 'RecordCount', 'Filter', 'Index', 'Sort', 'NoMatch', 'Type', 'Source', 'LockType', 'Options', 'Bookmark', 'AbsolutePage', 'PageCount', 'PageSize', 'RecordStatus', 'Status', 'BatchCollisionCount', 'BatchCollisionValues', 'BatchSize', 'CacheSize', 'CacheStart', 'DateCreated', 'EditMode', 'LastModified', 'LastUpdated', 'Name', 'Parent', 'PercentPosition', 'Restartable', 'TransactionUpdate', 'Updatable', 'ValidationRule', 'ValidationText', 'Field', 'Fields'])],
  ['DAO.TableDef', new Set(['Fields', 'CreateField', 'OpenRecordset', 'Indexes', 'CreateIndex', 'Relations', 'CreateRelation', 'Name', 'SourceTableName', 'DateCreated', 'LastUpdated'])],
  ['DAO.Field', new Set(['Name', 'Type', 'Value', 'Size', 'DefaultValue', 'Required', 'AllowZeroLength', 'OrdinalPosition', 'ValidationRule', 'ValidationText', 'SourceField', 'ForeignName'])],
  ['DAO.Workspace', new Set(['OpenDatabase', 'Close', 'Users'])],
  ['DAO.DBEngine', new Set(['Workspaces', 'Errors', 'SetOption', 'GetOption'])],
  ['DAO.Error', new Set(['Number', 'Description', 'Source', 'HelpContext', 'HelpFile', 'LastDbUpdate'])],
  ['DAO.Relation', new Set(['Fields', 'Name', 'Table', 'ForeignTable'])],
  ['DAO.Index', new Set(['Fields', 'Fields', 'CreateField', 'Name', 'Unique'])],
  ['DAO.Property', new Set(['Name', 'Value', 'Type'])],
  ['DAO.QueryDef', new Set(['Parameters', 'Fields', 'RecordsAffected', 'OpenRecordset', 'Execute', 'Name', 'SQL', 'Type', 'Updatable', 'ReturnsRecords', 'DateCreated'])],
  ['Dictionary', new Set(['Add', 'Exists', 'Keys', 'Items', 'Remove', 'RemoveAll', 'Count', 'Item', 'CompareMode', 'Key'])],
  ['Collection', new Set(['Add', 'Remove', 'Count', 'Item', 'NewEnum'])],
  ['FileSystemObject', new Set(['GetFile', 'GetFolder', 'GetDrive', 'CreateTextFile', 'OpenTextFile', 'FileExists', 'FolderExists', 'DriveExists', 'GetFileName', 'GetParentFolderName', 'GetExtensionName', 'GetBaseName'])],
  ['TextStream', new Set(['ReadAll', 'ReadLine', 'Write', 'WriteLine', 'WriteBlankLines', 'AtEndOfStream', 'Close', 'Skip'])],
  ['UserForm', new Set(['Show', 'Hide', 'Caption', 'Controls'])],
  ['MSForms.Control', new Set(['SetFocus', 'Visible', 'Enabled', 'Value', 'Caption', 'Text', 'BackColor', 'ForeColor', 'Add', 'Remove', 'Clear'])],
  ['Image', new Set(['Picture', 'PictureAlignment'])],
  ['ErrObject', new Set(['Number', 'Description', 'Source', 'HelpContext', 'HelpFile', 'Clear', 'Raise'])],
  ['String', new Set()],
  ['Integer', new Set()],
  ['Long', new Set()],
  ['Boolean', new Set()],
  ['Double', new Set()],
  ['Variant', new Set()],
  ['Object', new Set()],
  ['Date', new Set()],
  ['Currency', new Set()],
  ['Byte', new Set()],
  ['Single', new Set()],
]);

// Implicit late-bound call sites — these access magic members (forms
// reachable through Forms/Reports/DoCmd, control collections, etc.)
// and are not declared statically. Exclusion is the only sane option
// without parser-level type inference.
const IMPLICIT_MEMBERS = new Set([
  'Forms', 'Reports', 'Screen', 'ActiveControl', 'ActiveForm',
  'Controls', 'Properties', 'Pages',
]);

// --- Parse helpers ---------------------------------------------------------

// Match VBA Sub/Function/Property declarations, both `Public` and `Private`.
// Captures the identifier name. Property Get/Let/Set all live under
// `Property`, so this single regex covers the three subkinds.
//
// We deliberately ignore declarations preceded by a comment line
// `'` because rare inline-decl docs in this repo already match.
const DECL_RE = /^\s*(?:Public|Private)\s+(?:Function|Sub|Property(?:\s+(?:Get|Let|Set))?)\s+([A-Za-z_][A-Za-z0-9_]*)/gm;

// Match any `<ident>(` candidate — a call site or a function name
// appearing as a callee. We then filter by exclusions and module-level
// declaration registry.
const CALL_RE = /\b([A-Za-z_][A-Za-z0-9_]*)\s*\(/g;

// Match `<dotted>.<ident>(` candidates — i.e. method calls on a
// receiver expression. We use this to scan `.OpenRecordset(...)`,
// `.RegistrarNotasParaCalidad(...)`, etc.
const METHOD_CALL_RE = /\.([A-Za-z_][A-Za-z0-9_]*)\s*\(/g;

/**
 * Walk a directory recursively, returning every .cls / .bas file.
 * Sorted paths for deterministic output.
 *
 * @param {string} dir
 * @returns {string[]}
 */
function listVbFiles(dir) {
  const out = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...listVbFiles(full));
    } else if (entry.isFile() && (entry.name.endsWith('.cls') || entry.name.endsWith('.bas'))) {
      out.push(full);
    }
  }
  return out;
}

/**
 * Extract module name from filename. Convention in this repo:
 *   Form_FormXyz.cls    -> module "Form_FormXyz" (class)
 *   Form_FormXyz.bas    -> module "Form_FormXyz" (standard)
 *   ClassFoo.cls        -> module "ClassFoo"
 *   Foo.bas             -> module "Foo"
 * The VBA attribute VB_Name can override this; we honor it if present.
 *
 * @param {string} filePath
 * @param {string} text
 * @returns {string}
 */
function moduleNameOf(filePath, text) {
  const m = text.match(/^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/m);
  if (m) return m[1];
  const base = filePath.split(sep).pop();
  return base.replace(/\.(cls|bas)$/i, '');
}

/**
 * Parse a source file and return:
 *   - moduleName: parsed VB_Name or filename
 *   - declaredCallees: Set of identifier names declared as
 *      Public/Private Function|Sub|Property in this file
 *   - referencedCallees: array of { name, line, column, kind } for
 *      every potential callee found (unresolved yet)
 *
 * @param {string} filePath
 */
function parseFile(filePath) {
  const text = readFileSync(filePath, 'utf8');
  const moduleName = moduleNameOf(filePath, text);
  const declaredCallees = new Set();

  for (const m of text.matchAll(DECL_RE)) {
    declaredCallees.add(m[1]);
  }

  const referencedCallees = [];

  // Bare function-style calls — `Foobar(...)`, `Call Foobar(...)`.
  for (const m of text.matchAll(CALL_RE)) {
    referencedCallees.push({
      name: m[1],
      line: lineOf(text, m.index),
      column: columnOf(text, m.index),
      kind: 'call',
    });
  }

  // Dotted method calls — `.X(...)`. We only flag the right-hand
  // identifier so the exclusion list can resolve known runtime
  // members; the receiver itself (e.g. `db`, `GetDb()`) is not
  // resolved (a separate type-inference lint could close that gap).
  for (const m of text.matchAll(METHOD_CALL_RE)) {
    referencedCallees.push({
      name: m[1],
      line: lineOf(text, m.index),
      column: columnOf(text, m.index + 1), // +1 because lineOf of index before '.'; visually anchored to '.'
      kind: 'method',
    });
  }

  return { moduleName, declaredCallees, referencedCallees };
}

function lineOf(text, index) {
  let line = 1;
  for (let i = 0; i < index && i < text.length; i++) {
    if (text[i] === '\n') line++;
  }
  return line;
}

function columnOf(text, index) {
  let col = 1;
  for (let i = 0; i < index && i < text.length; i++) {
    if (text[i] === '\n') col = 1;
    else col++;
  }
  return col;
}

function isExcluded(name) {
  return VBA_KEYWORDS.has(name) || IMPLICIT_MEMBERS.has(name);
}

function isRuntimeMemberCall(member) {
  for (const [, members] of RUNTIME_MEMBERS) {
    if (members.has(member)) return true;
  }
  return false;
}

// --- Main ------------------------------------------------------------------

function lint() {
  // 1. Inventory all declarations across the source tree.
  /** @type {Map<string, { file: string, line: number }>} */
  const declarations = new Map();
  /** @type {{ module: string, file: string, declared: Set<string>, refs: { name: string, line: number, column: number, kind: string }[] }[]} */
  const modules = [];

  for (const dir of SOURCE_DIRS) {
    let files;
    try {
      files = listVbFiles(dir);
    } catch (e) {
      if (e.code === 'ENOENT') continue; // missing submodule directory is fine
      throw e;
    }
    for (const file of files) {
      const parsed = parseFile(file);
      modules.push({
        module: parsed.moduleName,
        file,
        declared: parsed.declaredCallees,
        refs: parsed.referencedCallees,
      });
      for (const name of parsed.declaredCallees) {
        // First-seen wins. Subsequent declarations on the same name are
        // tracked as ambiguous by the compiler, not by us — we just
        // surface the location.
        if (!declarations.has(name)) {
          declarations.set(name, { file: relative(PROJECT_ROOT, file), line: 1 });
        }
      }
    }
  }

  // 2. For each call site, decide whether it is "resolved" (i.e. a known
  // declaration or a known runtime/keyword) or "missing".
  /** @type {{ module: string, file: string, name: string, line: number, column: number, kind: string }[]} */
  const missing = [];
  // Track declarations that never appear as called callees — informational.
  /** @type {string[]} */
  const unused = [];

  const seen = new Set(declarations.keys());
  const allDeclaredAcrossModules = new Set();
  for (const m of modules) {
    for (const d of m.declared) allDeclaredAcrossModules.add(d);
  }

  // Compute unused AFTER resolving missing (we don't expect all declarations
  // to be called — e.g. Event sinks, helpers exported for tool-side use).
  for (const m of modules) {
    const localDeclared = new Set(m.declared);
    for (const ref of m.refs) {
      const { name, line, column, kind } = ref;
      if (isExcluded(name)) continue;
      if (allDeclaredAcrossModules.has(name)) continue;
      if (kind === 'method' && isRuntimeMemberCall(name)) continue;
      // Local declarations count too — but for forms, Public Sub X /
      // Function X is the entry point and `_Click` handlers are called
      // by Access runtime, not by user code, so they are implicit and
      // already excluded by the `_` suffix in isExcluded? No — they're
      // not. Add a soft heuristic: names ending in _Click, _Open, _Load,
      // _Close, _Unload, _Resize, _Enter, _Exit, _AfterUpdate, _BeforeUpdate,
      // _Change, _Current, _Dirty, _KeyPress, _KeyDown, _KeyUp, _MouseDown,
      // _MouseUp, _Timer, _DataChange, _Undo etc. are Access event handlers
      // resolved at runtime, not by callee lint.
      if (kind === 'call' && isAccessEventHandler(name)) continue;
      missing.push({
        module: m.module,
        file: relative(PROJECT_ROOT, m.file),
        name,
        line,
        column,
        kind,
      });
    }
  }

  // For "unused" informational reporting: declarations never referenced
  // anywhere. We exclude event handlers and names matching the standard
  // public-event-export pattern (Public Event X).
  for (const name of allDeclaredAcrossModules) {
    if (isAccessEventHandler(name)) continue;
    if (isExcluded(name)) continue;
    let referenced = false;
    for (const m of modules) {
      for (const r of m.refs) {
        if (r.name === name) { referenced = true; break; }
      }
      if (referenced) break;
    }
    if (!referenced) unused.push(name);
  }

  return { missing, unused, totalDeclarations: allDeclaredAcrossModules.size };
}

// Access form event handlers — these are entry points called by the
// Access runtime via the form's event binding, not by user code.
// Listing them explicitly avoids false positives.
function isAccessEventHandler(name) {
  // Convention: <handler>_<event>
  const idx = name.lastIndexOf('_');
  if (idx < 0) return false;
  const event = name.slice(idx + 1);
  return ACCESS_EVENTS.has(event);
}

const ACCESS_EVENTS = new Set([
  'Click', 'DblClick', 'MouseDown', 'MouseUp', 'MouseMove',
  'KeyPress', 'KeyDown', 'KeyUp',
  'Open', 'Load', 'Resize', 'Unload', 'Close',
  'Current', 'BeforeInsert', 'AfterInsert',
  'BeforeUpdate', 'AfterUpdate', 'BeforeDelConfirm', 'AfterDelConfirm',
  'BeforeChange', 'AfterChange',
  'Dirty', 'Undo', 'Change', 'Updated',
  'Error', 'Timer', 'GotFocus', 'LostFocus',
  'Enter', 'Exit',
  'Filter', 'ApplyFilter', 'BeforeScreenTip', 'AfterLayout',
  'BeforeQuery', 'AfterQuery', 'BeforeRender', 'AfterRender',
  'Deactivate', 'Activate', 'QueryClose', 'UnloadSync',
  'ViewChange', 'DataChange', 'DataSetChange', 'RecordChange', 'SelectionChange',
  'NotInList', 'BeforeUpdate', 'AfterUpdate',
  'OnClick', 'OnDblClick', 'OnOpen', 'OnClose',
  'OnGotFocus', 'OnLostFocus', 'OnEnter', 'OnExit',
]);

// --- CLI -------------------------------------------------------------------

const start = process.hrtime.bigint();
const result = lint();
const elapsedMs = Number(process.hrtime.bigint() - start) / 1e6;

const asJson = argv.includes('--json') || argv.includes('--machine');
if (asJson) {
  process.stdout.write(JSON.stringify({
    ok: result.missing.length === 0,
    elapsedMs: Math.round(elapsedMs * 100) / 100,
    totals: {
      declarations: result.totalDeclarations,
      missing: result.missing.length,
      unused: result.unused.length,
    },
    missing: result.missing,
    unused: result.unused,
  }, null, 2) + '\n');
} else {
  const lines = [];
  if (result.missing.length === 0) {
    lines.push(`OK: 0 missing callees across ${result.totalDeclarations} declarations (${elapsedMs.toFixed(0)}ms)`);
  } else {
    lines.push(`FAIL: ${result.missing.length} missing callee(s) across ${result.totalDeclarations} declarations (${elapsedMs.toFixed(0)}ms)`);
    lines.push('');
    for (const m of result.missing) {
      lines.push(`  src/${m.file}:${m.line}:${m.column}  missing callee: ${m.module}.${m.name} (${m.kind})`);
    }
  }
  if (process.env.LINT_VERBOSE === '1' && result.unused.length > 0) {
    lines.push('');
    lines.push(`INFO: ${result.unused.length} declared symbol(s) never called (informational, not a CI failure):`);
    for (const u of result.unused.slice(0, 50)) lines.push(`  - ${u}`);
    if (result.unused.length > 50) lines.push(`  ... and ${result.unused.length - 50} more`);
  }
  process.stdout.write(lines.join('\n') + '\n');
}

exit(result.missing.length === 0 ? 0 : 1);

```

> NOTA: pegar el codigo completo aqui en el issue -- el bloque anterior contiene el .mjs entero del consumer. Esto elimina la friccion de "tengo que navegar al consumer repo solo para ver que intentaron".

---

## Cross-session

**Round-9 NO reemplaza ni reabre ningun round previo.** Tras merge de este PR:
1. `gestion_riesgos#69` se cierra con merge-closure apuntando al SHA del PR de dysflow + version donde se incluye.
2. Consumer puede borrar el branch WIP `tools/lint-missing-callees-wip` y el archivo local (migrado a upstream).
3. El runbook de integration vive en el consumer como 3-lineas: `dysflow lint callees src/` en pre-commit / CI.

Si la fix del round-9 se mergear antes que #861 (v2.10.0 inestabilidad), el release que la incluye debe **no** ser v2.10.x (porque v2.10.x esta en proceso de patch #861). Target: **v2.11.0** como minor release.

Rounds a chequear antes de mergear este PR para no pisar trabajo en flight:
- #849 (round-6) -- CERRADO PR #854 -- no toca `tools/lint-*`
- #852 (round-7) -- CERRADO PR #853 -- no toca `tools/lint-*`
- #861 (round-8) -- OPEN -- merge despues o en paralelo si scopes son disjuntos
