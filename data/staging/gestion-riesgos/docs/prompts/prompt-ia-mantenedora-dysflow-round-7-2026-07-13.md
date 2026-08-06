Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Branch: <fix-rama-sugerida>. Versión: 2.9.2 (live runtime verificado 2026-07-13 en consumer `gestion_riesgos`).

## Contexto del round

Round 7 = dos bugs descubiertos al intentar automatizar la workaround documentada en #849 (round-6: `import_modules` reescribe `.form.txt` pero no re-vincula code-module → `Me.X undefined` al compile). Esos dos bugs **NO están cubiertos por #849**: distinta operación (`delete_module` no `import_modules`) y distinto modo de fallo dentro de `import_modules` (path resolver con triple prefijo + excepción opaca, no `Me.X undefined`). El consumer propuso este round para evitar fragmentar el trabajo del mantenedor y dejar consolidada la superficie completa de "form name resolution en dysflow 2.9.2".

Rounds previos del maintainer-prompt series:
- Round 1: capabilities block migration — cerrado.
- Round 2: correcciones de seguridad + scope post-round-1 — cerrado.
- Round 3: gap en `vba_inline_execution` — cerrado.
- Round 4: `lint_module` identifier-safety false positive para tildes/eñes — cerrado.
- Round 5: pivot consumer-side — cerrado.
- **Round 6 (#849, ABIERTO)**: `import_modules` no re-vincula code-module tras re-importar un form existente — workaround consumer-side = `delete_module` + `import_modules`. Este round NUNCA tocó `delete_module` ni el path resolver de source.
- **Round 7 (este)**: dos huecos que aparecen AL APLICAR la workaround de #849 masivamente (1 de 63 forms) y con nombres no canónicos (1 de 63 forms).

## Lo que YA funciona (NO tocar)

- Todo lo declarado en #849 (sección "Lo que YA funciona") sigue válido en este round.
- `delete_module(dryRun:false)` con nombres binarios canónicos (prefijo `Form` + resto): confirmado `status:"ok"` en este round para 35 forms consecutivos del consumer `gestion_riesgos`.
- `import_modules(dryRun:false, importMode:"auto")` con nombres source canónicos (`Form_<X>` ↔ binario `<X>`): confirmado `status:"ok"` en este round para 35 forms consecutivos.
- `verify_code`, `list_objects`, `list_vba_modules`, `get_capabilities` (live runtime verificado 2026-07-13, `adapterVersion: 2.9.2`, `toolsVisible: 80`, `effectiveDryRunDefault.delete_module: true`).
- Workaround `delete_module + import_modules` para los 3 forms confirmados en #849: sigue funcionando (verificado este round).
- Política `safe-by-default` y regla cross-project "human compiles".

## Lo que falta en este round

### Bug A: `delete_module` falla con HRESULT "valor fuera del intervalo esperado" cuando el nombre del form binario no sigue la convención `<Form_><X>`

#### Síntoma verificado

Consumer `gestion_riesgos` (DysTelefonica/GESTION_RIESGOS, branch `staging`). El form binario `frmSplash` (sin prefijo `Form_`, con `frm` minúscula) no se puede eliminar vía `delete_module`:

```js
await tools.dysflow.delete_module({
  projectId: "00-gestion-riesgos-staging",
  moduleNames: ["frmSplash"],  // binary name real (sin Form_)
  dryRun: false
});
// Resultado: VBA_MANAGER_FAILED exit code 1
// Mensaje: "DYSFLOW_RESULT [{\"code\":\"VBA_MANAGER_FAILED\",\"message\":...
//         \"No se pudo eliminar componente `Form_frmSplash`:
//          El valor no está fuera del intervalo esperado\"}]"
```

#### Causa raíz preliminar (NO verificada internamente)

La ruta interna de `delete_module` parece auto-prepender `Form_` al nombre cuando no lo tiene, produciendo `Form_frmSplash`. Pero el form binario real se llama `frmSplash`, así que el lookup llega a un VBComponent que no es la `Form_frmSplash` real (probablemente el `documentModule` residual) y devuelve HRESULT 0x80070057 (`E_INVALIDARG`, "valor fuera del intervalo esperado"). Confirmación cruzada: el `delete_module` masivo previo borró 62/63 forms OK; **el único que falló fue precisamente `frmSplash`** (el único cuyo nombre binario no empieza por `Form_` mayúscula).

#### Workaround consumer-side (NO verificado todavía)

- Candidatos pendientes de prueba, **no incluidos en este round** porque disparan write-class sobre el binary:
  1. `delete_module(moduleName: "Form_frmSplash", force: true)` con el nombre prefijado y `force:true` (el prefijo existe como documentModule residual tras la asignación del code-module).
  2. Excluir el form de los barridos bulk (lo que el consumer ya hizo: skip list).
- Conclusión: el consumer **bloqueó la automatización masiva** y mantiene `frmSplash` como skip list permanente hasta que el bug A esté resuelto upstream.

#### Riesgo

- Cross-fleet: cualquier consumer con un form cuyo binary name no siga `<Form_><X>` (por ejemplo: frm*, splash, dialog, login, etc.) sufrirá este fallo. En `gestion_riesgos` hay 1/63 = 1.6% de forms afectados. En `no_conformidades` se observó patrón idéntico (no verificado este round).
- Falla silenciosa: `delete_module` aborta todo el batch con `VBA_MANAGER_FAILED` y reporta UN SOLO error (no enumera fallos parciales). El consumer tuvo que instrumentar con `list_objects` post-delete para detectar el parcial.

### Bug B: `import_modules` con source `<Form_><X>` cuyo `.cls` no tiene `Attribute VB_Name` en línea 1 falla con excepción opaca `[object Object]` y path resolver triplica el prefijo

#### Síntoma verificado

Consumer `gestion_riesgos`. El source `Form_FormRiesgoBiblioteca.cls` (1580 líneas) empieza con:

```
Line 1: Option Compare Database
Line 2: Option Explicit
...
```

Es decir, **NO tiene la línea canónica `Attribute VB_Name = "Form_FormRiesgoBiblioteca"`** en la línea 1 del `.cls`. Ese atributo sí está presente pero enterrado dentro del `.form.txt` en la línea 1288.

Resultado al intentar import:

```js
await tools.dysflow.import_modules({
  projectId: "00-gestion-riesgos-staging",
  moduleNames: ["Form_FormRiesgoBiblioteca"],
  dryRun: false,
  importMode: "auto"
});
// Excepción cruda no tipada: [object Object]
// Cero envelope estructurado. Sin {code, message, phase}.
```

Comprobación cruzada con `inspect_form` (read-only):

```js
await tools.dysflow.inspect_form({
  projectId: "00-gestion-riesgos-staging",
  formName: "Form_FormRiesgoBiblioteca"
});
// FORM_NOT_FOUND: No form/report source found for "Form_FormRiesgoBiblioteca"
//   under source root "src".
// Checked: forms/Form_Form_FormRiesgoBiblioteca.form.txt
```

El path resolver intentó buscar **`forms/Form_Form_FormRiesgoBiblioteca.form.txt`** (triple prefijo). El archivo en disco es `forms/Form_FormRiesgoBiblioteca.form.txt` (doble prefijo correcto). El path resolver duplica `Form_Form_` cuando no encuentra el `Attribute VB_Name` en la línea 1 del `.cls`.

#### Causa raíz preliminar (NO verificada internamente)

El resolver de source probablemente:
1. Lee el primer línea del `.cls`.
2. Si la primera línea NO es `Attribute VB_Name = "<name>"`, no puede extraer el nombre canónico.
3. Como fallback, intenta prepender `Form_Form_` al filename (filename ya empieza con `Form_Form_`), generando triple prefijo.
4. Como el archivo no existe, dispara una excepción interna no envuelta (de ahí `[object Object]`).

Nota: 58/63 forms de este consumer también carecen de `Attribute VB_Name` en línea 1 del `.cls`, pero el path resolver **funcionó correctamente** para 34 de ellos en este round. La diferencia es que `Form_FormRiesgoBiblioteca` es el único donde el nombre del archivo (`Form_FormRiesgoBiblioteca`) tras prepender `Form_Form_` se convierte en un nombre que parece plausible pero no existe. El resolver no detecta el filename ya prefijado y duplica el prefijo.

Hipótesis alternativa: existe un bug general en el path resolver para cualquier source `.cls` con naming unusual, pero los 34 forms exitosos enmascararon el patrón porque su prefijo canónico `Form_` permitía que el filename original matcheara con el path canónico una sola vez.

#### Workaround consumer-side (NO verificado)

- Añadir `Attribute VB_Name = "Form_FormRiesgoBiblioteca"` como línea 1 del `.cls` antes del `import_modules`. No probado por el consumer porque requiere edición source-side y el contrato de este round excluye cambios source.
- Excluir el form de los barridos bulk (skip list, lo que el consumer ya hizo).

#### Riesgo

- Cross-fleet: cualquier consumer con un `.cls` sin `Attribute VB_Name` en línea 1 está expuesto. Estimación: muy común en proyectos legacy donde los `.cls` se generan/editaron directamente desde VBE sin pasar por un exportador que insertara el atributo en su lugar canónico.
- Excepción opaca: cualquier consumer que use `import_modules` en modo seguro (try/catch alrededor de la llamada) **no recibe un error tipado** que pueda clasificar (recoverable vs fatal, retry vs abort). El consumer tuvo que abrir el log de stderr para detectar el `[object Object]` y cruzarlo con `inspect_form` para entender qué pasó.

### Relación con #849 (round-6)

- **#849**: el form YA EXISTE en el binary; tras `import_modules`, `Me.X` resuelve contra instancia fantasma del pre-import. **Workaround**: delete + reimport.
- **Bug A de este round**: el workaround de #849 FALLA en `delete_module` cuando el nombre binario no empieza por `Form_`. Imposible completar la fix de #849 en estos forms.
- **Bug B de este round**: el workaround de #849 FALLA en `import_modules` cuando el `.cls` no tiene `Attribute VB_Name` línea 1. Imposible completar la fix de #849 en estos forms.
- **Conclusión**: mientras bugs A y B no estén resueltos, la fix de #849 NO es aplicable masivamente. El consumer está bloqueado.

## Tests RED sugeridos

```ts
// Test A1: delete_module acepta nombres binarios que no siguen la convención Form_<X>
describe('delete_module: non-canonical form name handling', () => {
  it('deletes form whose binary name starts with frm* (e.g. frmSplash)', async () => {
    await setupTestBinary({ forms: [{ name: 'frmSplash', moduleName: 'Form_frmSplash' }] });
    await expect(client.call('delete_module', { moduleNames: ['frmSplash'], dryRun: false }))
      .resolves.toMatchObject({ status: 'ok' });
  });

  it('returns typed error instead of HRESULT when delete hits an unexpected VBComponent', async () => {
    // Si un lookup equivocado produce HRESULT, debe envolverse en { code, message, candidate }
    await setupTestBinary({ forms: [{ name: 'frmSplash', moduleName: 'Form_frmSplash' }] });
    await expect(client.call('delete_module', { moduleNames: ['frmSplash_BAD'], dryRun: false }))
      .rejects.toMatchObject({ code: expect.stringMatching(/MODULE_NOT_FOUND|NOT_FOUND|INVALID_ARG/) });
  });
});

// Test B1: import_modules resuelve source correctamente cuando .cls no tiene VB_Name línea 1
describe('import_modules: source path resolver handles cls without line-1 Attribute VB_Name', () => {
  it('imports form whose .cls lacks Attribute VB_Name at line 1 but has it in .form.txt', async () => {
    await setupTestBinary({
      forms: [{ sourceName: 'Form_FormRiesgoBiblioteca', vbNameInFormTxtOnly: true }]
    });
    await client.call('import_modules', {
      moduleNames: ['Form_FormRiesgoBiblioteca'],
      dryRun: false, importMode: 'auto'
    });
    // Assert: import tuvo éxito o error tipado (NO [object Object])
    await expect(client.call('list_objects', {})).resolves.toMatchObject({
      forms: expect.arrayContaining([{ name: expect.stringMatching(/RiesgoBiblioteca/) }])
    });
  });

  it('returns typed error when source cannot be resolved (no opaque exception)', async () => {
    await setupTestBinary({
      forms: [{ sourceName: 'Form_FormRiesgoBiblioteca_BAD', missing: true }]
    });
    // La respuesta debe ser { code: 'SOURCE_NOT_FOUND', ... } - nunca [object Object]
    await expect(client.call('import_modules', { moduleNames: ['Form_FormRiesgoBiblioteca_BAD'], dryRun: false }))
      .rejects.toMatchObject({ code: expect.stringMatching(/SOURCE_NOT_FOUND|FORM_NOT_FOUND|FORM_RESOLVE_FAILED/) });
  });

  it('does not triplicate Form_Form_ prefix when filename already starts with Form_Form_', async () => {
    // Después de fix, inspect_form({formName:'Form_FormRiesgoBiblioteca'}) debe buscar
    //   forms/Form_FormRiesgoBiblioteca.form.txt (no triple prefix)
    const before = await client.call('inspect_form', { formName: 'Form_FormRiesgoBiblioteca' });
    expect(before.sourcePath).toBe('forms/Form_FormRiesgoBiblioteca.form.txt');
    expect(before.sourcePath).not.toMatch(/Form_Form_Form_/);
  });
});

// Test C1: regresión — el workaround de #849 sigue funcionando en forms canónicos
describe('regression: delete+import workaround for #849 still works', () => {
  it('re-establishes Me.X resolvable for canonical Form_<X> forms', async () => {
    await setupTestBinary({ forms: [{ name: 'FormPublicacionCalidadPublicarEjecutar' }] });
    await simulateBulkImportBreak();
    await client.call('delete_module', { moduleNames: ['FormPublicacionCalidadPublicarEjecutar'], dryRun: false });
    await client.call('import_modules', { moduleNames: ['Form_FormPublicacionCalidadPublicarEjecutar'], dryRun: false });
    await expect(runCompileCheck()).resolves.toEqual([]);
  });
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR). Los tests RED deben fallar con el comportamiento actual y pasar tras la fix.
- Conventional commits con scope `delete-module` (bug A) y `import-modules` (bug B).
- NO tocar el comportamiento de `delete_module` ni `import_modules` para forms canónicos (regresión cubierta por Test C1).
- NO reintroducir `compile_vba` (regla cross-project).
- Si la fix cambia el shape del error envelope (de HRESULT crudo a tipado), documentar el cambio en `references/error-codes.md`.
- Cross-consumer safety: si la fix toca el path resolver de source, validar con `assets/scripts/verify-examples-vs-runtime.ps1` (debe pasar exit 0).

## Acceptance output

- PR con **5 tests verdes** (A1, A2, B1, B2, B3, C1 — son 6, entregar los 5 obligatorios + 1 de regresión).
- Changelog en `CHANGELOG.md` con bullets separados:
  - `fix(delete_module): accept non-canonical form binary names (e.g. frmSplash) and surface typed error envelope instead of HRESULT (#<issue-A>)`
  - `fix(import_modules): resolve source path correctly when .cls lacks Attribute VB_Name at line 1; no triple Form_Form_ prefix; no opaque [object Object] exceptions (#<issue-B>)`
- Version bump: **patch** (`v2.9.3`) si las fixes son internas; **minor** (`v2.10.0`) si cambia el shape observable del error envelope o el comportamiento para forms canónicos.
- Documentación actualizada: sincronizar `assets/write-flags-matrix.md` y `references/error-codes.md` con `assets/scripts/verify-examples-vs-runtime.ps1` (debe pasar exit 0).
- Si la fix requiere un nuevo code-path en `dysflow-usage` skill (consumer-side) para evitar el skip list manual de forms con nombre no canónico, coordinar con el consumer ANTES de publicar el changelog.

## Quick start

```bash
git clone <repo path>
cd <repo>
git checkout -b fix/delete-module-non-canonical-and-import-modules-source-resolver
<comando install>
<comando test>
```

Test repro contra el dev (necesita Access real):

```bash
# Setup: .accdb de prueba con Form_frmSplash (binary name "frmSplash", module name "Form_frmSplash")
node -e 'import("./dist/index.js").then(m => m.client.call("delete_module", {moduleNames: ["frmSplash"], dryRun: false}))'
# Esperado: status:"ok"  /  Actual: VBA_MANAGER_FAILED + HRESULT 0x80070057

# Setup: .accdb de prueba con Form_FormRiesgoBiblioteca (cls sin Attribute VB_Name línea 1)
node -e 'import("./dist/index.js").then(m => m.client.call("import_modules", {moduleNames: ["Form_FormRiesgoBiblioteca"], dryRun: false, importMode: "auto"}))'
# Esperado: status:"ok" o error tipado FORM_RESOLVE_FAILED
# Actual: excepción cruda [object Object]
```

## Reinforcement

Mismo principio que #849: discovery cost por consumer es alto (el human-compile gate o un rollback manual descubre el bug). Las fixes deben venir con **tests RED automatizables** que el CI pueda correr **sin abrir Access** (vía mocks del COM object o un harness que simule los VBComponents). Si la fix requiere Access real para validar, marcalo explícitamente como gap de testing y sugerí la estrategia de mock más cercana. Los bugs A y B son particularmente dolorosos porque aparecen SOLO cuando se intenta automatizar la workaround de #849 — exactamente el camino que cualquier consumer va a tomar tras leer #849.

## Cross-session

Round-7 NO cierra ni reemplaza #849. Tras la fix de este round, el consumer debería poder:
1. Aplicar la workaround de #849 masivamente (delete + reimport).
2. Hacerlo también para los forms con nombre no canónico (bug A).
3. Hacerlo también para los forms cuyo .cls no tiene VB_Name línea 1 (bug B).
4. Llegar a un estado donde `Debug → Compile VBA Project` no produzca `Me.X undefined`.

Si la fix de este round es suficiente para resolver masivamente el problema expuesto en #849 (no requiere fix adicional), marcar #849 como cerrado al mergear este PR.
