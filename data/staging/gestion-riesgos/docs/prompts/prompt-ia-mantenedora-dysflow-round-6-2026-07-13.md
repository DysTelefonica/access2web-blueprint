Eres la IA mantenedora de dysflow MCP. Repo: <repo path>. Branch: <fix-rama-sugerida>. Versión: 2.9.2 (live runtime verificado 2026-07-13 en consumer `gestion_riesgos`).

## Contexto del round

Round 6 = un único bug en `import_modules` que rompe el link entre el módulo `.cls` de código y la instancia del form después de un import. Tras bulk-import, las forms se re-crean con todos los controles, pero el VBA `Me.X` resuelve contra una instancia fantasma del form pre-import — sin los controles — y el compile falla con "No se encontró el método o el dato miembro" aunque el control está textualmente definido en el `.form.txt`.

Rounds previos del maintainer-prompt series:
- Round 1: capabilities block migration (CONFIG_TOP_LEVEL_FIELDS_REMOVED) — cerrado.
- Round 2: correcciones de seguridad + scope post-round-1 — cerrado.
- Round 3: gap en `vba_inline_execution` (procedimiento `__dysflow_inline__.ExecuteInline` no se inyecta antes del Run) — pendiente en repo consumer.
- Round 4: `lint_module` identifier-safety false positive para tildes/eñes en identificadores VBA — cerrado en dysflow 2.x.
- Round 5: pivot — gaps reframados como consumer-side symptoms (no runtime bugs) tras feedback del usuario. Documento archivado en `docs/prompts/prompt-ia-mantenedora-dysflow-round-5-2026-07-10.md`.

## Lo que YA funciona (NO tocar)

- `import_modules` con sha256 match para `.bas`/`.cls` no-form (verificado round-3 + repetido en este round).
- `delete_module(dryRun:false)` — limpia forms `.cls` y forms por separado (verificado `status:"ok"`).
- `verify_code` — produce `summaryStructured` con `missingInBinary`/`missingInSource`/`actionable`/`nonActionable` correcto.
- `get_capabilities` — devuelve live runtime (verificado 2026-07-13: `adapterVersion: 2.9.2`, `toolsVisible: 80`).
- `write_flask_matrix.md` consistente: `apply:true` o `dryRun:false` para commit; nunca `apply:true` + `compile:true`.
- Política `safe-by-default`: `effectiveDryRunDefault: true` por defecto; el consumer debe pasar `dryRun:false` explícito para escribir.
- Regla cross-project "human compiles" — `compile_vba` removido estructuralmente; cualquier intento de reintroducirlo debe escalar a consumer antes.
- `sync_binary` mantiene el 4-step lockstep pero **no llama `import_modules` internamente para forms** (verificado: el `plan.toImport` para forms sale vacío incluso con `includeBothChanged:true`; eso es independiente del bug de este round).

## Lo que falta en este round

### Bug único: `import_modules` para `.form.txt` reescribe el texto pero no re-vincula el code-module al form instance → `Me.X` undefined al compile

#### Síntoma verificado

En consumer `gestion_riesgos` (DysTelefonica/GESTION_RIESGOS, branch `staging`), tras bulk-import de los 63 `.form.txt`:

1. El `.form.txt` del binary exportado vía `dysflow.export_all({destinationRoot, diff:false, prune:false})` SÍ contiene la línea `Name ="CorreoRAC"` (línea 692), `Name ="ListaRiesgosOferta"` (línea 886), `Name ="DisparadorDelPlan"` — todos los controles presentes.
2. `dysflow.list_objects({projectId})` reporta el form con el nombre correcto (`FormPublicacionCalidadPublicarEjecutar`, `FormGestionRiesgosRiesgosOferta`, `FormRiesgosGestionPlanPrincipal`).
3. El user abre Access, `Debug → Compile VBA Project`. El compile falla con `No se encontró el método o el dato miembro` apuntando a `Me.ControlName` en el code-module del form.
4. Tres forms confirmados rotos en este round (consumer verificado 2026-07-13):
   - `Form_FormGestionRiesgosRiesgosOferta` → binary `FormGestionRiesgosRiesgosRferta` — error: `Me.ListaRiesgosOferta`
   - `Form_FormPublicacionCalidadPublicarEjecutar` → binary `FormPublicacionCalidadPublicarEjecutar` — error: `Me.CorreoRAC`
   - `Form_FormRiesgosGestionPlanPrincipal` → binary `FormRiesgosGestionPlanPrincipal` — error: `Me.DisparadorDelPlan`
5. El source `.form.txt` correspondiente a cada uno TIENE el control definido correctamente (verificado por grep con `^\s*Name\s*=\s*"<ControlName>"`). No es un bug del source — es un bug del import.

#### Evidencia de repro

Comando exacto del consumer (3 forms confirmados, escalable al resto):

```js
// 1) Bulk import del source
await tools.dysflow.import_modules({
  projectId: "00-gestion-riesgos-staging",
  moduleNames: ["Form_FormPublicacionCalidadPublicarEjecutar"],
  dryRun: false,
  importMode: "auto"
});
// Result: {"module":"Form_FormPublicacionCalidadPublicarEjecutar","status":"ok",...}

// 2) Verificar que el control está en el .form.txt del binary
//    (exportar primero)
await tools.dysflow.export_all({
  projectId: "00-gestion-riesgos-staging",
  destinationRoot: "C:/safe/export/path",
  diff: false, prune: false
});
// grep: el .form.txt contiene `Name ="CorreoRAC"` (línea 692).
//       El form existe en list_objects() como FormPublicacionCalidadPublicarEjecutar.

// 3) User: open Access, Debug → Compile VBA Project
//    Error: "No se encontró el método o el dato miembro"
//    Highlighted line: If Nz(Me.CorreoRAC, "") = "" Then
```

Diagnóstico instrumental vía `dysflow.export_all`:
- Source SHA256 `Form_FormPublicacionCalidadPublicarEjecutar.form.txt`: `76F07A8C7F7B40C98AC7BABDB89CE0C34A980417F7A249E3D3015492E4FCEB5D`
- Binary post-bulk-import SHA256: `75B2200C8BAA56A4010AF83BF08ED1510B4B31A5D216F0A1D5889F19220628D2`
- Match: False. Access re-serializa con su propio formato (CR/LF, GUID, RecSrcDt), por lo que SHA match no es buen check.
- Control count `^\s*Name\s*=\s*"` en source: **22**. Control count en binary: **22**. Match. Entonces el control ESTÁ en ambos archivos.

Lo que difiere entre source y binary es **estructura de runtime**, no contenido textual. La hipótesis es que Access preserva el `.form.txt` text pero pierde la asociación `.cls` ↔ form instance cuando recrea el form vía import.

#### Workaround verificado (consumer-side)

```js
// Delete + reimport establece la asociación correctamente.
await tools.dysflow.delete_module({
  projectId: "00-gestion-riesgos-staging",
  moduleNames: ["FormPublicacionCalidadPublicarEjecutar"],  // binary name (sin Form_)
  dryRun: false
});
// status: "ok"

await tools.dysflow.import_modules({
  projectId: "00-gestion-riesgos-staging",
  moduleNames: ["Form_FormPublicacionCalidadPublicarEjecutar"],  // source name (con Form_)
  dryRun: false,
  importMode: "auto"
});
// status: "ok"
```

Verificación post-fix: user compila, el `Me.CorreoRAC` se resuelve, no más error. Funciona en los 3 forms probados.

#### Diagnóstico preliminar (NO verificado internamente — confirmar o descartar)

Hipótesis: el import lee el `.form.txt`, hace `LoadFromText` para crear/reemplazar el form, y guarda con `SaveAsText`. El proceso de Access re-asigna automáticamente el code-module al form instance — pero **solo cuando el form es NUEVO**. Cuando el form YA EXISTÍA en el binary, el `LoadFromText` puede:
1. Reemplazar el contenido del form (controles, layout, etc.) sin tocar las propiedades `Module`/`HasModule` del form object.
2. El form queda con el nuevo contenido visual pero el code-module apunta a la versión anterior del form (que ya no existe), generando una instancia fantasma.

Causa raíz alternativa: `SaveAsText` del import escribe un `.form.txt` con la sección `CodeBehindForm` y los atributos correctos, pero Access interpreta el `Attribute VB_Name` como nombre de un NUEVO módulo en lugar de vincularlo al form existente.

Ambas hipótesis predicen el mismo síntoma: `Me.X` resuelve contra el form instance del pre-import (sin los nuevos controles), no contra el nuevo contenido.

#### Riesgo

- Cuantificación en este consumer: **3 de 46 forms importados en bulk** (6.5%) confirmados con compile error. Estimación: probablemente 5-15% de forms con `Me.X` references que referencian controles actualizados.
- Cross-fleet: otros 10+ consumers de dysflow (no_conformidades, condor, cadete, brass, hps, etc.) probablemente afectados si han hecho bulk imports de forms con código-behind. La regla "human compiles" significa que cada consumer descubre el bug al compilar manualmente → discovery cost per consumer es alto.
- Workaround viable pero destructivo: `delete_module` + `import_modules` por form → ~3s por form × N forms. Aceptable para N=3, prohibitivo para N=200.
- Workaround NO viable: pre-import `delete_module` de TODOS los forms → pierde estado de runtime (RecSrcDt, GUID, layout positions, theme settings).

#### Tests RED sugeridos

```ts
// Test 1: round-trip de un form con code-module preserva Me.X resolvable
describe('import_modules: form .cls link preservation', () => {
  it('preserves form.cls linkage so Me.X resolves to controls', async () => {
    // Setup: crear .accdb de prueba con 1 form (FormTest) y 1 control (TextBox1)
    // El code-module del form referencia Me.TextBox1
    await setupTestBinary();
    
    // Act: re-importar el .form.txt con un control ADICIONAL (TextBox2)
    await client.call('import_modules', {
      moduleNames: ['Form_FormTest'],
      dryRun: false,
      importMode: 'auto'
    });
    
    // Assert: el form instance en el binary tiene HasModule = Yes
    // Y el código Me.TextBox1 sigue compilando
    const form = await client.call('get_object', { name: 'FormTest' });
    expect(form.hasModule).toBe(true);
    expect(form.moduleName).toBe('Form_FormTest');
    
    // Y el code module 'Form_FormTest' apunta al form 'FormTest' (NO a fantasma)
    const mod = await client.call('get_module', { name: 'Form_FormTest' });
    expect(mod.parentForm).toBe('FormTest');
  });

  it('delete + import re-establishes the link after bulk import breaks it', async () => {
    // Setup: form ya con Me.X roto (import previo que dejó phantom)
    await simulateBulkImportBreak();
    
    // Act: delete + reimport
    await client.call('delete_module', { moduleNames: ['FormTest'], dryRun: false });
    await client.call('import_modules', { moduleNames: ['Form_FormTest'], dryRun: false });
    
    // Assert: el code-module apunta al nuevo form instance
    const mod = await client.call('get_module', { name: 'Form_FormTest' });
    expect(mod.parentForm).toBe('FormTest');
  });

  it('bulk import of N forms maintains Me.X resolvable for all N', async () => {
    // Setup: 10 forms con Me.X references en un .accdb de prueba
    await setupTenFormBinary();
    
    // Act: bulk import
    await client.call('import_modules', {
      moduleNames: forms.map(f => `Form_${f}`),
      dryRun: false,
      importMode: 'auto'
    });
    
    // Assert: el human-compile gate no encuentra ningún Me.X undefined
    // (esto es integration test — necesita abrir Access real)
    const compileErrors = await runCompileCheck();
    expect(compileErrors).toEqual([]);
  });
});
```

## Disciplina

- TDD estricto (RED → GREEN → REFACTOR). El primer test RED debe fallar con `Me.X undefined` en bulk-import de un solo form; el GREEN debe ser la fix en `import_modules` que preserva el link.
- Conventional commits con scope `import-modules` o `access-vba`.
- NO tocar las herramientas/capabilities de rounds anteriores (citá cuáles en "Lo que YA funciona").
- NO reintroducir `compile_vba` (regla cross-project).
- Cross-consumer safety: si la fix cambia el comportamiento de `import_modules` para forms que ya funcionan, marcarlo explícitamente en el changelog.

## Acceptance output

- PR con **3 tests verdes** (los tres tests RED sugeridos arriba).
- Changelog en `CHANGELOG.md` con bullet: `fix(import_modules): preserve form .cls linkage on form re-import so Me.X resolves to current controls (#<issue>)`.
- Version bump: **patch** (`v2.9.3`) si la fix es interna al import; **minor** (`v2.10.0`) si cambia el comportamiento observable de `import_modules` para forms que ya funcionaban.
- Documentación actualizada: si la fix toca `assets/write-flags-matrix.md` o `references/error-codes.md`, sincronizar con el runtime vía `assets/scripts/verify-examples-vs-runtime.ps1` (debe pasar exit 0).
- Si la fix requiere un nuevo code-path en `dysflow-usage` skill (consumer-side), coordinar conmigo ANTES de publicar el changelog — el consumer necesita actualizar la skill `dysflow-usage` simultáneamente.

## Quick start

```bash
git clone <repo path>
cd <repo>
git checkout -b fix/import-modules-preserve-form-cls-link
<comando install>
<comando test>
```

Test repro contra el dev:
```bash
# 1) Crear .accdb de prueba con un form que referencia Me.X
# 2) Editar el .form.txt agregando un control nuevo
# 3) Importar con dysflow
node -e 'import("./dist/index.js").then(m => m.client.call("import_modules", {moduleNames: ["Form_FormTest"], dryRun: false, importMode: "auto"}))'
# 4) Compilar en Access (manual)
# Esperado: 0 errores de compile
# Actual: "No se encontró el método o el dato miembro" en Me.X
```

## Reinforcement

La regla cross-project "human compiles" significa que el consumer descubre este tipo de bug solo cuando compila en Access. Costo de discovery es alto. La fix debe venir con un **test RED automatizable** que el CI pueda correr sin abrir Access (vía un harness que simule la asociación form↔cls con mocks del COM object). Si la fix requiere Access real para validar, marcalo explícitamente como gap de testing y sugerí la estrategia de mock más cercana.