# Prompt para la IA mantenedora de dysflow — Round 3 (MCP runtime, post 1.15.6)

> **Origen**: sesión 2026-07-06 con `00_GESTION_RIESGOS_staging`, proyecto `00-gestion-riesgos-staging`. Round de **verificación post-release** + nuevos hallazgos.
>
> **Modo**: bug-hunt + auditoría de contrato MCP.
>
> **Severidad global**: P0 / bloqueante. La fricción con `test_vba`/`run_vba` PERSISTE en 1.15.6.

---

Eres la IA mantenedora de **dysflow**. Repo: `C:\Proyectos\dysflow`. Branch de trabajo: `main`. Versión verificada en el campo: **1.15.6** (runtime actualizado automáticamente desde 1.15.4 durante la sesión del consumidor).

## Contexto de la ronda

Esta ronda NO es continuación automática de round-1 ni round-2. Es una **verificación de subsanación** post-release de `dysflow 1.15.6` (auto-update detectado en el snapshot vivo). El round-3 incluye hallazgos nuevos que no estaban en rounds previos y referencia los issues ya abiertos para evitar duplicación.

**Tabla de subsanación verificada en 1.15.6** (commit del consumidor: `00_GESTION_RIESGOS_staging`):

| Hallazgo | Prompt original | Estado en 1.15.6 |
|---|---|---|
| `dryRunDefault:false` cuando writes habilitados | round-1 Item 2 | ✅ **SUBSANADO** (ahora `true`) |
| `compact_repair` con `dryRun:true` mutaba el .accdb | round-1 Item 1 | ✅ **SUBSANADO** (md5 ALIGNED post-dry-run) |
| `delete_module` per-entry error reporting | round-1 Item 3 | ❓ No verificable sin `test_vba` funcional |
| `projectIdResolution: unresolved` | round-2 Item 3 | ❌ **NO subsanado** + ver Issue #718 (mismo síntoma) |
| `test_vba` gate con `allowedProcedures` válido | round-2 Item 1 | ❌ **NO subsanado** (mismo `MCP_INPUT_INVALID`) |
| `test_vba` no acepta `dryRun:true` | round-2 Item 1 sub | ❌ **NO subsanado** (sigue `dryRun is not allowed`) |
| `run_vba` con `dryRun:true` → `OpenCurrentDatabase failed` | round-2 Item 2 | ❌ **NO subsanado** (persiste, falso positivo) |
| **NUEVO**: `doctor({ includeEnvironment: true })` modifica el .accdb | round-3 Item 3 | 🆕 Sin issue previo |

Severidad agregada de los items no subsanados + nuevo: **P0**. El consumidor (`00_GESTION_RIESGOS_staging`) sigue bloqueado para correr TDD porque `test_vba` no funciona.

## Lo que YA funciona (NO tocar)

- **`dysflow_doctor` (CLI)** retorna `access-open: opened` correctamente. La CLI no tiene el bug del MCP.
- **`dysflow_setup` (CLI)** configura correctamente el runtime, apunta al `.accdb` correcto.
- **`dysflow_compact_repair` con `dryRun:true`** ya no muta el binario (round-1 Item 1 subsanado).
- **`dysflow_dysflow_get_capabilities.dryRunDefault: true`** ahora coincide con el AGENTS.md global (round-1 Item 2 subsanado).
- **`dysflow_export_all`, `dysflow_import_modules`**, `dysflow_list_objects`, `dysflow_list_access_files`, etc.: read-class tools funcionan (con los caveats de la regla dysflow-write-serialize).
- **`dysflow_dysflow_form_serialize`, `dysflow_form_add_control`**, etc.: form tools round-trip OK (con caveat de #718).
- **Round-1 y round-2 hallazgos ya documentados** siguen vigentes como contrato esperado.

## Lo que falta en este round

### Item 1 — `projectIdResolution: unresolved` persiste en 1.15.6 (P0)

**Síntoma verificado**: el consumidor configuró `.dysflow/project.json` con `id: "00-gestion-riesgos-staging"` (correspondiente a la rama staging + path local `00_GESTION_RIESGOS_staging`). El id coincide con la rama actual (`b2-punto-15-anexos-helper` basada en staging) y con el `projectId` que pasa en cada llamada. Sin embargo, `dysflow_dysflow_get_capabilities` retorna:

```json
{
  "projectIdResolution": { "projectId": null, "outcome": "unresolved" },
  ...
}
```

**El campo `projectId: null` es la pista**: el MCP ni siquiera está leyendo el `projectId` que el caller pasa, o lo lee pero no lo puede resolver contra el project config.

**Causa probable**: cache obsoleto del MCP runtime tras el auto-update de 1.15.4 a 1.15.6, o bug del resolver que no lee el `.dysflow/project.json` actualizado.

**Relación con issue #718**: ese issue pide que las form tools resuelvan path desde `projectId` (problema similar pero distinto alcance). El problema del consumidor es MÁS amplio: el MCP no resuelve `projectId` para NINGÚN tool, ni siquiera `test_vba` o `doctor`. Si #718 se aprueba y mergea, podría resolver parcialmente este item, pero el alcance debería ampliarse a todos los tools.

**Test RED sugerido** (en `test/adapters/mcp/tools.test.ts` o equivalente):
- Given `.dysflow/project.json` con `id` válido que coincide con el `projectId` que pasa el caller
- When `dysflow_dysflow_get_capabilities` se invoca con `projectId`
- Then `projectIdResolution.outcome === "resolved"`
- And `projectIdResolution.projectId` debe ser el id pasado (no `null`)
- And `projectIdResolution.source` debe ser `"file"` (vs `"default"` o `"cache"`)

### Item 2 — `dysflow_run_vba` con `dryRun:true` reporta bloqueo falso (P0)

**Síntoma verificado** (post-cambio de projectId): el consumidor llamó `dysflow_run_vba` con `procedureName: "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID"`, `argsJson: []`, `dryRun: true`, `projectId: "00-gestion-riesgos-staging"`. La respuesta fue:

```
RUNNER_FAILED: PowerShell runner failed with exit code 1: OpenCurrentDatabase failed for 'Gestion_Riesgos.accdb': Excepción al llamar a "OpenCurrentDatabase" con los argumentos "3": "Microsoft Access no puede abrir la base de datos porque falta o porque está abierta de forma exclusiva por otro usuario, o no es un archivo ADP."
```

**Estado OS verificado** (NO había bloqueo real):
- `Get-Process -Name MSACCESS` → 0 procesos
- `Get-ChildItem . -Filter "*.laccdb"` → 0 archivos
- `dysflow doctor` (CLI) → `✓ access-open: opened` — **la CLI confirma que el .accdb se puede abrir**
- `dysflow setup` → apunta correctamente a `C:\00repos\codigo\00_GESTION_RIESGOS_staging\Gestion_Riesgos.accdb`

**Conclusión**: el MCP runtime no logra abrir el binario aunque la CLI puede. Es un bug del runner MCP, no del sistema.

**Riesgo**: el consumidor deduce bloqueo falso, intenta acciones destructivas (`cleanup_access_operation` con `force`, o peor, `Stop-Process -Name MSACCESS` que está baneado). Empuja al consumidor a soluciones agresivas para un problema que no existe.

**Test RED sugerido**:
- Given un `.accdb` válido sin procesos vivos ni lock files, con `dysflow doctor` (CLI) retornando `access-open: opened`
- When `dysflow_run_vba` con `dryRun:true` se invoca
- Then el runner debe planificar sin tocar Access (no `OpenCurrentDatabase` exception), y devolver el plan JSON

### Item 3 — `dysflow_dysflow_doctor({ includeEnvironment: true })` modifica el .accdb (P1)

**Síntoma verificado**: el consumidor llamó `dysflow_dysflow_doctor` con `projectId: "00-gestion-riesgos-staging"` y `includeEnvironment: true`. Después de la llamada, el md5 del `.accdb` cambió de `C9F1EF02AC77DF4BD2D55F32EE7CD84F` (HEAD) a `FB5BBA598FAD5F5724DE537C940C9691`. **Doctor modificó el binario.**

**Comportamiento esperado**: doctor debería ser **read-only**. La doc del tool dice "Read-only and dry-run" (verificar), pero en 1.15.6 no se cumple cuando `includeEnvironment:true`.

**Riesgo**: cualquier consumidor que use `doctor({ includeEnvironment: true })` para diagnóstico rutinario muta el binario silenciosamente. Si el binario tiene cambios valiosos sin commitear, los pierde. Si el binario está limpio, queda con metadata Access modificada (timestamps, internal stats).

**Test RED sugerido**:
- Given un `.accdb` con md5 conocido
- When `dysflow_dysflow_doctor({ includeEnvironment: true })` se invoca
- Then el md5 del `.accdb` debe permanecer idéntico
- And no deben crearse `.laccdb` files residuales

### Item 4 — `get_capabilities` retorna `projectId: null` cuando unresolved (P2)

**Síntoma verificado**: el snapshot retorna `"projectId": null` (no el id que el caller pasó). Sin información sobre qué id intentó resolver, qué path leyó, o por qué falló.

**Esperado**: el campo `projectId` debe ser el id que el caller pasó (string), no null. Y `projectIdResolution` debería incluir un campo `reason` con la causa del fallo (e.g., "project.json not found", "id mismatch", "cache stale").

**Test RED sugerido**:
- Given cualquier estado
- When `get_capabilities` se invoca con `projectId`
- Then `projectIdResolution.projectId` es el id pasado (string, no null)
- And si `outcome: unresolved`, `projectIdResolution.reason` debe ser uno de `["project.json not found", "id mismatch", "cache stale", "path invalid", "unknown"]`

### Item 5 — `test_vba` y `delete_module` no aceptan `dryRun` (P2)

**Síntoma verificado**:
- `test_vba` con `dryRun:true` → `dryRun is not allowed`
- `delete_module` con `dryRun:true` → `dryRun is not allowed`

**Inconsistencia**: otros write-class tools (`compact_repair`, `import_modules`, `run_vba`) SÍ aceptan `dryRun:true` como escape hatch. Pero `test_vba` y `delete_module` no. Esto rompe el patrón consistente de "dry-run cuando hay duda".

**Riesgo**: el consumidor no puede planear cambios en `delete_module` (debe ejecutar de verdad, con riesgo de side-effects no reportados honestamente, ver round-1 Item 3). El consumidor no puede ver el plan de `test_vba`.

**Test RED sugerido**:
- Given `.dysflow/project.json` con `allowedProcedures` válido
- When `test_vba({ dryRun: true, proceduresJson: [...] })` se invoca
- Then el tool debe retornar el plan JSON sin ejecutar los tests
- And `dysflow_dysflow_access_operations_list` no debe mostrar la operación como `running`

## Disciplina

- **Strict TDD**: test RED primero, código que lo hace pasar, refactor.
- **Conventional commits**: mensajes en inglés, formato `<type>(<scope>): <subject>`, body con referencia al issue y al test.
- **Backward compat**: si el fix de Item 1 introduce `projectIdResolution.source`, el default debe ser `"unknown"` o `"file"` para no romper consumers que no lean el campo nuevo.
- **No reintroducir el bug de round-1**: la subsanación de `dryRunDefault:true` debe preservarse en todos los tools.
- **No agregar más fricción**: el fix del Item 1 (projectIdResolution) debe resolver automáticamente las fricciones de Items 2, 3, 4, 5 (porque todos se derivan de que el MCP no resuelve el project config).

## Acceptance output

1. **PR 1 (P0) — Fix de `projectIdResolution`**:
   - Tests RED primero. Cambios:
     - `src/adapters/mcp/tools.ts` o equivalente: el resolver del `projectId` debe leer el `.dysflow/project.json` actualizado y retornar `outcome: resolved` cuando el id coincide.
     - Invalidar cache obsoleto del MCP al detectar auto-update.
     - `get_capabilities` debe retornar el `projectId` que el caller pasó (no null) más el campo `source: "file" | "cache" | "default"`.
   - Tests E2E que verifiquen: (a) `get_capabilities` retorna `projectIdResolution: resolved` con id válido en `.dysflow/project.json`; (b) tras cambio de id, `resolved` se actualiza sin reiniciar MCP.
   - Version bump: `1.15.6` → `1.16.0` (P0, breaking si se introduce el campo `source`).
   - CHANGELOG entry.

2. **PR 2 (P0) — Fix de `run_vba` con `dryRun:true`**:
   - Tests RED primero. Cambios:
     - El runner PowerShell debe respetar `dryRun:true` sin ejecutar `OpenCurrentDatabase`.
     - La respuesta del plan debe ser JSON con la firma propuesta (o `null` con `dryRun:true` acknowledged).
   - Tests que verifiquen: `run_vba` con `dryRun:true` y binario OK según `dysflow doctor` CLI → plan válido, sin `OpenCurrentDatabase` exception.
   - Version bump: `1.16.0` → `1.16.1`.

3. **PR 3 (P1) — `dysflow_doctor({ includeEnvironment: true })` read-only**:
   - Tests RED primero.
   - Doctor no debe modificar el `.accdb` ni crear `.laccdb` files.
   - Tests que verifiquen: md5 ALIGNED pre/post-doctor.
   - Version bump: `1.16.1` → `1.16.2`.

4. **PR 4 (P2) — Documentación + campos adicionales**:
   - `get_capabilities`: agregar campo `projectIdResolution.reason` cuando `outcome: unresolved`.
   - `test_vba`, `delete_module`: aceptar `dryRun:true` consistente con otros write-class tools.
   - Version bump: `1.16.2` → `1.16.3` (doc + param extension).

5. **Comunicación al consumidor**: cuando los PRs estén mergeados, notificar al consumidor `00_GESTION_RIESGOS_staging` para que retire los workarounds manuales (revertir .accdb pre/post doctor, saltar `test_vba`, etc.).

## Quick start

```bash
# 1. Clonar el repo de trabajo
cd C:\Proyectos\dysflow

# 2. Crear rama de fix
git checkout -b fix/round-3-mcp-runtime-fixes

# 3. Verificar runtime actual
dysflow --version
# esperado: 1.15.6

# 4. Correr test suite para baseline
pnpm install
pnpm test

# 5. Para reproducir Item 1:
# - En un repo cualquiera con .dysflow/project.json con id válido
# - Llamar dysflow_dysflow_get_capabilities({ projectId: "..." })
# - Verificar projectIdResolution.outcome debe ser "resolved" con id correcto
# - En 1.15.6 retorna "unresolved" con projectId: null

# 6. Para reproducir Item 2:
# - Verificar dysflow doctor (CLI) retorna access-open: opened
# - Llamar dysflow_run_vba({ procedureName: "...", dryRun: true, projectId: "..." })
# - En 1.15.6 retorna OpenCurrentDatabase failed (falso positivo)

# 7. Para reproducir Item 3:
# - Capturar md5 del .accdb
# - Llamar dysflow_dysflow_doctor({ includeEnvironment: true, projectId: "..." })
# - Verificar md5 post-call
# - En 1.15.6 cambia (drift)

# 8. Tests RED primero, luego implementar fix, luego verificar.

# 9. Coverage gate: no bajar la cobertura global.
pnpm test --coverage
```

## Queries de verificación que el orquestrador va a correr después del fix

```javascript
// Verificar 1: get_capabilities resuelve projectId
dysflow_dysflow_get_capabilities()
// esperado: projectIdResolution.outcome === "resolved", projectId === "00-gestion-riesgos-staging", source === "file"

// Verificar 2: doctor read-only
const md5_before = md5("Gestion_Riesgos.accdb");
dysflow_dysflow_doctor({ projectId: "00-gestion-riesgos-staging", includeEnvironment: true });
const md5_after = md5("Gestion_Riesgos.accdb");
assert md5_before === md5_after;  // no drift

// Verificar 3: test_vba con allowedProcedures válido
dysflow_test_vba({
  projectId: "00-gestion-riesgos-staging",
  proceduresJson: '["Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos"]'
})
// esperado: ok: true con results por procedure (no MCP_INPUT_INVALID)

// Verificar 4: run_vba con dryRun no toca Access
dysflow_run_vba({
  procedureName: "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID",
  argsJson: "[]",
  dryRun: true,
  projectId: "00-gestion-riesgos-staging"
})
// esperado: plan JSON sin OpenCurrentDatabase exception

// Verificar 5: test_vba con dryRun:true
dysflow_test_vba({
  projectId: "00-gestion-riesgos-staging",
  proceduresJson: '["Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos"]',
  dryRun: true
})
// esperado: plan JSON, no ejecutado
```

## Referencias cruzadas

- **Issue #718** (OPEN, approved): "feat(forms): resolve form source paths from projectId across form tools". Mismo síntoma para form tools, scope más limitado. Si se aprueba y mergea, debe coordinarse con Item 1 de este round (alcance ampliado a todos los tools).
- **Issue #691** (CLOSED): "fix(http): mirror test_vba default-deny allowlist gate". Ya cerrado.
- **Issue #738** (CLOSED): "[Bug] test_vba MCP contract metadata still claims the allowlist gate is pending — PR1b already landed". PR1b cerrado pero el comportamiento del gate persiste en el consumidor.
- **Round-1 prompt** (`prompt-ia-mantenedora-dysflow-round-1-2026-07-06.md`): hallazgos previos (subsanados parcialmente en 1.15.6).
- **Round-2 prompt** (`prompt-ia-mantenedora-dysflow-round-2-2026-07-06.md`): hallazgos sobre gate `test_vba` y modelo multi-proyecto. Items 1, 2, 3 NO subsanados en 1.15.6.

## Lección operativa del consumidor (para el mantenedor)

El consumidor (`00_GESTION_RIESGOS_staging`) durante esta sesión perdió tiempo diagnosticando los P0 de rounds previos en el sitio equivocado (atribuyó al runtime de dysflow bugs que eran del project config local). El consumer aprendió:

1. **Siempre verificar `.dysflow/project.json` `id` antes de empezar a usar dysflow** — si no coincide con la rama/path, corregir ANTES de culpar al runtime.
2. **Usar `dysflow doctor` (CLI) como baseline de "puede el sistema abrir el binario"** — si la CLI dice `access-open: ok` pero el MCP dice `OpenCurrentDatabase failed`, el bug es del MCP, no del sistema.
3. **El snapshot `get_capabilities.projectIdResolution` debe ser el primer chequeo** — si `outcome: unresolved`, no usar NINGÚN tool del MCP hasta arreglar el id del project.json o reiniciar el MCP.

El consumidor sugiere que el MCP exponga un helper `dysflow_dysflow_resolve_project(projectId)` explícito que retorne:

```json
{
  "projectId": "...",
  "outcome": "resolved" | "unresolved",
  "reason": "project.json not found" | "id mismatch" | "cache stale" | "path invalid" | "unknown",
  "accessPath": "...",
  "projectRoot": "...",
  "sourceRoot": "..."
}
```

Esto reduciría el tiempo de diagnóstico del consumidor de ~30 min a ~1 min.

## Notas operativas del consumidor

- El consumidor (`00_GESTION_RIESGOS_staging`) tiene cambios sin commitear en `src/classes/Edicion.cls` (implementación de `Function ColAnexosTotales` siguiendo SOLID, listo para import cuando `test_vba` funcione).
- Working tree limpio a nivel de `.accdb` (md5 ALIGNED con HEAD).
- Reunión con Calidad del cliente programada pronto; los tests del Punto 15 (Opción C) son prerequisito para merge a main.
- 12 issues OPEN en el repo del consumidor; el cierre de varios depende de `test_vba` funcional.

---

**Fin del round 3. Cuando los PRs estén mergeados, abrir round 4 con la auditoría completa de los 35+ write-class tools (round-1 Item 4 deferred) más los nuevos findings del round-3 (si los hay).**