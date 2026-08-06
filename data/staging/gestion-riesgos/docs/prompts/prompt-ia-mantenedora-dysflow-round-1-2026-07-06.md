# Prompt para la IA mantenedora de dysflow — Round 1 (P0)

> **Origen**: sesión 2026-07-06 con `00_GESTION_RIESGOS_staging`, proyecto `00-gestion-riesgos-develop`. Tres hallazgos verificados durante una limpieza rutinaria del binario Access.
>
> **Modo**: bug-hunt (P0).
>
> **Severidad global**: bloqueante — el contrato de dry-run no es confiable, los consumidores no pueden usar el patrón seguro habitual.

---

Eres la IA mantenedora de **dysflow**. Repo: `C:\Proyectos\dysflow`. Branch de trabajo: `main` (o tu rama de release habitual). Versión actual del runtime en el consumidor: **1.15.4** (adapter `1.15.4`, paquete `@telefonica/dysflow` MCP stdio).

## Contexto de la ronda

Un consumidor (orquestador opencode sobre proyecto Access/VBA) ejecutó una limpieza rutinaria del binario (`Gestion_Riesgos.accdb`) usando herramientas write-side de dysflow. Detectó que **el contrato `dryRun:true` no se respeta de forma fiable**: una llamada explícita con `dryRun:true` modificó el contenido del binario a pesar de pedir plan. El resultado fue revertido con `git checkout`, pero el patrón seguro de uso (revisar plan → aplicar) está roto. Tres hallazgos relacionados:

1. **P0 — `compact_repair` ignora `dryRun:true` cuando el proyecto tiene `writesProject.allowWrites:true` en `.dysflow/project.json`**. El md5 del binario cambió durante el dry-run (no durante el apply), aunque se pasó el flag explícitamente.
2. **P0 — `dryRunDefault` documentado como `true` en el AGENTS.md global y en tu propio CHANGELOG, pero el runtime retorna `false` en `dysflow_dysflow_get_capabilities()` cuando el proyecto tiene `writesProject.allowWrites:true`**. Esto rompe el contrato esperado por los consumidores que leen el AGENTS.md al inicio de la sesión.
3. **P1 — `delete_module` reporta `status:"error"` con mensaje "No existe objeto/componente para eliminar" para ciertos objetos (caso verificado: `TempSccObj1..6`), pero un `list_objects` posterior confirma que SÍ fueron borrados**. El reporte de éxito/error por entry no es fiable cuando la eliminación ocurre como side-effect de otra (probablemente al borrar `__dysflow_inline__`).

Severidad agregada: **P0 / bloqueante** para los dos primeros. Cualquier consumidor que asuma que `dryRun:true` es seguro perderá datos o corromperá el binario silenciosamente.

## Lo que YA funciona (NO tocar)

- Read-side tools respetan sus contratos: `verify_code`, `list_objects`, `get_schema`, `query_sql` (modo read), `dysflow_get_capabilities`, `vba_orphan_audit` (read-only), `dysflow_access_operations_list`, `dysflow_access_force_cleanup_orphaned` (modo listing).
- `import_modules` con `dryRun:true` parece respetar el flag según la nota del CHANGELOG ("standardized dryRun defaults" en v1.14). Esto NO es parte del round, solo se documenta como punto de referencia.
- El snapshot `dysflow_get_capabilities` es estructuralmente correcto (`writesProcess`, `writesProject`, `toolsVisible`, `writeClassToolsPermitted`). Solo el valor de `dryRunDefault` es engañoso.
- El write gate (`MCP_WRITES_DISABLED` cuando `writesProcess.enabled:false`) sí funciona — el bug es específicamente en el flag `dryRun` cuando writes están habilitadas.

## Lo que falta en este round

### Item 1 — `compact_repair` con `dryRun:true` no debe mutar el binario (P0)

**Síntoma verificado**: el consumidor llamó `dysflow_compact_repair({ projectId: "00-gestion-riesgos-develop", databasePath: "Gestion_Riesgos.accdb", backupFirst: true, dryRun: true })`. El tool retornó el plan correcto (`dryRun:true, sourcePath, targetPath, backupFirst:true, wouldReplaceSource:true`). Pero el md5 del archivo `Gestion_Riesgos.accdb` CAMBIÓ:

| Estado | md5 |
|---|---|
| Pre-call (HEAD) | `C9F1EF02AC77DF4BD2D55F32EE7CD84F` |
| Post-call (después del dry-run) | `5E463DAD5AA0CEB707566046D256075F` |

Mismo tamaño (50,708,480 bytes), contenido distinto. El consumidor revirtió con `git checkout -- Gestion_Riesgos.accdb` y recuperó el HEAD.

**Evidencia de repro**: `.dysflow/project.json` del consumidor tiene `"writesProject": { "allowWrites": true }` y `"writesProcess": { "enabled": true }` (snapshot vivo). El bug solo se reproduce cuando ambos flags están activos.

**Riesgo**: cualquier consumidor que use el patrón "plan primero, apply después" verá la revisión del plan (que parece correcto) pero el binario ya estará mutado. La fase "apply" se ejecuta sobre un estado alterado no intencionalmente.

**Test RED sugerido** (en `test/core/runner/compact-repair.test.ts` o equivalente):
- Given un `.accdb` con hash conocido H, `writesProject.allowWrites:true`, `writesProcess.enabled:true`.
- When `compact_repair({ dryRun: true })` se ejecuta.
- Then el hash del archivo debe seguir siendo H (no debe cambiar), y la respuesta debe ser la metadata del plan sin haber escrito nada en el filesystem del binario.

### Item 2 — `dryRunDefault` debe ser `true` por contrato (P0)

**Síntoma verificado**: `dysflow_dysflow_get_capabilities()` retorna actualmente:
```json
{ "dryRunDefault": false, "writesProject": { "allowWrites": true }, "writesProcess": { "enabled": true } }
```
Pero el AGENTS.md global opencode (línea 781) y el CHANGELOG de dysflow (release con "standardized dryRun defaults") dicen que el default debe ser `true`. El consumidor asume `true` al inicio de la sesión, lee la herramienta con intención "plan primero, apply después" y termina ejecutando mutaciones no intencionales.

**Evidencia de repro**: cualquier sesión dysflow con writes habilitadas a nivel proceso Y proyecto. El snapshot siempre retorna `dryRunDefault:false`. El AGENTS.md nunca cambia.

**Riesgo**: comportamiento "fail open" silencioso. Consumidores que asumen safe-by-default ejecutan escrituras destructivas sin su intención explícita. En el caso de la sesión que disparó este round, el `delete_module` que asumió dry-run terminó borrando `__dysflow_inline__` del binario (que era huérfano real, así que el resultado fue correcto por suerte, pero fue accidente).

**Test RED sugerido** (en `test/adapters/mcp/capabilities.test.ts` o equivalente):
- Given un proyecto con `writesProject.allowWrites:true`.
- When el consumidor llama `dysflow_dysflow_get_capabilities`.
- Then `dryRunDefault` debe ser `true` (NO `false`), para que el AGENTS.md y el snapshot sean coherentes.

### Item 3 — `delete_module` per-entry error status debe reflejar el resultado real (P1)

**Síntoma verificado**: el consumidor llamó `delete_module` con `moduleNames: ["__dysflow_inline__", "TempSccObj1", "TempSccObj2", "TempSccObj3", "TempSccObj4", "TempSccObj5", "TempSccObj6"]`. La respuesta fue:
```json
[
  { "module": "__dysflow_inline__", "status": "ok", "deleted": "__dysflow_inline__", "kind": "VBComponent", "tempSccObjectsCleaned": [] },
  { "module": "TempSccObj1", "status": "error", "error": "No existe objeto/componente para eliminar: TempSccObj1" },
  { "module": "TempSccObj2", "status": "error", "error": "..." },
  { "module": "TempSccObj3", "status": "error", "error": "..." },
  { "module": "TempSccObj4", "status": "error", "error": "..." },
  { "module": "TempSccObj5", "status": "error", "error": "..." },
  { "module": "TempSccObj6", "status": "error", "error": "..." }
]
```

Un `list_objects` posterior confirmó que los seis `TempSccObj1..6` SÍ fueron eliminados del binario. El status por entry miente.

**Riesgo**: consumidores que reaccionan al `status:"error"` ejecutando reintentos o rollback terminan borrando cosas dos veces o borrando cosas que no querían.

**Test RED sugerido**:
- Given un binario con seis `TempSccObj*` huérfanos registrados en `list_objects`.
- When `delete_module` se invoca con la lista completa.
- Then todas las entries de la respuesta deben tener `status:"ok"` (o un estado que refleje la realidad post-eliminación), y un `list_objects` posterior debe confirmar la eliminación.

### Item 4 — Auditoría completa de write-side tools (P1)

**Solicitado**: extender la auditoría a TODOS los write-class tools listados en `dysflow_dysflow_get_capabilities().writeClassToolsPermitted` (35+ herramientas). Para cada uno, verificar si respeta `dryRun:true` cuando `writesProject.allowWrites:true`. Resultado esperado: tabla con columnas `[tool, dryRun_respetado, evidencia]`. Cualquier herramienta que no respete el flag se considera otro ítem de fix.

**Test RED sugerido**: harness parametrizado que ejecuta cada write-side tool con `dryRun:true` contra un fixture con hash conocido y verifica invariantes (md5, listas de objetos).

## Disciplina

- **Strict TDD**: test RED primero, código que lo hace pasar, refactor. No parchar el binario del consumidor ni del mantenedor sin test que lo justifique.
- **Conventional commits**: mensajes en inglés, formato `<type>(<scope>): <subject>`, body con referencia al issue y al test.
- **Contrato primero**: define el comportamiento esperado de `dryRun:true` (read-only real, sin side-effects en el binario) en un test que falle antes de tocar código.
- **No reintroducir `DRY_RUN_DEFAULT:` warning** sin actualizar también el AGENTS.md global opencode (que está fuera de tu repo, pero el comportamiento debe coincidir).
- **Compatibilidad**: si el cambio requiere que los consumidores pasen `apply:true` explícitamente para escribir, considera un `deprecationWarning` antes del breaking change. Si decides hacer el fix como "ahora hay que pasar `apply:true`", asegúrate de que `compact_repair` con `dryRun:true` realmente no haga nada (el plan debe ser 100% no-mutating).

## Acceptance output

1. **PR 1 (P0)** — Arreglo de `compact_repair` + `dryRunDefault`. Tests RED primero. Cambios:
   - `src/adapters/mcp/tools.ts` o equivalente: cuando `dryRun:true` se pasa explícitamente, NO debe ejecutar mutaciones en el binario, independientemente del estado de writes.
   - `src/adapters/mcp/handlers/capabilities.ts` o equivalente: `dryRunDefault` debe ser `true` siempre (no flip por `allowWrites`).
   - Tests E2E que verifiquen: `compact_repair({dryRun:true})` no cambia md5 del binario, y `get_capabilities` retorna `dryRunDefault:true` incluso con writes habilitadas.
   - Version bump: `1.15.4` → `1.16.0` (P0, breaking implícito del contrato).
   - CHANGELOG entry que refleje el cambio de contrato.

2. **PR 2 (P1)** — Arreglo del per-entry error reporting en `delete_module`. Tests RED primero.
   - El `status` por entry debe reflejar el resultado real (objeto eliminado o no).
   - El campo `tempSccObjectsCleaned` debe ser preciso o eliminarse si no es útil.
   - Version bump: `1.16.0` → `1.16.1` (P1, no breaking).

3. **Tabla de auditoría (P1)** — entregable aparte (puede ser un markdown en `docs/audit/write-side-dry-run-2026-07-06.md`) con el resultado de auditar todos los write-side tools. Para cada herramienta: `dryRun_respetado: bool`, `evidencia: path_test_o_repro`, `acción: mantener/arreglar/documentar`.

4. **Comunicación al consumidor**: cuando los PRs estén mergeados, notificar a la IA del proyecto consumidor (`00_GESTION_RIESGOS_staging`) para que retire el snapshot pre/post manual que tuvo que aplicar como workaround.

## Quick start

```bash
# 1. Clonar el repo de trabajo si no lo tienes
cd C:\Proyectos\dysflow

# 2. Crear rama de fix
git checkout -b fix/round-1-dryrun-contract-p0

# 3. Verificar runtime actual
dysflow --version
# esperado: 1.15.4

# 4. Correr test suite para baseline
pnpm install
pnpm test

# 5. Para reproducir el bug P0 (sin abrir un binario real del consumidor):
# - Crear un .accdb fixture (test fixture o AAAA_test_compact_dryrun.accdb)
# - Capturar md5 antes
# - Llamar compact_repair con dryRun:true vía dysflow MCP
# - Verificar md5 después
# - Si md5 cambió, reproducir confirmado

# 6. Tests RED primero: escribir tests que fallen con el comportamiento actual,
# luego implementar el fix para que pasen.

# 7. Coverage gate: no bajar la cobertura global del repo.
pnpm test --coverage
```

## Queries de verificación que el orquestrador va a correr después del fix

```javascript
// Verificar 1: dryRunDefault debe ser true
dysflow_dysflow_get_capabilities()
// esperado: { "dryRunDefault": true, "writesProject": { "allowWrites": true }, ... }

// Verificar 2: compact_repair dry-run no muta el binario
const before = md5("Gestion_Riesgos.accdb");
dysflow_compact_repair({ projectId, databasePath, backupFirst: true, dryRun: true });
const after = md5("Gestion_Riesgos.accdb");
assert before === after;

// Verificar 3: delete_module per-entry status refleja realidad
const beforeList = dysflow_list_objects({ projectId });
const response = dysflow_delete_module({ projectId, moduleNames: ["TempSccObj1", "..."] });
const afterList = dysflow_list_objects({ projectId });
// Para cada TempSccObj en el input, debe aparecer status:"ok" en response Y
// no aparecer en afterList.
```

## Referencias cruzadas

- **AGENTS.md global opencode** (`~/.config/opencode/AGENTS.md`): línea 781 documenta `dryRunDefault:true`. El fix debe alinearse con esta promesa.
- **dysflow CHANGELOG.md**: release v1.14 "Standardized dryRun defaults. Writing tools (`import_modules`, `import_all`, and `generateForm`) now consistently default to plan mode (`dryRun: true`) unless `apply === true` or `dryRun === false` is explicitly supplied." El fix debe respetar esta promesa.
- **dysflow `AGENTS.md`** (`C:\Proyectos\dysflow\AGENTS.md`): línea 174 dice "Use `apply: true` only for intentional writes after reviewing the dry-run plan." El fix debe respetar este flujo.

## Notas operativas

- El proyecto consumidor (`00_GESTION_RIESGOS_staging`) revertirá con `git checkout -- Gestion_Riesgos.accdb` cualquier cambio no intencional hasta que el fix esté mergeado y desplegado.
- El consumidor NO va a promover el `delete_module` accidental de `__dysflow_inline__` a main porque es un huérfano real, pero NO recomendamos que el comportamiento se repita.
- Reunión con Calidad del cliente está programada pronto; el fix debería estar mergeado antes de cualquier promoción a main.

---

**Fin del round 1. Cuando los PRs estén mergeados, abrir round 2 con las herramientas que la auditoría marque como needing fix.**