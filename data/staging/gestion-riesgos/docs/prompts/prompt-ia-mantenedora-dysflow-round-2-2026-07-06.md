# Prompt para la IA mantenedora de dysflow — Round 2 (P0 + diseño)

> **Origen**: sesión 2026-07-06 con `00_GESTION_RIESGOS_staging`, proyecto `00-gestion-riesgos-develop`. Sesión de **diseño fundamental** sobre el modelo de coexistencia de múltiples proyectos, no un fix puntual.
>
> **Modo**: bug-hunt + arquitectura.
>
> **Severidad global**: P0 / bloqueante. La fricción bloqueó trabajo válido del consumidor.

---

Eres la IA mantenedora de **dysflow**. Repo: `C:\Proyectos\dysflow`. Branch de trabajo: `main` (o tu rama de release). Versión actual del runtime: **1.15.4**.

## Contexto de la ronda

Esta ronda NO es continuación mecánica de round-1 (prompt separado `prompt-ia-mantenedora-dysflow-round-1-2026-07-06.md`). Es una ronda de **diseño fundamental** sobre cómo dysflow maneja la coexistencia de múltiples proyectos que apuntan a distintos binarios Access. Tres fricciones descubiertas en una sesión que NO son fixes puntuales sino síntomas de un modelo mal planteado:

1. **P0 — `test_vba` rechaza TODOS los planes con `MCP_INPUT_INVALID`** aunque `allowedProcedures` esté configurado en `.dysflow/project.json` y los procedures del plan estén en la lista. El gate exige `dryRun:true` como escape hatch, pero `test_vba` rechaza `dryRun` con `dryRun is not allowed`. Resultado: el consumidor queda bloqueado, sin camino de salida documentado.
2. **P0 — `dysflow_run_vba` con `dryRun:true` falla con `OpenCurrentDatabase failed`** aunque el binario NO está realmente bloqueado. Estado OS verificado: `MSACCESS.EXE` count = 0, `*.laccdb` files = 0, no hay procesos vivos con handle sobre el archivo. La fricción es 100 % runtime de dysflow, no del sistema.
3. **Arquitectura — dysflow fue diseñado asumiendo "un binario activo por sesión"** y el consumidor espera "múltiples proyectos paralelos con distintos binarios". El creador de dysflow confirma la expectativa: nunca dos agentes deben tocar el mismo binario Access; pero con distintos binarios la operación no debe generar ruido ni fricciones falsas.

Severidad agregada: P0 para los tres. Sin un fix, el consumidor queda bloqueado y la IA no puede usar `test_vba` ni `run_vba` en modo lectura ni con `dryRun`.

## Lo que YA funciona (NO tocar)

- **Write-gate `MCP_WRITES_DISABLED`**: bloquea writes cuando `writesProcess.enabled:false`. Eso sí funciona correctamente.
- **`force_cleanup_orphaned` (read-only con `confirmPid`)**: lista MSACCESS huérfanos y pwsh workers que tienen el `accessPath`. Verificado: lista vacía cuando no hay bloqueos reales.
- **`dysflow_export_all` / `dysflow_import_modules`**: funcionan correctamente con paths absolutos o relativos.
- **Snapshot de `get_capabilities`**: estructuralmente correcto (`writesProcess`, `writesProject`, `toolsVisible`, `writeClassToolsPermitted`).
- **Round-1 hallazgos (siguen vigentes, ya documentados)**: `dryRunDefault:false` cuando `writesProject.allowWrites:true` (contradice AGENTS.md), `compact_repair` ignora `dryRun:true`, `delete_module` per-entry error status es unreliable.

## Lo que falta en este round

### Item 1 — `test_vba` gate rechaza plans con `allowedProcedures` válido (P0)

**Síntoma verificado**: el consumidor llamó `dysflow_test_vba` con `proceduresJson: ["Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos", ...]` y `projectId: "00-gestion-riesgos-develop"`. El project.json tiene `"allowedProcedures"` con 18 entries, incluyendo los procedures del plan. La respuesta fue:

```
MCP_INPUT_INVALID: Refusing to execute test_vba plan [...]: project config must declare allowedProcedures (with every procedure in the list) OR caller must pass dryRun:true. Set allowedProcedures in .dysflow/project.json to allow these procedures.
```

Mismo error con un solo procedure, con `filter` o con `proceduresJson`. La regla **niega que `allowedProcedures` esté siendo leído** aunque esté físicamente en el project.json.

**Comprobaciones adicionales**:
- `dysflow_dysflow_get_capabilities()` retornó `"projectIdResolution": {"projectId": null, "outcome": "unresolved"}` — el runtime NO está resolviendo el project config del projectId.
- `dysflow_dysflow_doctor` con `projectId` retornó OK para `access-db-path` y `access-open` (verifica solo el binario, no la resolución del project config).
- Probado con `accessPath` absoluto y relativo: mismo error.
- Probado pasar `dryRun:true` como escape hatch: `dryRun is not allowed` (test_vba no soporta dry-run).

**Riesgo**: cualquier consumidor que intente correr tests queda bloqueado. La IA no tiene documentación sobre cómo desbloquearse.

**Test RED sugerido** (en `test/adapters/mcp/tools.test.ts` o equivalente):
- Given un `.dysflow/project.json` con `allowedProcedures` válido conteniendo los procedures del plan
- When `dysflow_test_vba` se invoca con `proceduresJson` cuyos procedures están en `allowedProcedures`
- Then el plan debe ejecutarse (no `MCP_INPUT_INVALID`)
- And si `projectId` se resuelve correctamente, `projectIdResolution.outcome === "resolved"`

### Item 2 — `run_vba` con `dryRun:true` reporta bloqueo falso (P0)

**Síntoma verificado**: el consumidor llamó `dysflow_run_vba` con `procedureName: "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID"`, `argsJson: []`, `dryRun: true`. La respuesta fue:

```
RUNNER_FAILED: PowerShell runner failed with exit code 1: OpenCurrentDatabase failed for 'Gestion_Riesgos.accdb': Excepción al llamar a "OpenCurrentDatabase" con los argumentos "3": "Microsoft Access no puede abrir la base de datos porque falta o porque está abierta de forma exclusiva por otro usuario, o no es un archivo ADP."
```

**Estado OS verificado** (NO había bloqueo real):
- `Get-Process -Name MSACCESS` → 0 procesos
- `Get-ChildItem . -Filter "*.laccdb"` → 0 archivos
- `dysflow_dysflow_access_force_cleanup_orphaned` (read-only) → `[]` (sin huérfanos)
- `handle.exe` (Sysinternals) → no disponible, pero la combinación de 0 procesos + 0 lock files descarta bloqueo real

**Riesgo**: el consumidor deduce falsamente que tiene un bloqueo, intenta acciones destructivas (`cleanup_access_operation` con `force`, o peor, `Stop-Process -Name MSACCESS` que está baneado por el bloque `dysflow-msaccess-cleanup-only`). Es una fricción peligrosa: empuja al consumidor a soluciones agresivas para un problema que no existe.

**Test RED sugerido**:
- Given un `.accdb` válido sin procesos vivos ni lock files
- When `dysflow_run_vba` con `dryRun:true` se invoca
- Then el runner debe planificar sin tocar Access (no `OpenCurrentDatabase`), y devolver el plan JSON o `null` con `dryRun:true` acknowledged
- And `dysflow_dysflow_access_operations_list` no debe mostrar la operación zombie post-call

### Item 3 — Patrón raíz: dysflow asume "un binario activo" (arquitectura)

**Observación del creador** (citado literal):
> "La idea mía como creador de dysflow es que nunca nadie toque el mismo binario access, que se pueda trabajar con varios proyectos con el dysflow pero con distintos binarios, por lo que eso no debería dar ruido a la IA, ya que no afecta realmente."

**Lo que el runtime hace**:
- Mantiene un solo `accessPid` por `projectId` en el registry
- Locks implícitos sobre el binario abierto
- `force_cleanup_orphaned` lista zombies pero la heurística falla cuando el pwsh worker tiene el lock y no es MSACCESS.EXE
- `MCP_INPUT_INVALID` se dispara incluso cuando el project config declara `allowedProcedures` correctamente

**Lo que debería hacer**:
- Reconocer desde el projectId + accessPath absoluto si el binario está realmente abierto o no
- Si NO está abierto, los read-class tools (`test_vba`, `run_vba` con `dryRun`, `vba_inline_execution` con `dryRun`) deben abrir Access *transitoriamente* para el plan, no exigir un binario persistente
- Los errores de "no se puede abrir" deben distinguir entre "bloqueado por otro proceso real" (entonces sí falla) vs "runtime no reconoce que no hay bloqueo" (entonces debe abrir transitoriamente)
- El gate de `allowedProcedures` debe verificar contra el project config REAL leido, no contra un snapshot cacheado o null

**Sugerencia de diseño** (abierta a tu mejor criterio):
- `dysflow_dysflow_doctor` debe incluir un check `project-config-resolution` que retorne el `allowedProcedures` efectivo (o `null` si no resuelve). Si el consumer recibe `null`, sabe que el gate va a fallar antes de intentarlo.
- `test_vba` con `allowedProcedures` válido debe bypassear `OpenCurrentDatabase` en modo dry-run.
- Para `run_vba` con `dryRun:true`: el runner debe planificar sin abrir Access (es un plan, no ejecución).
- Documentar en `AGENTS.md` y `docs/` los escenarios donde el binario NO está bloqueado pero el runtime lo asume.

**Test RED sugerido** (test de integración, no unit):
- Given el consumidor llama `dysflow_dysflow_doctor({ projectId })` en una sesión con `writesProject.allowWrites:true`
- Then el resultado debe incluir `project_config_resolution: { allowedProcedures_resolved: [...], source: "file" | "default" }` — el consumer puede verificar ANTES de llamar `test_vba` que el allowlist está cargado

## Disciplina

- **Strict TDD**: test RED primero, código que lo hace pasar, refactor. No parchar el binario del consumidor ni del mantenedor sin test que lo justifique.
- **Conventional commits**: mensajes en inglés, formato `<type>(<scope>): <subject>`, body con referencia al issue y al test.
- **Contrato primero**: definir el comportamiento esperado (gate se aplica SOLO si `allowedProcedures` está cargado y los procedures NO están en él) en un test que falle antes de tocar código.
- **Compatibilidad**: el fix de los Items 1 y 2 debe preservar el comportamiento de round-1 (que `dryRun:true` sea respetado en write-class tools, que `delete_module` per-entry status sea honesto, etc.).
- **No reintroducir fricciones**: el fix del Item 3 debe eliminar ruido en el camino feliz del consumidor, no agregar validaciones adicionales.

## Acceptance output

1. **PR 1 (P0) — Fix del gate de `test_vba`**. Tests RED primero. Cambios:
   - `src/adapters/mcp/tools.ts` o equivalente: el gate de `test_vba` debe leer `allowedProcedures` desde el project config REAL (no null) cuando se pasa `projectId` válido.
   - `src/core/config/dysflow-config.ts`: agregar método público `getAllowedProcedures(projectId): string[] | null` que retorna la lista efectiva (cargada o null).
   - `dysflow_dysflow_doctor`: agregar check `project-config-resolution` que retorne `{ allowedProcedures_resolved: string[] | null, source: "file" | "default" }`.
   - Tests E2E que verifiquen: (a) con `allowedProcedures` válido + procedures en la lista → `test_vba` ejecuta; (b) con `allowedProcedures` válido + procedures NO en la lista → `MCP_PROCEDURE_NOT_ALLOWED`; (c) sin `allowedProcedures` → `MCP_INPUT_INVALID` con escape hatch `dryRun:true`.
   - **Version bump**: `1.15.4` → `1.16.0` (P0, breaking implícito del contrato de gate).
   - CHANGELOG entry que refleje el cambio de contrato del gate.

2. **PR 2 (P0) — Fix del false-positive de `run_vba` con `dryRun:true`**. Tests RED primero.
   - El PowerShell runner debe respetar `dryRun:true` sin tocar Access ni ejecutar `OpenCurrentDatabase`.
   - La respuesta del plan debe ser un JSON con la firma propuesta (o `null` si no hay plan ejecutable).
   - Tests que verifiquen: `run_vba` con `dryRun:true` y binario no bloqueado → plan válido, sin excepción de OpenCurrentDatabase.
   - Version bump: `1.16.0` → `1.16.1` (P0).

3. **PR 3 (arquitectura) — Documentación del modelo de coexistencia**. Cambios:
   - `AGENTS.md` global (en `~/.config/opencode/AGENTS.md` línea 781): clarificar que `allowedProcedures` solo aplica cuando se resuelve el projectId; si `projectIdResolution.outcome === "unresolved"`, `test_vba` debe retornar `null` plan (no `MCP_INPUT_INVALID`).
   - `docs/architecture/multi-project-coexistence.md`: documento nuevo explicando el modelo de diseño (un binario por sesión, múltiples proyectos en paralelo sin ruido).
   - `docs/security/adapter-write-gates.md`: ampliar la sección sobre `test_vba` con los tres escenarios del Item 1.
   - Version bump: `1.16.1` → `1.16.2` (doc only, no breaking).

4. **Comunicación al consumidor**: cuando los PRs estén mergeados, notificar a la IA del proyecto consumidor para que retire el workaround manual que tuvo que aplicar (skip `test_vba`, workarounds con `dysflow_run_vba`, etc.).

## Quick start

```bash
# 1. Clonar el repo de trabajo si no lo tienes
cd C:\Proyectos\dysflow

# 2. Crear rama de fix
git checkout -b fix/round-2-test-vba-gate-p0

# 3. Verificar runtime actual
dysflow --version
# esperado: 1.15.4

# 4. Correr test suite para baseline
pnpm install
pnpm test

# 5. Para reproducir el bug P0 (sin abrir un binario del consumidor):
# - Levantar un test harness con un .accdb mínimo + .dysflow/project.json con allowedProcedures
# - Llamar dysflow_test_vba con proceduresJson de procedures válidos
# - Verificar que retorna MCP_INPUT_INVALID (reproducir el bug)

# 6. Tests RED primero: escribir tests que fallen con el comportamiento actual,
# luego implementar el fix para que pasen.

# 7. Coverage gate: no bajar la cobertura global del repo.
pnpm test --coverage
```

## Queries de verificación que el orquestrador va a correr después del fix

```javascript
// Verificar 1: dysflow_dysflow_doctor reporta project-config-resolution
dysflow_dysflow_doctor({ projectId: "00-gestion-riesgos-develop", includeEnvironment: false })
// esperado: checks[] incluye "project-config-resolution" con allowedProcedures_resolved = [...]

// Verificar 2: test_vba ejecuta cuando procedures están en allowedProcedures
dysflow_test_vba({
  projectId: "00-gestion-riesgos-develop",
  proceduresJson: '["Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos"]'
})
// esperado: ok: true con results por procedure (no MCP_INPUT_INVALID)

// Verificar 3: run_vba con dryRun:true no toca Access
dysflow_run_vba({
  procedureName: "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID",
  argsJson: "[]",
  dryRun: true
})
// esperado: plan JSON sin OpenCurrentDatabase exception
```

## Referencias cruzadas

- **AGENTS.md global opencode** (`~/.config/opencode/AGENTS.md`): línea 781 documenta `dryRunDefault:true` (sigue siendo el piso). El fix del Item 3 debe coordinarse con el bloque `<!-- gentle-ai:dysflow-reference -->`.
- **Round-1 prompt** (`prompt-ia-mantenedora-dysflow-round-1-2026-07-06.md`): hallazgos previos sobre `dryRunDefault`, `compact_repair`, `delete_module`. **Siguen vigentes**.
- **Sistema prompt del proyecto `00_GESTION_RIESGOS_staging`** (`AGENTS.md` local): confirma `projectId: "00-gestion-riesgos-develop"`, `allowWrites: true`. El `allowedProcedures` del project.json fue añadido por PR1 del proyecto (commit `7e8c9b4`).

## Notas operativas del consumidor

- El consumidor (yo, en `00_GESTION_RIESGOS_staging`) **no pudo correr `test_vba` durante la sesión del 2026-07-06**. Quedó bloqueado en el Item 1. Las features del Punto 15 que requieren verificar con TDD no se completaron.
- El consumidor tuvo que `git revert` el binario Access por el false-positive de bloqueo del Item 2 (combined con el `compact_repair` dry-run issue del round-1).
- Reunión con Calidad del cliente programada pronto; el fix debería estar mergeado antes de cualquier promoción a main.
- **No hay workaround oficial** para los P0 del round-2. La única salida es saltar la verificación TDD (que el proyecto no acepta per AGENTS.md local rule 4).

## Cita del creador

> "nunca nadie toque el mismo binario access, que se pueda trabajar con varios proyectos con el dysflow pero con distintos binarios, por lo que eso no debería dar ruido a la IA, ya que no afecta realmente"

Esta es la directriz de diseño para el round-2. El runtime debe alinearse con esta expectativa.

---

**Fin del round 2. Cuando los PRs estén mergeados, abrir round 3 con la auditoría de los 35+ write-class tools (Item 4 del round-1).**