# Migración ModuloCacheIndicadores → ModuloCacheIndicadoresIssue18 — Plan de implementación

> **Para el implementador (humano o IA):**
>
> - **REQUIRED SUB-SKILL #1:** `access-vba-tdd` — patrón de tests VBA (sandbox-safe, fixture-first, JSON contract).
> - **REQUIRED SUB-SKILL #2:** `work-unit-commits` — un commit por función legacy migrada + un commit por bloque de callsites actualizado.
> - **REQUIRED SUB-SKILL #3:** `access-vba-e2e-methodology` — puente TDD⇄UAT (cada átomo TDD debe espejar un escenario UAT-validable).
> - **Reglas del proyecto (no negociables):**
>   - El usuario compila manualmente después de cada `import_modules`/`delete_module`. La IA **NUNCA** usa `compile_vba`.
>   - `accessPath`/`backendPath` siempre desde `.dysflow/project.json`. `projectId = "00-no-conformidades-staging-clean"` (NUNCA `"no_conformidades"`).
>   - Sandbox-safe: usar `ForceLocalBackend` o `m_TestingMode`. Ningún test depende de filas preexistentes.
>   - Schema-first: antes de seed, inspeccionar el schema real de las tablas tocadas.

**Goal:** Migrar las 14 funciones exclusivas del módulo legacy `ModuloCacheIndicadores.bas` a la arquitectura de `ModuloCacheIndicadoresIssue18.bas`, completar el refactor del Issue #18, eliminar el choque de nombres `Cache_Indicadores_SincronizarDesdeAC`, y purgar el módulo legacy. Todo bajo TDD estricto con tests de paridad.

**Architecture:**

- **Source of truth post-migración:** `src/modules/ModuloCacheIndicadoresIssue18.bas` absorbe las 14 funciones exclusivas del legacy + consolida las 7 funciones comunes (eliminando duplicación).
- **Tests de paridad primero:** Escritos ANTES de migrar. Validan el comportamiento actual del legacy. Sirven como "frozen snapshot" — si pasan contra legacy, deben pasar contra la nueva implementación tras la migración.
- **Ciclo TDD por función:** RED (test falla, función nueva no existe) → GREEN (implementar copiando del legacy, ajustes mínimos) → REFACTOR (limpiar duplicación con código Issue18 existente si aplica) → COMMIT (work unit individual).
- **Migración de callsites:** Después de migrar cada función, actualizar callsites uno por uno con `dysflow_test_vba` corriendo entre cambios para detectar regresiones temprano.
- **Borrado del legacy:** Solo cuando TODAS las 14 funciones tienen reemplazo + TODOS los callsites migrados + tests verdes en el manifest completo.

**Tech Stack:** VBA 7.x + DAO + Scripting.Dictionary + JsonConverter, dysflow v1.4.1 MCP, Microsoft Access 2019+, repo en Git con OpenSpec mirror.

---

## Contexto del problema

El 18/06/2026 el commit `276e2bc feat(issue-18): ModuloCacheIndicadoresIssue18 — per-NC sync, AC/AR resolvers, full rebuild, read/filter API` introdujo un nuevo módulo como reemplazo del legacy `ModuloCacheIndicadores.bas`, pero la migración quedó **incompleta**: solo se consolidaron 7 funciones comunes (entre ellas `Cache_Indicadores_SincronizarDesdeAC`, causa del error de compilación "nombre ambiguo"). Quedaron **14 funciones exclusivas del legacy** con **~80 callsites activos en producción** que el nuevo módulo no cubre.

El 19/06/2026, al importar los 15 módulos `missingInBinary` durante la sesión de sync, `ModuloCacheIndicadoresIssue18.bas` entró al binario y se materializó el choque de nombres. La mantenedora de dysflow confirmó en v1.4.1 que el comportamiento del adapter es correcto (`Auto` ahora sincroniza `.cls` + `.form.txt`), pero el bug de duplicación de símbolos es un **problema del proyecto no_conformidades**, no de dysflow.

---

## File Structure

### Archivos a modificar (en `src/`)

| Archivo | Acción | Notas |
|---|---|---|
| `src/modules/ModuloCacheIndicadoresIssue18.bas` | Modificar (crecer) | Absorbe las 14 funciones legacy + consolida las 7 comunes |
| `src/modules/ModuloCacheIndicadores.bas` | BORRAR al final (Fase 5) | Solo cuando todo lo demás esté verde |
| `src/modules/Módulo1.bas` | BORRAR (Fase 6) | Basura automática de Access, encoding corrupto |
| ~30 archivos `.cls`/`.bas` con callsites | Modificar (callsites) | Prefijar con `ModuloCacheIndicadoresIssue18.` para evitar el choque y permitir coexistencia temporal |

### Archivos de tests a crear

| Archivo | Propósito |
|---|---|
| `tests/legacy/Test_LegacyCache_Parity_*.bas` | 14 tests de paridad, uno por función legacy. Corren contra el legacy ANTES de migrar. |
| `tests/migration/Test_Migration_Cache_*.bas` | Tests de regresión post-migración. Validan que la nueva implementación conserva comportamiento. |
| `tests/manifests/test-legacy-parity.json` | Manifest para los 14 tests de paridad (Fase 1) |
| `tests/manifests/test-migration-cache.json` | Manifest para los tests de migración (Fases 3-4) |

### Archivos de docs a actualizar

| Archivo | Acción | Notas |
|---|---|---|
| `docs/superpowers/plans/2026-06-20-cache-legacy-to-issue18-migration.md` | Este archivo | Plan operativo |
| `openspec/changes/cache-legacy-to-issue18-migration/` | (opcional, Fase 0) | Solo si se requiere trazabilidad SDD formal |

---

## Inventario de funciones

### Las 14 funciones legacy exclusivas (migración obligatoria)

Cada una tiene callsites activos en producción. **NO se pueden borrar sin migrar primero.**

| # | Función | Callsites externos | Complejidad estimada |
|---|---|---|---|
| 1 | `Cache_Indicadores_Proyecto` | 18 | Media (devuelve Dictionary con shape específico) |
| 2 | `Cache_Indicadores_Auditoria` | 14 | Media (idem, shape de auditoría) |
| 3 | `Cache_IndicadoresProyectoMaterializado_SincronizarNC` | 8 | Alta (escribe en tabla materializada, side effects DB) |
| 4 | `Cache_IndicadoresProyectoMaterializado_CargarConteos` | 7 | Media (lectura con cache de conteos) |
| 5 | `Cache_Proyecto_EstaCargado` | 6 | Baja (bool simple) |
| 6 | `Cache_InvalidarProyecto` | 5 | Baja (clear/invalidate cache) |
| 7 | `Cache_IndicadoresAuditoriaMaterializado_CargarConteos` | 5 | Media |
| 8 | `Cache_IndicadoresAuditoriaMaterializado_Sincronizar` | 4 | Alta |
| 9 | `Cache_IndicadoresProyectoMaterializado_Sincronizar` | 4 | Alta |
| 10 | `Cache_InvalidarAuditoria` | 3 | Baja |
| 11 | `Cache_InvalidarTodo` | 3 | Baja |
| 12 | `Cache_Auditoria_EstaCargado` | 3 | Baja |
| 5 funciones `Cache_Test_*` | Tests internos del legacy | (no requieren migración — son tests, no API) |

**Total a migrar con callsites en producción: 12 funciones** + 2 sin callsites (`Cache_IndicadoresProyectoMaterializado_SincronizarAC/AR` que se borran sin migrar).

### Las 7 funciones comunes (mismo nombre en ambos módulos)

Estas causan el **"nombre ambiguo"** y deben consolidarse en el Issue18:

| # | Función | Resolución |
|---|---|---|
| 1 | `Cache_Indicadores_SincronizarDesdeAC` | Issue18 gana (mantiene firma + usa TestHelper) |
| 2 | `Cache_Indicadores_SincronizarDesdeAR` | Issue18 gana |
| 3 | `Cache_Indicadores_ResolverNCDesdeAC` | Issue18 gana |
| 4 | `Cache_Indicadores_ResolverNCDesdeAR` | Issue18 gana |
| 5 | `Cache_Indicadores_CargarBucket` | Issue18 gana (si existe) o migrar |
| 6 | `Cache_Indicadores_CargarDetalle` | Idem |
| 7 | `Cache_Indicadores_ReconstruirTodo` | Idem |

### Las 2 funciones legacy sin callsites (borrado directo en Fase 5)

- `Cache_IndicadoresProyectoMaterializado_SincronizarAC` (0 callsites)
- `Cache_IndicadoresProyectoMaterializado_SincronizarAR` (0 callsites)

---

## Fases

### Fase 0: Setup y decisiones previas (15 min)

**Archivos:**
- Modificar: `tests/manifests/tests.vba.json` (registrar nuevos manifests)
- Crear: `openspec/changes/cache-legacy-to-issue18-migration/proposal.md` (opcional)

- [ ] **Step 1: Crear el directorio del plan y los manifests**

```bash
mkdir -p tests/legacy tests/migration
```

- [ ] **Step 2: Verificar dysflow runtime v1.4.1**

Llamar `dysflow_dysflow_doctor` con `projectId="00-no-conformidades-staging-clean"`. Esperado: `dysflowVersion: "1.4.1"`, `adapterVersion: "1.4.1"`. Si reporta otra versión, **STOP** y actualizar con `dysflow update`.

- [ ] **Step 3: Capturar baseline de tests pre-migración**

Llamar `dysflow_test_vba` con `proceduresJson='["Test_RegressionAnchor"]'` o el manifest baseline que ya exista. Anotar el resultado (esperado: todos verdes, posiblemente con 2 fallos pre-existentes del Issue #38 según memoria del proyecto). Guardar el JSON de respuesta en `docs/superpowers/plans/2026-06-20-cache-legacy-baseline.json` como evidencia.

- [ ] **Step 4: Crear rama de feature** (opcional, recomendado)

```bash
git checkout -b feature/cache-legacy-to-issue18-migration
```

Si preferís mantener todo en `staging` directamente, omitir este step.

- [ ] **Step 5: Commit de setup**

```bash
git add tests/legacy tests/migration
git commit -m "chore(tests): scaffold legacy parity + migration manifests"
```

---

### Fase 1: Tests de paridad contra el legacy (~3-4 horas)

**Objetivo:** Para cada una de las 12 funciones legacy con callsites + las 2 sin callsites (total 14), escribir un test que:
- Cree el fixture mínimo necesario (cache config, NC padre, AC hija, etc.)
- Llame a la función legacy
- Verifique el output esperado (sea Dictionary, JSON, bool, o lo que retorne)
- **Pase verde contra el legacy actual** (frozen snapshot del comportamiento)

**Patrón de test de paridad (template):**

```vba
' === tests/legacy/Test_LegacyCache_Parity_<Función>.bas ===
Attribute VB_Name = "Test_LegacyCache_Parity_<Función>"
Option Explicit

' Test parity: <Función>(input) == expected_output_against_legacy
' Written BEFORE migration. Must pass green against legacy code.

Public Function Test_<Función>_HappyPath() As String
    Dim logs As Collection
    Dim p_Error As String
    Dim result As Variant
    Dim expected As Variant

    Set logs = New Collection
    On Error GoTo errores

    ' === ARRANGE: create fixtures (sandbox-safe, ForceLocalBackend) ===
    Dim db As DAO.Database
    Dim ncId As Long, acId As Long
    Set db = getdb(p_Error)  ' getdb must route to local backend in test mode
    If p_Error <> "" Then Err.Raise 1000

    ncId = TestHelper.InsertFixtureNC(db, p_Error)  ' TODO: helper real
    If p_Error <> "" Then Err.Raise 1000

    acId = TestHelper.InsertFixtureAC(db, ncId, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    ' === ACT: call legacy function ===
    result = <Función>(<args>)  ' call against ModuloCacheIndicadores (legacy)

    ' === ASSERT: verify expected behavior ===
    ' TODO: assertions específicas por función (ver tabla de shapes abajo)
    If IsNull(result) Then Err.Raise 1001
    ' ... más asserts ...

    Test_<Función>_HappyPath = TestHelper.BuildJsonOk(logs, "parity_ok")
    Exit Function
errores:
    Test_<Función>_HappyPath = TestHelper.BuildJsonFail("Test_<Función>_HappyPath: " & Err.Description, logs)
End Function

Public Function Test_<Función>_SadPath() As String
    ' === Edge case: input inválido (id<=0, parent not found, etc.) ===
    ' ... debe pasar contra el legacy para validar comportamiento actual ...
End Function
```

**Tabla de shapes esperados por función** (para el `ASSERT`):

| Función | Tipo retorno | Shape esperado |
|---|---|---|
| `Cache_Indicadores_Proyecto(p_IDProyecto)` | `Scripting.Dictionary` | `{ "ok": bool, "value": Dictionary<NC>, "error": string/Null }` |
| `Cache_Indicadores_Auditoria(p_IDAuditoria)` | `Scripting.Dictionary` | `{ "ok": bool, "value": Dictionary<Auditoria>, "error": string/Null }` |
| `Cache_Auditoria_EstaCargado(p_IDAuditoria)` | `Boolean` | true si cache hit, false si no |
| `Cache_Proyecto_EstaCargado(p_IDProyecto)` | `Boolean` | idem |
| `Cache_InvalidarAuditoria(p_IDAuditoria, p_Error)` | `Boolean` | true si OK |
| `Cache_InvalidarProyecto(p_IDProyecto, p_Error)` | `Boolean` | idem |
| `Cache_InvalidarTodo(p_Error)` | `Boolean` | idem |
| `Cache_IndicadoresProyectoMaterializado_CargarConteos(p_IDNC)` | `Dictionary` | conteos por estado/acción |
| `Cache_IndicadoresAuditoriaMaterializado_CargarConteos(p_IDAud)` | `Dictionary` | idem |
| `Cache_IndicadoresProyectoMaterializado_Sincronizar(p_IDNC, p_Error)` | `Boolean` | true si OK |
| `Cache_IndicadoresAuditoriaMaterializado_Sincronizar(p_IDAud, p_Error)` | `Boolean` | idem |
| `Cache_IndicadoresProyectoMaterializado_SincronizarNC(p_IDNC, p_Error)` | `Boolean` | idem |

**Tasks de Fase 1 (12 tasks, una por función):**

- [ ] **Task 1.1:** Escribir `Test_LegacyCache_Parity_Cache_Indicadores_Proyecto.bas` con happy + sad path
- [ ] **Task 1.2:** Escribir `Test_LegacyCache_Parity_Cache_Indicadores_Auditoria.bas`
- [ ] **Task 1.3:** Escribir `Test_LegacyCache_Parity_Cache_Proyecto_EstaCargado.bas`
- [ ] **Task 1.4:** Escribir `Test_LegacyCache_Parity_Cache_Auditoria_EstaCargado.bas`
- [ ] **Task 1.5:** Escribir `Test_LegacyCache_Parity_Cache_InvalidarProyecto.bas`
- [ ] **Task 1.6:** Escribir `Test_LegacyCache_Parity_Cache_InvalidarAuditoria.bas`
- [ ] **Task 1.7:** Escribir `Test_LegacyCache_Parity_Cache_InvalidarTodo.bas`
- [ ] **Task 1.8:** Escribir `Test_LegacyCache_Parity_Cache_IndicadoresProyectoMaterializado_CargarConteos.bas`
- [ ] **Task 1.9:** Escribir `Test_LegacyCache_Parity_Cache_IndicadoresAuditoriaMaterializado_CargarConteos.bas`
- [ ] **Task 1.10:** Escribir `Test_LegacyCache_Parity_Cache_IndicadoresProyectoMaterializado_Sincronizar.bas`
- [ ] **Task 1.11:** Escribir `Test_LegacyCache_Parity_Cache_IndicadoresAuditoriaMaterializado_Sincronizar.bas`
- [ ] **Task 1.12:** Escribir `Test_LegacyCache_Parity_Cache_IndicadoresProyectoMaterializado_SincronizarNC.bas`

Por cada task:

- [ ] **Step A:** Escribir el test siguiendo el patrón (fixtures + act + assert + JSON contract)
- [ ] **Step B:** Llamar `dysflow_test_vba` con `proceduresJson='["Test_<Función>_HappyPath","Test_<Función>_SadPath"]'` para verificar que **pasa verde contra el legacy actual** (es el frozen snapshot — debe pasar)
- [ ] **Step C:** Si el test falla contra el legacy, investigar: ¿hay bug pre-existente? Documentar y NO arreglar todavía (registrar en `docs/superpowers/plans/2026-06-20-cache-parity-findings.md`)
- [ ] **Step D:** Commit work unit:

```bash
git add tests/legacy/Test_LegacyCache_Parity_<Función>.bas
git commit -m "test(legacy-parity): freeze <Función> behavior snapshot"
```

- [ ] **Step E (final de Fase 1):** Crear manifest `tests/manifests/test-legacy-parity.json`:

```json
[
  "Test_Cache_Indicadores_Proyecto_HappyPath",
  "Test_Cache_Indicadores_Proyecto_SadPath",
  "Test_Cache_Indicadores_Auditoria_HappyPath",
  "Test_Cache_Indicadores_Auditoria_SadPath",
  "Test_Cache_Proyecto_EstaCargado_HappyPath",
  "Test_Cache_Auditoria_EstaCargado_HappyPath",
  "Test_Cache_InvalidarProyecto_HappyPath",
  "Test_Cache_InvalidarAuditoria_HappyPath",
  "Test_Cache_InvalidarTodo_HappyPath",
  "Test_Cache_IndicadoresProyectoMaterializado_CargarConteos_HappyPath",
  "Test_Cache_IndicadoresAuditoriaMaterializado_CargarConteos_HappyPath",
  "Test_Cache_IndicadoresProyectoMaterializado_Sincronizar_HappyPath",
  "Test_Cache_IndicadoresAuditoriaMaterializado_Sincronizar_HappyPath",
  "Test_Cache_IndicadoresProyectoMaterializado_SincronizarNC_HappyPath"
]
```

Y commit:

```bash
git add tests/manifests/test-legacy-parity.json
git commit -m "test(legacy-parity): manifest for all 14 legacy cache parity tests"
```

**Criterio de salida Fase 1:** Todos los 14 tests de paridad pasan verdes contra el legacy actual. Si alguno falla por bug pre-existente, documentar y decidir: ¿se arregla el bug antes de migrar (puede romper otros), o se migra con el bug conocido y se arregla después?

---

### Fase 2: Decisión de módulo destino (15 min)

**Objetivo:** Para cada una de las 14 funciones legacy, decidir si va al `ModuloCacheIndicadoresIssue18.bas` (consolidación) o a un nuevo módulo `ModuloCacheIndicadoresLegacyBridge.bas` (capa de compatibilidad temporal).

**Criterio de decisión:**

- **Va al Issue18** si la función es API "moderna" (usa `getdb`, `TestHelper`, retorna JSON/logs) o si ya tiene contraparte directa en Issue18 que solo necesita扩展.
- **Va a LegacyBridge** si la función es legacy-only y no encaja en el patrón de Issue18 (e.g., `Cache_IndicadoresProyectoMaterializado_SincronizarNC` que escribe a tabla materializada específica).

**Tasks:**

- [ ] **Task 2.1:** Listar las contrapartes existentes en Issue18 para cada función legacy (qué hay, qué falta). Output: tabla con la decisión por función.
- [ ] **Task 2.2:** Documentar la decisión en `docs/superpowers/plans/2026-06-20-cache-migration-decisions.md`:

```markdown
# Decisión módulo destino por función

| Función legacy | ¿Issue18 o LegacyBridge? | Razón |
|---|---|---|
| Cache_Indicadores_Proyecto | Issue18 (extender) | API similar ya existe |
| Cache_Indicadores_Auditoria | Issue18 (extender) | Idem |
| ... | ... | ... |
```

- [ ] **Task 2.3:** Si se decide crear LegacyBridge, crear el archivo `src/modules/ModuloCacheIndicadoresLegacyBridge.bas` con header mínimo:

```vba
Attribute VB_Name = "ModuloCacheIndicadoresLegacyBridge"
Option Compare Database
Option Explicit
' Bridge layer for legacy cache functions pending full migration to Issue18.
' Each function here MUST have a parity test in tests/legacy/.
' New code should use ModuloCacheIndicadoresIssue18 directly.
```

Y commit:

```bash
git add src/modules/ModuloCacheIndicadoresLegacyBridge.bas docs/superpowers/plans/2026-06-20-cache-migration-decisions.md
git commit -m "chore(cache): decision on legacy function destinations (Issue18 vs bridge)"
```

**Criterio de salida Fase 2:** Tabla de decisiones firmada (mental o explícitamente). Sin esto, no arrancamos Fase 3.

---

### Fase 3: Migración TDD función por función (~5-6 horas)

**Objetivo:** Para cada una de las 14 funciones legacy:
1. Crear la nueva función en el módulo destino (Issue18 o LegacyBridge)
2. Implementar copiando del legacy (ajustes mínimos para encajar en el nuevo contexto)
3. Test de paridad debe pasar contra la nueva (frozen snapshot del comportamiento)
4. Test de regresión debe pasar (cubrir los side effects del migration)
5. Commit work unit

**Ciclo TDD por función:**

- [ ] **Step A:** Importar el archivo al binario (si es nuevo): `dysflow_import_modules` con `importMode="Auto"`. Usuario compila.
- [ ] **Step B:** Crear la función nueva con la firma exacta del legacy:

```vba
Public Function Cache_<Función>(...) As <TipoRetorno>
    ' Implementación copiada del legacy, ajustada al contexto del módulo destino
    ' ... código idéntico al legacy por ahora ...
End Function
```

- [ ] **Step C:** Llamar `dysflow_test_vba` con `proceduresJson='["Test_LegacyCache_Parity_<Función>_HappyPath","Test_LegacyCache_Parity_<Función>_SadPath"]'`. Esperado: PASA verde (la función nueva tiene la misma lógica).
- [ ] **Step D:** Si falla, comparar la salida del legacy vs la nueva. Ajustar hasta que el parity test pase.
- [ ] **Step E:** Escribir test de regresión en `tests/migration/Test_Migration_Cache_<Función>.bas` que valide los side effects (escritura a tabla materializada, cache hit/miss, etc.).
- [ ] **Step F:** Commit work unit:

```bash
git add src/modules/ModuloCacheIndicadoresIssue18.bas tests/migration/Test_Migration_Cache_<Función>.bas
git commit -m "feat(cache): migrate <Función> to Issue18 (parity-verified)"
```

**Tasks de Fase 3 (14 tasks, una por función):** Orden sugerido de mayor a menor riesgo:

- [ ] **Task 3.1:** Migrar `Cache_IndicadoresProyectoMaterializado_SincronizarNC` (8 callsites, Alta complejidad)
- [ ] **Task 3.2:** Migrar `Cache_Indicadores_Proyecto` (18 callsites, Media)
- [ ] **Task 3.3:** Migrar `Cache_Indicadores_Auditoria` (14 callsites, Media)
- [ ] **Task 3.4:** Migrar `Cache_IndicadoresProyectoMaterializado_CargarConteos` (7 callsites, Media)
- [ ] **Task 3.5:** Migrar `Cache_IndicadoresAuditoriaMaterializado_CargarConteos` (5 callsites, Media)
- [ ] **Task 3.6:** Migrar `Cache_IndicadoresAuditoriaMaterializado_Sincronizar` (4 callsites, Alta)
- [ ] **Task 3.7:** Migrar `Cache_IndicadoresProyectoMaterializado_Sincronizar` (4 callsites, Alta)
- [ ] **Task 3.8:** Migrar `Cache_Proyecto_EstaCargado` (6 callsites, Baja)
- [ ] **Task 3.9:** Migrar `Cache_Auditoria_EstaCargado` (3 callsites, Baja)
- [ ] **Task 3.10:** Migrar `Cache_InvalidarProyecto` (5 callsites, Baja)
- [ ] **Task 3.11:** Migrar `Cache_InvalidarAuditoria` (3 callsites, Baja)
- [ ] **Task 3.12:** Migrar `Cache_InvalidarTodo` (3 callsites, Baja)
- [ ] **Task 3.13:** Consolidar las 7 funciones comunes (eliminar del legacy, dejar solo Issue18)
- [ ] **Task 3.14:** Borrar las 2 funciones legacy sin callsites (`_SincronizarAC`, `_SincronizarAR`) — solo del disco, no afecta el binario hasta Fase 5

**Criterio de salida Fase 3:** Las 14 funciones tienen su reemplazo en el módulo destino + tests de paridad verdes + tests de regresión pasando.

---

### Fase 4: Migrar callsites (~3-4 horas)

**Objetivo:** Actualizar los ~80 callsites en `src/classes/*.cls` y `src/modules/*.bas` que apuntan a las funciones legacy, para que apunten a las nuevas (Issue18 o LegacyBridge).

**Estrategia:** Prefijar con el nombre del módulo destino para eliminar el "nombre ambiguo" de forma permanente. Ejemplo:

```vba
' Antes (en ACAuditoriaOperaciones.cls:229):
Cache_Indicadores_SincronizarDesdeAC CLng(.IdAccionCorrectiva), indicatorSyncErr

' Después:
ModuloCacheIndicadoresIssue18.Cache_Indicadores_SincronizarDesdeAC CLng(.IdAccionCorrectiva), indicatorSyncErr
```

**Criterio:** Esto debe hacerse función por función, en work units, después de que la función esté migrada (Fase 3) y antes de borrar el legacy (Fase 5).

**Tasks de Fase 4 (14 tasks, una por función legacy):**

- [ ] **Task 4.1:** Actualizar callsites de `Cache_IndicadoresProyectoMaterializado_SincronizarNC` (8 callsites)
- [ ] **Task 4.2:** Actualizar callsites de `Cache_Indicadores_Proyecto` (18 callsites)
- ... etc, mismo orden que Fase 3.

Por cada task:

- [ ] **Step A:** `grep` los callsites: `grep -rn "\b<Función>\b" src/classes src/modules | grep -v "ModuloCacheIndicadores\."`
- [ ] **Step B:** Para cada callsite, agregar el prefijo `ModuloCacheIndicadoresIssue18.` (o `ModuloCacheIndicadoresLegacyBridge.` según Fase 2).
- [ ] **Step C:** Llamar `dysflow_test_vba` con el manifest `test-legacy-parity.json` y el `test-migration-cache.json`. Esperado: todos verdes.
- [ ] **Step D:** Llamar `dysflow_test_vba` con `proceduresJson='["Test_RegressionAnchor"]'` (baseline original). Esperado: misma cantidad de fallos pre-existentes que en baseline (no nuevos).
- [ ] **Step E:** Commit work unit:

```bash
git add src/classes src/modules
git commit -m "refactor(callsites): migrate <Función> callsites to Issue18 prefix"
```

**Criterio de salida Fase 4:** Todos los callsites están prefijados, todos los tests siguen verdes, baseline de regresión sin nuevos fallos.

---

### Fase 5: Borrar el legacy (30 min)

**Objetivo:** Eliminar `ModuloCacheIndicadores.bas` del disco y del binario. Solo se hace cuando Fases 1-4 están verdes.

**Tasks:**

- [ ] **Task 5.1:** Borrar del binario: `dysflow_delete_module` con `moduleNames=["ModuloCacheIndicadores"]`, `force=true`. **Usuario compila.**

- [ ] **Task 5.2:** Borrar del disco:

```bash
git rm src/modules/ModuloCacheIndicadores.bas
```

- [ ] **Task 5.3:** `dysflow_test_vba` con manifest completo (`tests.vba.json` + `test-legacy-parity.json` + `test-migration-cache.json`). Esperado: todos verdes.

- [ ] **Task 5.4:** `dysflow_verify_binary` con `diff=true`. Esperado: `ok=true`, sin drift, sin `missingInSource`.

- [ ] **Task 5.5:** Commit:

```bash
git add NoConformidades.accdb  # si verify_binary cambió el binario
git commit -m "feat(cache): purge ModuloCacheIndicadores legacy module"
```

**Criterio de salida Fase 5:** `verify_binary ok=true`, 14 funciones migradas funcionando, tests verdes, binario más chico (~75 KB liberados).

---

### Fase 6: Cleanup final (30 min)

**Objetivo:** Limpiar `Módulo1` (basura automática de Access) + compact_repair + push.

**Tasks:**

- [ ] **Task 6.1:** `dysflow_delete_module` con `moduleNames=["Módulo1"]`, `force=true`. **Usuario compila.**

- [ ] **Task 6.2:** `git rm src/modules/Módulo1.bas`

- [ ] **Task 6.3:** `dysflow_compact_repair` con `backupFirst=true`, `databasePath=NoConformidades.accdb`. Recupera ~4 MB.

- [ ] **Task 6.4:** `dysflow_verify_binary` final. Esperado: `ok=true`, drift cero.

- [ ] **Task 6.5:** Commit:

```bash
git add NoConformidades.accdb src/modules/Módulo1.bas
git commit -m "chore(cleanup): delete Módulo1 + compact-repair binary"
```

- [ ] **Task 6.6 (opcional):** Push a origin/staging (solo con tu OK explícito, regla del proyecto).

- [ ] **Task 6.7:** Actualizar `openspec/changes/cache-legacy-to-issue18-migration/` (si se creó en Fase 0) con el cierre + commits.

- [ ] **Task 6.8:** Cerrar el SDD change:

```bash
mv openspec/changes/cache-legacy-to-issue18-migration openspec/changes/archive/
# Editar archive-report.md con los commits
```

**Criterio de salida Fase 6:** Working tree limpio, binario compacto, todos los tests verdes, listo para promoción a main (que requiere tu OK explícito).

---

## Riesgos conocidos

1. **Side effects no documentados del legacy:** Las funciones `Cache_IndicadoresProyectoMaterializado_Sincronizar*` escriben a tablas materializadas específicas. El test de paridad valida el comportamiento, pero los side effects sobre tablas reales podrían no estar cubiertos. **Mitigación:** usar sandbox (`NoConformidades_Datos.accdb` local con `ForceLocalBackend`).

2. **Dependencias circulares:** Es posible que durante la migración detectemos que alguna función legacy llama a otra que también estamos migrando. **Mitigación:** migrar en orden topológico (las de Baja complejidad primero, que no llaman a las de Alta).

3. **Tests flaky pre-existentes:** El baseline del 18/06 mostró 2 fallos pre-existentes del Issue #38 (fuera de scope). **Mitigación:** documentar el baseline y comparar contra él en cada fase, no contra cero fallos.

4. **Compact-repair puede fallar si hay referencias corruptas:** Las tablas materializadas pueden tener índices huérfanos. **Mitigación:** backup automático con `backupFirst=true`.

5. **El usuario debe compilar 5+ veces:** Después de cada import/delete. Esto es regla del proyecto. **Mitigación:** planificar bloques para minimizar interrupciones.

---

## Métricas de éxito

| Métrica | Target |
|---|---|
| Tests de paridad verdes contra legacy (Fase 1) | 14/14 |
| Funciones migradas con tests verdes (Fase 3) | 14/14 |
| Callsites actualizados (Fase 4) | ~80/~80 |
| `verify_binary` final | `ok=true`, drift cero |
| Tamaño del binario post-compact | Reducido vs pre-migration (estimado ~3-4 MB liberados) |
| Commits work units | ~30-40 (1 por función migrada + 1 por bloque de callsites + cleanup) |
| Cobertura de tests para funciones migradas | 100% (cada función tiene su parity test + regression test) |

---

## Próximos pasos después de aprobar este plan

1. **Vos:** aprobás el plan (o lo ajustás).
2. **Yo:** ejecuto Fase 0 (setup) en este orden. Te aviso antes de cada operación que toca el binario para que compiles.
3. **Fase 1:** Te paso un script con los 14 tests de paridad para revisión antes de ejecutar `dysflow_test_vba`.
4. **Fases 2-6:** Avanzo por fase, con checkpoints entre cada fase. En modo Interactivo del proyecto SDD (default), te pregunto antes de cada fase si seguís o ajustás.

**Modo de ejecución:** Sugiero **Interactivo** (default del proyecto) por la magnitud del cambio. Si querés que avance sin pausa hasta encontrar un blocker, decime y pongo `auto`.

**Una sola pregunta:** ¿**Aprobás el plan tal cual**, **ajustás algo antes de arrancar**, o **empezamos por Fase 0 (setup) ya**?