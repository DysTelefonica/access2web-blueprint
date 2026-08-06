# Deuda crítica de pruebas Access/VBA TDD v2.4.2

> Proyecto: CONDOR
> Fecha de registro: 2026-06-15
> Alcance: documentación de deuda transversal de pruebas. No se modifica VBA ni se ejecutan operaciones Access en esta tarea.

## 1. Resumen ejecutivo

La auditoría de alineación con `access-vba-tdd` v2.4.2 detecta que las suites existentes de CONDOR no deben aceptarse de forma general como nueva evidencia `Verified-runtime` hasta completar una migración del harness, del aislamiento de datos y de los manifests.

La evidencia histórica puede conservarse como contexto acotado por capacidad, pero no debe promocionarse ni reutilizarse como cierre de regresión para nuevas reglas sin pasar por el contrato v2.4.2: ciclo de vida canónico, fixture propio, revisión schema-first, `DAO.Database` explícito, cardinalidad de mutaciones y separación clara entre pruebas atómicas y smoke.

## 2. Impacto sobre la evidencia

| Evidencia | Estado bajo v2.4.2 | Uso permitido |
|---|---|---|
| Suites existentes en general | No aceptables como nueva evidencia `Verified-runtime` sin auditoría y migración | Contexto histórico o evidencia parcial explícitamente acotada |
| Runner probe | Diagnóstico técnico de invocación | Solo confirma que el runner puede llamar procedimientos; no prueba comportamiento funcional |
| E2E de plantillas documentales | Candidato de evidencia read-only | Puede sostener contrato de plantilla/mapeo si se ejecuta en verde; no cubre mutaciones ni persistencia de solicitudes |
| Pruebas de mutación PCSUB/CDCA | Evidencia incompleta mientras falte cardinalidad o harness canónico | Deben migrarse antes de cerrar regresión runtime nueva |

## 3. Bloqueadores críticos

1. **Harness legacy**: persisten `SuiteSetup` y `SuiteTeardown`; no hay evidencia suficiente del ciclo canónico `BeginTestSession` / `EndTestSession` / `ResetTestSession` con `Test_EVE(True/False)`.
2. **Mutación de configuración productiva**: SetTestBackend en src/modules/TestHelper.bas mutaba TbConfiguracionBackends; este patrón queda prohibido por v2.4.2 — **RESUELTO 2026-06-15: Public Sub SetTestBackend() eliminado de src/modules/TestHelper.bas; solo se conserva ForceLocalBackend que no muta la configuración**. Adicionalmente, SandboxLogger.bas y la tabla tbSandboxLog asociada fueron eliminados del slice 2026-06-15: las 9 llamadas SandboxLogger.Log* de SandboxGestor.cls se redirigieron a TraceEngine existente; las 2 llamadas de SandboxValidator.cls se reemplazaron por Debug.Print inline; los helpers privados EnsureSandboxMetadata y EnsureTbSandboxLogInSandbox se eliminaron; VerifyPostBuildCompleteness se simplificó a exigir solo tablas configuradas; ValidateMetadata se eliminó. TbConfiguracionBackends permanece activa como configuración de selección de backend y no debe ser mutada por tests.
3. **Precheck de sandbox incompleto**: el guard general debe demostrar rechazo de rutas UNC, rechazo de fingerprints productivos y exigencia de fingerprint CONDOR antes de tocar datos — **RESUELTO para `Test_PCSUB_Strict` en commit `f27f693` mediante bloqueo de UNC, `\\datoste\\` y rutas sin `condor_datos.accdb`; pendiente revisar/migrar el resto de suites**.
4. **CDCA pasa `Nothing` como `DAO.Database`**: varias pruebas fuerzan el camino interno de `getdb()` en lugar de probar inyección explícita de base de datos.
5. **Cardinalidad de mutaciones incompleta**: PCSUB y CDCA no demuestran de forma homogénea `countBefore` / `countAfter` para cada `INSERT`, `UPDATE` o `DELETE` bajo prueba.
6. **Deuda de manifests**: `tests/tests.pcsub.json` ya fue migrado a 29 pruebas atómicas BR-001..BR-011 en commit `f27f693`; queda pendiente corregir el resto de manifests legacy, incluyendo CDCA.

## 4. Orden obligatorio de migración

1. Migrar el harness a `BeginTestSession` / `EndTestSession` / `ResetTestSession` e integrar `Test_EVE(True)` en apertura y `Test_EVE(False)` en cierre.
2. Endurecer `ForceLocalBackend`/guards equivalentes para bloquear UNC, fingerprints productivos y rutas que no contengan el fingerprint esperado de CONDOR. **PCSUB strict resuelto en `Test_PCSUB_Strict.SetupSandbox`; pendiente generalizar al resto de suites**.
3. ~~Retirar~~ Retirado `SetTestBackend` y cualquier escritura de tests contra `TbConfiguracionBackends` — **COMPLETADO 2026-06-15**.
4. Migrar CDCA a `DAO.Database` explícito en todas las llamadas bajo prueba; eliminar el uso de `Nothing` como base de datos.
5. Añadir `countBefore` / `countAfter` a todas las pruebas de mutación y fallar con mensaje claro si la cardinalidad esperada no se cumple.
6. Separar manifests: pruebas atómicas en manifests atómicos, agregadores `*_RunAll` solo en manifests smoke.
7. Solo después de los pasos anteriores, promover evidencia a `Verified-runtime`.

## 5. Criterio de salida

Una capacidad podrá declarar nueva evidencia `Verified-runtime` cuando su manifest ejecute funciones públicas globales únicas, devuelva JSON canónico, use fixture propio en sandbox, documente schema-first para los datos tocados, inyecte `DAO.Database` explícito, verifique cardinalidad de mutaciones y no muten configuración productiva (nótese que `SetTestBackend` fue eliminado el 2026-06-15) ni dependan de datos existentes por azar ni probes diagnósticos.

## 6. Referencias de capacidad

- [CAP-001 — PCSUB](../capabilities/CAP-001-pcsub.md): BR-001..BR-011 de servicio/repositorio ya tienen manifest atómico real, cardinalidad en mutaciones y evidencia runtime (`f27f693`); queda deuda UI/formulario y decisiones de producto sobre completitud/RAC delegado.
- [CAP-002 — Documentos Word y mapeos](../capabilities/CAP-002-documentos-word-mapeo.md): el E2E de plantilla es candidato read-only, no evidencia de mutación o persistencia.
- [CAP-003 — CDCA](../capabilities/CAP-003-cdca.md): requiere migración prioritaria por uso de `Nothing` como `DAO.Database`, harness legacy y deuda de cardinalidad.
