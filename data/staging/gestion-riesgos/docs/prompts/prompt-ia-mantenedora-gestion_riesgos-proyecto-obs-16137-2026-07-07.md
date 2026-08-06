Eres la IA mantenedora del proyecto VBA/Access `00_GESTION_RIESGOS_staging` (rama `b2-punto-15-anexos-helper`, PR #84). El proyecto se testea con dysflow MCP + TDD estricto, pero el bug que necesito resolver está en el código VBA del proyecto (no en dysflow).

## Contexto del round

He estado haciendo merge de varias rondas de trabajo en la feature B2 / Punto 15 / Anexos. El data layer y la UI están listos:
- 12/12 tests verdes en `tests.vba.friction-F15-F17-F18-F20.json`
- 2/2 tests sad-path verdes en `tests.vba.modAnexosListPresenter.json`
- 4/4 tests happy-path ROJOS con DAO error 3027 ("matriz es fija o se encuentra temporalmente bloqueada")
- Helper `modAnexosListPresenter.bas` con preload maps anti-N+1
- Botón "Ver Anexos" en `Form_FormRiesgosGestionEdicion`
- 3 fixes de issue #83 aplicados (Anexo.Tipo, EliminarAnexo cache, Riesgo.ColAnexosTotales re-raise)

## El bug que necesito que resuelvas

**4 tests del manifest `tests.vba.modAnexosListPresenter.json` fallan consistentemente con error DAO 3027 ("Esta matriz es fija o se encuentra temporalmente bloqueada"):**
- `Test_Presenter_Happy_5Anexos_5Filas`
- `Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo`
- `Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto`
- `Test_Presenter_CacheFirst_NoRompeContrato`

## Lo que YA funciona (NO tocar)

- Data layer completo: cierre de F15/F17/F18/F20 + 3 fixes de #83, todos verdes.
- UI: `modAnexosListPresenter.bas` con preload maps anti-N+1, `Form_FormAnexos.EstablecerLista` thin form, botón Ver Anexos en `Form_FormRiesgosGestionEdicion`.
- `tests.vba.friction-F15-F17-F18-F20.json`: 12/12 verde.
- `tests.vba.modAnexosListPresenter.json`: 2 sad-path atoms verdes (Nothing/Vacio).

## El bug en detalle

### Síntoma exacto

Mensaje de fallo del test runner:
```
Test_Presenter_Happy_5Anexos_5Filas — Test_Presenter_Happy_5Anexos_5Filas fallo: errMsg='' | Err.Description='Esta matriz es fija o se encuentra temporalmente bloqueada' | logs= |  |  |  |  |  |  |
```

- `errMsg=''` (vacío) → el parámetro `ByRef p_Error` de `Constructor.getAnexosTotalesDeEdicion` quedó vacío.
- `Err.Description='Esta matriz es fija...'` → la error global de VBA tiene este mensaje.
- `logs= |  | ...` → el array de logs del test está vacío, indicando que la línea `logs(0) = "..."` nunca se ejecutó.

### Repro exacta

**Path del manifest (serializado, una dysflow call a la vez):**

```bash
dysflow test_vba projectId=00-gestion-riesgos-staging testsPath=tests/tests.vba.friction-F15-F17-F18-F20.json
dysflow test_vba projectId=00-gestion-riesgos-staging testsPath=tests/tests.vba.modAnexosListPresenter.json
```

Después de ejecutar el manifest de fricción (12/12 verde), ejecutar el manifest del presenter (4 fallan).

### La pista clave: Mismo código funciona en friction, falla en presenter

| Aspecto | Friction (funciona) | Presenter (falla) |
|---|---|---|
| Módulo | `Test_modFormRiesgoDocumentosHelper` atom 11 | `Test_modAnexosListPresenter` atom 3 |
| Test | `Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica` | `Test_Presenter_Happy_5Anexos_5Filas` |
| SQL | `SELECT IDAnexo, ... FROM TbAnexos WHERE IDEdicion=910000 OR IDRiesgo IN (...)` | Idéntica |
| Función llamada | `Constructor.getAnexosTotalesDeEdicion(910000, errMsg)` | `Constructor.getAnexosTotalesDeEdicion(910000, errMsg)` |
| Setup | `EnsureTestingContext` + `SeedGrafoB2` | `EnsureTestingContext` + (sin SeedGrafoB2, asume datos de friction) |
| Resultado | 5/5 anexos, `col.Count = 5` | `col = Nothing`, `errMsg = ""`, `Err.Description = 'matriz es fija'` |
| Manifest | `tests.vba.friction-F15-F17-F18-F20.json` | `tests.vba.modAnexosListPresenter.json` |

**Diferencia clave:** el manifest del presenter carga `modAnexosListPresenter.bas` y `Test_modAnexosListPresenter.bas`, mientras que el manifest de fricción solo carga `Test_modFormRiesgoDocumentosHelper.bas`. El manifest de fricción NO carga `modAnexosListPresenter.bas`.

## Diagnóstico de investigación

He intentado lo siguiente sin éxito:
1. ❌ Schema-first audit (obs #16136): schema correcto, `Anexo.Tipo` no existe en SQL (lo hice propiedad computada).
2. ❌ 3 fixes de #83: aplicados, no resuelven "matriz es fija" — confirma que es problema DIFERENTE.
3. ❌ Eliminar `SeedGrafoB2` del presenter (confiar en friction): mismo error.
4. ❌ Correr presenter solo (sin friction antes): mismo error.
5. ❌ Correr presenter antes que friction: mismo error.
6. ❌ Test super-simplificado en atom 3 (BuildOk sin Constructor): PASA. Constructor.getAnexosTotalesDeEdicion es el problema.
7. ❌ Verificado que `EnsureTestingContext` no resuelve el problema.

## 5 hipótesis candidatas a validar

### Hipótesis 1: Recordset cacheado en `getdb().m_CachedDB`

`getdb()` retorna un `DAO.Database` cacheado (ver `Variables Globales.bas:1069-1140`). El manifest del presenter carga `modAnexosListPresenter.bas`, que en sus preloads (`PreloadEdicionDisplays`, `PreloadRiesgoCodigos`) llama `.Edicion` y `.riesgo` Property Gets en cada Anexo. Estos Property Gets abren recordsets via `Constructor.getEdicion`/`Constructor.getRiesgo`. Si la cache `m_CachedDB` no se cierra correctamente entre operaciones, podría tener un recordset bloqueado.

**Cómo verificar:** ejecutar `?getdb().m_CachedDB.Name` en el Immediate window durante el test. Si muestra error o un nombre diferente, el cache está corrupto.

### Hipótesis 2: Lock por transacción implícita al importar `modAnexosListPresenter`

Cuando `dysflow_import_modules` carga `modAnexosListPresenter.bas`, ¿deja una transacción implícita abierta? El módulo tiene un body grande con 218 líneas. Si la importación abre una transacción para validar el módulo, podría dejarla sin commit.

**Cómo verificar:** ejecutar `CurrentDb.Transactions(0).ActiveConnection` antes y después de importar el módulo.

### Hipótesis 3: Permisos del usuario sandbox

`getdb()` con sandbox abre la DB con `;PWD=`. Si el usuario del sandbox tiene permisos read-only sobre `TbAnexos`, `OpenRecordset` retornaría un recordset read-only (DAO error 3027).

**Cómo verificar:** después de `EnsureTestingContext`, ejecutar `CurrentDb.TableDefs("TbAnexos").Updatable` o `?db.TableDefs.Refresh`. Si retorna False, el sandbox user no tiene permisos de update.

### Hipótesis 4: Mala configuración de `m_TestingMode`

`m_TestingMode = True` debería hacer que `getdb()` use `m_BackendSandboxURL`. Pero si `m_ActiveBackendURL` no se actualiza correctamente después de `Test_EVE`, `getdb()` podría usar la DB de producción con permisos diferentes.

**Cómo verificar:** ejecutar `?getdb().m_ActiveBackendURL` y `?getdb().m_BackendSandboxURL` después de `EnsureTestingContext`. Si el primero no coincide con el segundo, hay un bug en `Test_EVE`.

### Hipótesis 5: DAO cache stale (cierre accidental de `m_CachedDB` durante test anterior)

Si un test anterior (friction u otro) cierra `m_CachedDB` (vía `ResetGlobals` o `ResetGetDbCache`), el siguiente `getdb()` intentaría reabrir la DB durante el `OpenRecordset`, pero podría haber un recordset huerfano que bloquea tablas.

**Cómo verificar:** ejecutar `?getdb().m_CachedDB Is Nothing` antes de cada test. Si está Nothing al final de un test, hay un bug de cleanup.

## Disciplina

- TDD estricto (RED → GREEN). No tocar tests que ya pasan.
- Conventional commits con SHA + test reference.
- Mantener cache-first: NO queries directas a DAO sin pasar por el helper cacheado.
- Mantener la convención de que `Constructor.getAnexosDe*Cached` es el ÚNICO path de lectura de anexos.
- Idempotencia: si no hay datos seedeados, los tests deben fallar limpio con `assert col.Count=5`, no con error de DAO.

## Acceptance output

1. Diagnóstico de la causa raíz de obs #16137: ¿cuál de las 5 hipótesis (u otra) es la correcta?
2. Fix mínimo: cambio quirúrgico que NO rompa los 12/12 del data layer NI los 2 sad-path del presenter.
3. 3 tests adicionales (RED first): tests que cubran el caso "matriz es fija" antes del fix, para que si vuelve a aparecer, tengamos cobertura.
4. 4 tests verdes (los happy-path del presenter) + 12/12 friction + 2/2 sad-path.
5. Actualización de `tests.vba.modAnexosListPresenter.json` con los 6 tests (incluyendo el fix).
6. Comentario en obs #16137 en engram con la causa raíz identificada.
7. Comentario en #83 cerrando el issue con SHA del fix (ya aplicado en `d45dbac`).
8. Manifest en `tests.vba.modAnexosListPresenter.json` añadido a `tests.vba.friction-F15-F17-F18-F20.json` (o manifest nuevo) si corresponde, para que el CI corra todos los tests.

## Archivos relevantes

- `src/modules/Test_modAnexosListPresenter.bas` — 4 tests failing
- `src/modules/modAnexosListPresenter.bas` — helper con preload maps
- `src/modules/Constructor.bas:1595-1657` — `getAnexosTotalesDeEdicion` (función que falla)
- `src/modules/Variables Globales.bas:1069-1140` — `getdb()` con cache `m_CachedDB`
- `src/modules/Test_Helper.bas:33-72` — `ResolveBackendSandbox`
- `src/modules/Test_Helper.bas:192-286` — `ForceLocalBackend` (incluye `Test_EVE`)
- `src/modules/Test_Fixtures.bas:57-102` — `GetTestDb` (cache de test)
- `docs/uat/estado-planificacion-NEW_2026-07-07.html` — planning actualizado
- `src/classes/Anexo.cls` — `Tipo` (computada), `EliminarAnexo` (invalida cache edicion padre)
- `src/classes/Riesgo.cls:1118-1139` — `ColAnexosTotales` (re-raise Err 1000)

## Quick start

```bash
# Repro
cd C:\00repos\codigo\00_GESTION_RIESGOS_staging
git log origin/b2-punto-15-anexos-helper..HEAD --oneline

dysflow test_vba projectId=00-gestion-riesgos-staging testsPath=tests/tests.vba.friction-F15-F17-F18-F20.json
dysflow test_vba projectId=00-gestion-riesgos-staging testsPath=tests/tests.vba.modAnexosListPresenter.json

# Investigar cache de conexión
dysflow dysflow_vba_execute projectId=00-gestion-riesgos-staging procedureName=GetTestDb moduleName=Test_Fixtures
dysflow dysflow_vba_execute projectId=00-gestion-riesgos-staging procedureName=ForceLocalBackend moduleName=Test_Helper

# Examinar fuente
cat src/modules/Test_modAnexosListPresenter.bas | head -300
cat src/modules/modAnexosListPresenter.bas
cat src/modules/Constructor.bas | grep -A 60 "getAnexosTotalesDeEdicion"
cat src/modules/Variables Globales.bas | grep -A 70 "Public Function getdb"
```

## Output esperado del mantenedor

1. **Diagnóstico de la causa raíz de obs #16137**: ¿cuál de las 5 hipótesis (u otra) es la correcta?
2. **Fix mínimo**: cambio quirúrgico que NO rompa los 12/12 del data layer NI los 2 sad-path del presenter.
3. **3 tests adicionales** (RED first): tests que cubran el caso "matriz es fija" antes del fix, para que si vuelve a aparecer, tengamos cobertura.
4. **4 tests verdes** (los happy-path del presenter) + 12/12 friction + 2/2 sad-path.
5. **Actualización de `tests.vba.modAnexosListPresenter.json`** con los 6 tests (incluyendo el fix).
6. **Comentario en obs #16137 en engram** con la causa raíz identificada.
7. **Comentario en #83 cerrando el issue** con SHA del fix (ya aplicado en `d45dbac`).
