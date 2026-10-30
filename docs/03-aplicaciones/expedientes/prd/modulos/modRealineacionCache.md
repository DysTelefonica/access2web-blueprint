# modRealineacionCache

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/modules/modRealineacionCache.bas` (209 líneas) · **PRD:** [README](../README.md)

Rutina de mantenimiento: realinea las copias derivadas (`TbExpedientes.CadenaPecal`/`Pecal` y `TbExpedientesConEntidades`) de los expedientes con sus tablas base, por el mismo punto de sincronización que usa toda escritura (issue #237). Las citas son `modRealineacionCache.bas:línea`.

| Función | Qué hace | Evidencia |
|---|---|---|
| `RealinearCacheExpedientes(db, desde, hasta)` | Recorre `TbExpedientes` en el rango de IDs (`0` = abierto), carga cada expediente con `constructor.getExpediente` (`:91`) y llama a `ExpedienteOperaciones.SincronizarDerivados` con ámbito `Todo` y la base inyectada (`:97-99`); un fallo se anota en `fallidos` y la pasada sigue. Devuelve JSON `{"ok","backend","desde","hasta","procesados","fallidos","error"}` | `:21-156` |
| `RealinearTodaLaCache(db)` | Lo anterior sobre todo el backend | `:157-209` |

La base de datos es **obligatoria e inyectada** por quien llama (`p_db is Nothing` es error, `:59-60`): no hay `getdb()` oculto ni marcador de pruebas. El rango existe para no saturar a los ejecutores remotos.

**Sin llamadores en `src` fuera de las pruebas** (`Test_Paso5DropHashCacheSuite.bas`, `Test_Tier2_ErrorPropagation.bas`): es una herramienta de operación, no un flujo de usuario.

## Comportamiento que la migración debe conservar

El concepto de copias derivadas, no el mecanismo: en web esas copias son vistas o columnas calculadas y no requieren realineación manual (propuesta pendiente de decisión).
