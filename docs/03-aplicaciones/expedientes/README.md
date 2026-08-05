# 03 · Expedientes

## Propósito

Evidencia de descubrimiento de **Expedientes**, aplicación Access/VBA que crea y mantiene el expediente de contratación y sus relaciones. Este lote no propone diseño futuro: establece el inventario funcional, técnico y de migración que deberá conservarse.

## Estado

- **Fase:** descubrimiento completo, Batch 2.
- **Fecha de evidencia:** 2026-08-05.
- **Staging:** `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` (`origin/staging`), rama limpia salvo `Expedientes.accdb` modificado previamente.
- **Main de comparación:** `535c38a04da5f40aedfd0f1fe6a8464199807918` (`origin/main`), rama limpia salvo `Expedientes.accdb` modificado previamente.
- **Frontend:** `C:\00repos\codigo\00_EXPEDIENTES\staging\Expedientes.accdb`.
- **Backend autoritativo:** `C:\00repos\datos\Expedientes_datos.accdb`; no se ha tratado ningún backend repo-local como autoridad.
- **Dysflow:** `2.35.3`, envelope `dysflow.result/v1`, `projectId=expedientes-staging`; únicamente lecturas.

## Lote asociado

Lote 2 de `exploration.md` — Expedientes.

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES` (documentación leída antes de reverse engineering).
2. CodeGraph-VBA propio de `staging`.
3. Dysflow solo lectura sobre frontend y `C:\00repos\datos\Expedientes_datos.accdb`.
4. `staging` frente a `00_main`, comparados separadamente.
5. Engram como contexto histórico, no como fuente de comportamiento.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen de la documentación valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- Todo campo/registro del backend queda como `preservar hasta decisión`; no se declara obsoleto sin decisión explícita.

## Checklist

- [x] Backend autoritativo resuelto bajo `C:\00repos\datos`.
- [x] Inventario funcional, formularios, clases, módulos y tablas documentado.
- [x] Perfilado agregado ejecutado sin copiar filas personales.
- [x] Staging/main y suciedad preexistente registrados.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.

## Siguiente paso

Siguiente fase: revisar con negocio las filas `needs business decision` de la matriz, sin convertirlas todavía en diseño, propuesta o especificación.

## Descubrimiento de staging (2026-08-05, codegraph-vba)

Inspección del árbol staging (`0946b6a0a40acf4fb88eb62e435ec3f76f8a234d`) con codegraph-vba reveló que **el legacy VBA ya tiene forma hexagonal implícita**. El mapeo a la nueva plataforma hexagonal (D8, D66-D68) es uno a uno:

| Capa legacy VBA | Hexagonal | Equivalente nuevo (FastAPI) |
|---|---|---|
| Clases de dominio (`Expediente.cls`, `ExpedienteAGEDYS.cls`, `ExpedienteResponsable.cls`, `USUARIO.cls`, `Entorno.cls`) con `ColCampos`/`getPropiedad`/`SetPropiedad` | Dominio | Modelos SQLAlchemy 2.0 + Pydantic |
| `constructor.bas` (factory `getExpediente`, `getUsuario`, `getMostrarEstado`, `getEntorno`; 135 callers de `getdb()`) | Composition root | Composition root FastAPI con `Depends()` |
| `ExpedienteOperaciones.cls` con `Registrar` transaccional (`BeginTrans` + `CommitTrans` + `Rollback`) | Servicios / casos de uso | Casos de uso Python con `AsyncSession` |
| Helpers por dominio (`modExpedienteHelper.bas`, `modExpedienteEntidadesHelper.bas`, `modExpedienteHitosHelper.bas`...) | Adaptadores / handlers | Adaptadores FastAPI por agregado |
| `getdb()` + `TbConfiguracionBackends` (`BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `IDAplicacion`, `PasswordBackend`) | Puerto driven de persistencia | Puerto de BD con adapter de config |
| Forms (`Form_FormExpediente.cls`, `Form_FormTareas.cls`) | Adaptador driving UI | Routes HTMX + Jinja2 SSR |
| `CorreoOperaciones`, `ExpedienteAGEDYS`, integraciones externas | Adaptadores driven | Adaptadores de notificación, integraciones |
| Tests VBA (`Test_BackendCache`, `Test_BackendResolver`, `Test_ExpedienteCacheTransacciones`, un `Test_*` por cada `Helper_*`) | Evidencia de comportamiento | pytest + pytest-asyncio + httpx |

Mecanismo de sandbox/testing en staging:

- `m_TestingMode`, `m_TestOnlyBackendActivo`, `m_TestingBackendURL` en `Variables Globales.bas`.
- `TestOnly*` methods (`TestOnlySetBackendConfigOverride`, `TestOnlyOverrideActiveBackendURL`, `TestOnlyIsSandboxCacheStale`, etc.) que permiten a los tests cambiar backend sin tocar configuración real.
- Manifiestos: `tests.vba.json`, `tests.vba.responsable-71.json` (este último ligado al feature 71 responsable de aplicación), `Attribute-VB-Name-BaseName.Tests.ps1`.
- `Idempotencia de invalidación de caché`: `InvalidateBackendCache` se invoca en `LeeConfiguracionLocal` para garantizar que `getdb()` reabra contra el backend activo.

Implicaciones para el Lote 3 (Gestion_Riesgos) y siguientes:

- **D86**: la forma hexagonal del legacy se preserva como referencia de mapeo a la nueva plataforma. El blueprint hexagonal NO es una invención nueva; el legacy ya tenía esta forma.
- **D87**: los `Test_*` de VBA son evidencia de comportamiento que se preserva como referencia para los nuevos tests pytest. Ningún `Test_*` se descarta sin trazabilidad.

Esta capa no introduce diseño ni propuesta: deja el mapeo conceptual explícito para que las fases SDD posteriores (proposal, spec, design, tasks) arranquen con la forma destino ya caracterizada.
