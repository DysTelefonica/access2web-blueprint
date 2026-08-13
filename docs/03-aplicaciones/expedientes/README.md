# 03 · Expedientes

## Propósito

Evidencia y planificación de la migración de **Expedientes**, aplicación Access/VBA que crea y mantiene el expediente de contratación y sus relaciones. El descubrimiento alimenta un cambio SDD completo orientado a paridad de capacidades, no a réplica de formularios ni de herramientas legacy.

## Estado

- **Fase:** planificación SDD completa: proposal, diez specs, design, tasks y backlog ejecutable.
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
7. [Ledger de capacidades](capability-ledger.md)
8. [Backlog de issues de migración](issue-backlog.md)

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

Ejecutar las unidades documentales D01–D54 por orden de dependencia y resolver cada decisión bloqueante antes del slice que la consume. Las 110 unidades ya tienen issue enlazada en el backlog. Proposal, specs, design y tasks ya existen bajo `openspec/changes/expedientes-web-migration/`.

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

## Inventario real Dysflow (2026-08-05, segunda pasada)

**Volúmenes principales del backend autoritativo** (`C:\00repos\datos\Expedientes_datos.accdb`):

- **`TbExpedientes`**: **453 filas** (expedientes en producción).
- **`TbExpedientesAnexos`**: **712 filas** (anexos, más que expedientes por anexo múltiple).
- **49 tablas totales** (segunda app más grande después de Gestion_Riesgos con 71). Fueron observadas en runtime el 2026-08-05; esta fase de planificación no ha repetido esa validación, por lo que D40–D44 deben refrescar schema, ownership, volumen y uso antes del DDL final.

**Hallazgos críticos del esquema de `TbExpedientes`** (55 columnas reales, esquema de procurement regulatorio):

- **Jerarquía AM/lote**: `IDExpedientePadre` (Long) — jerarquía recursiva. La nueva plataforma debe **formalizar con CTE recursivo** en PostgreSQL.
- **Importes como `Currency` (type 7)**: `ImporteLicitacion`, `ImporteContratacion` (tamaño 8 bytes). **Crítico para auditoría regulatoria**.
- **Códigos múltiples**: `CodProyecto`, `CodExp`, `CodExpLargo`, `CodS4H` (4 códigos por expediente). **Migración**: cada código se preserva como columna o como tabla de normalización.
- **Fechas del workflow regulatorio** (11 fechas, algunas en MAYÚSCULAS ⚠️): `FechaInicioContrato`, `FechaFinContrato`, `FechaFinGarantia`, `FECHAPREOFERTA` (⚠️), `FECHAINICIOLICITACION` (⚠️), `FECHAOFERTA` (⚠️), `FECHAADJUDICACION` (⚠️), `FECHAFIRMACONTRATO` (⚠️), `FECHACERTIFICACION` (⚠️), `FECHAPERDIDA` (⚠️), `FECHADESESTIMADA` (⚠️).
- **Flags booleanos como Text(2)**: 14 columnas con valores 'Sí'/'No' (⚠️ cross-cutting con HPS, NoConformidades, Condor). Inconsistencia a normalizar en PostgreSQL.
- **Hashes E2E** ⚠️⚠️: `HashActual` (Text 64), `HashUltimaExportacion` (Text 64) — **patrón de sincronización bidireccional** entre Expedientes y Lanzadera. La nueva plataforma debe **preservar este patrón** (sincronización E2E con verificación de hash).
- **FKs conceptuales sin constraint**: `IdGradoClasificacion`, `IDOrganoContratacion`, `IDOficinaPrograma`, `IDEjercito`, `IDResponsableCalidad`, `IDResponsableSeguridad`, `IDEstado` (a catálogos internos). Decisión D94 sigue aplicando.
- **Audit con `IDUsuario` como Text(255)**: `IDUsuarioCreacion`, `IDUsuarioUltimoCambio` (⚠️ NO son FK numéricas a `tbUsuarios` — son **strings** que probablemente contienen el email o el `UsuarioRed` del usuario). Migración: agregar FK numérica real.

**Sub-tablas de Expedientes** (14+ tablas que comparten `IDExpediente` con la cabecera):

- `TbExpedientesAnexos` (712 filas), `TbExpedientesAnualidades` (174), `TbExpedientesCadenaContratacion`, `TbExpedientesCodigoCompras`, `TbExpedientesComerciales`, `TbExpedientesConEntidades`, `TbExpedientesCPVs`, `TbExpedientesE2E`, `TbExpedientesHitos` (46), `TbExpedientesJefaturas`, `TbExpedientesJuridicas` (417), `TbExpedientesLugaresEjecucion` (194), `TbExpedientesModificados` (37), `TbExpedientesPECAL`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbExpedientes_antes` (legacy, presumido).

**Inconsistencias cross-cutting detectadas en flags Sí/No** (coherente con HPS, NoConformidades, Condor): 14 columnas Text(2) en `TbExpedientes` + columnas análogas en las sub-tablas. **D102 propuesta**: estandarizar a `BOOLEAN` en PostgreSQL con regla de migración explícita (similar a D92 cross-cutting).

**Sentinels detectados** (de `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `ListaPrevia`, `TbAusExpPostAGEDO`, `TbAuxEstadosMartina`, `TbAuxNemotecnico`, `TbConfMostrarEstado`): patrón legacy de copia antes de cambios masivos + tablas auxiliares.

**`.dysflow/project.json` creado en esta pasada** (con `setup_project` autorizado) en `00_EXPEDIENTES/00_main/.dysflow/`. `projectId: 00-expedientes-staging`, `frontendFile: Expedientes.accdb`, `allowWrites: true`.
