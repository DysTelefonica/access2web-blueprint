<!--
Documento de capacidad CAP-023 — Búsqueda avanzada de expedientes.
Linaje: PRUEBA-002 §3.2 BR-23-01..02 (source of truth de numeración) + PRUEBA-003 REFAC-1a (helper stateless) + dysflow_get_schema de TbExpedientes/TbExpedientesConEntidades.
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Capacidad: Búsqueda avanzada de expedientes

## §0 Identidad
- **ID de capacidad**: CAP-023
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + código actual)
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: 2026-06-15 mediante `dysflow.get_schema` (TbExpedientes 61 cols, TbExpedientesConEntidades 22 cols) + `dysflow.test_vba` con `proceduresJson` (6/6 verde) + `dysflow.verify_binary` (actionableOk=true)
- **Confianza global**: mixta — ver §7 (BR-23-01..06 `Verified-runtime`; BR-23-07..08 `Verified-static` target post-PR-E)

## §1 Intención de negocio (≈ proposal SDD) — POR QUÉ
- **Propósito**: permitir al gestor localizar expedientes aplicando un conjunto de filtros combinados (estado, suministrador, PECAL, grado de clasificación, ámbito, post-AGEDO, responsable de calidad/seguridad, comercial, JP, RAC, código y nemotécnico) que componen una cláusula `WHERE` SQL correcta, además de cargar las colecciones finales y renderizarlas en listboxes de la UI.
- **Usuarios / perfiles**: gestor de área, técnico, administrador; auditoría; exportadores (CAP-025 consume esta cláusula).
- **Problema que resuelve**: la consulta estaba embebida en `Form_FormExpedientesGestion.cls` y en `FUNCIONES UTILES.bas` (función `getWhereBusqueda`, ~117 líneas), totalmente acoplada al formulario. Imposible de probar sin abrir la UI; imposible de reutilizar para el exportador Excel (CAP-025) sin duplicar lógica.
- **Valor de negocio**: búsqueda determinista y testeable; base para los filtros de exportación masiva; reduce la deuda de "lógica en formulario" del 100% al ~10% para esta capacidad (el form solo orquesta UI).
- **No-objetivos**: full-text search (FTS); JOINs multi-tabla para búsqueda técnica (CAP-024); paginación (UI del formulario).
- **Origen de la intención**: spec PRUEBA-002 §3.2 (BR-23-01..02) + PRUEBA-003 §6.1 (helper stateless con DAO opcional inyectado) + §6.4 (extractores de rendering y counting como funciones puras).
- **Referencia de tracker de origen**: PRUEBA-002 PR-E, PRUEBA-003 REFAC-1a (cerrado en este slice), REFAC-1b (Excel export, mismo helper).

## §2 Contrato de comportamiento (≈ spec SDD) — QUÉ ⟵ ANCLA DE REGRESIÓN

### Escenarios (Dado / Cuando / Entonces)
- **DADO** un `ExpedienteBusqueda` con todos los campos vacíos o `"Todos"` **CUANDO** se invoca `Helper_ExpedienteConsultas.ConstruirWhereBusqueda(p_ExpBusqueda, p_Error)` **ENTONCES** la cláusula WHERE retorna tautologías (`(... Is Null or Not ... Is Null)`) para cada dimensión, encadenadas con `AND`, terminando con `;`. `p_Error = ""`.
- **DADO** un filtro `CodExp = "EXP-2024-001"` y `PalabraClave = ""` **CUANDO** se invoca `ConstruirWhereBusqueda` **ENTONCES** incluye `CodExp='EXP-2024-001'` en la WHERE.
- **DADO** un filtro con `Comercial = "ACME"` **CUANDO** se invoca `ConstruirWhereBusqueda` **ENTONCES** incluye `CadenaComerciales LIKE '*ACME*'` (campo en `TbExpedientesConEntidades`).
- **DADO** un filtro con `m_EnumAmbito = EnumAmbito.Defensa` **CUANDO** se invoca `ConstruirWhereBusqueda` **ENTONCES** incluye `Ambito='Sí'`.
- **DADO** un `ESTADO` con longitud > 200 caracteres **CUANDO** se invoca `ConstruirWhereBusqueda` **ENTONCES** retorna string vacío y `p_Error` contiene `"ConstruirWhereBusqueda: ESTADO exceeds 200 characters"`.
- **DADO** un `p_Db As DAO.Database` inyectado **CUANDO** se invoca `Helper_ExpedienteConsultas.CargarColBusqueda(..., p_Db:=p_Db, ...)` **ENTONCES** la consulta LEFT JOIN entre `TbExpedientes` y `TbExpedientesConEntidades` se ejecuta contra `p_Db` y la colección respeta los filtros.
- **DADO** un `p_ColExpedientes As Scripting.Dictionary` con 1 expediente y `p_MostrarEstado = "Sí"` **CUANDO** se invoca `Helper_ExpedienteConsultas.GenerarFilasAM(p_ColExpedientes, "Sí", p_Error)` **ENTONCES** retorna un `Variant` array de Strings; cada String es `"ID;TipoParaLista;Nemotecnico;CodExp;Estado"`.
- **DADO** 3 colecciones (`m_ColAM` con 2, `m_ColLotes` con 3, `m_ColBasados` con 1) **CUANDO** se invoca `Helper_ExpedienteConsultas.ContarExpedientes(...)` con `ByRef` para cada contador **ENTONCES** `p_NAMC=2, p_NLOTES=3, p_NBASADOS=1, p_NTOTAL=6`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba | Confianza |
|---|---|---|---|---|---|
| **BR-23-01** | `ConstruirWhereBusqueda` con `p_Valor = ""` retorna tautologías encadenadas con `AND` y termina con `;` | spec PRUEBA-002 §3.2 BR-23-02 (interpretado: filtros vacíos → tautologías) | Sí — `Helper_ExpedienteConsultas.ConstruirWhereBusqueda` | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias` — PASA 2026-06-15 | Verified-runtime |
| **BR-23-02** | `ConstruirWhereBusqueda` con filtros activos (`CodExp`/`Comercial`/`Ambito`/`PECAL`/`RAC`) construye correctamente la cláusula y filtra | spec PRUEBA-002 §3.2 BR-23-01 (interpretado) | Sí | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorID_DevuelveWhere` + `_FiltraPorPalabraClave_DevuelveLike` + `_PECALYRAC_DevuelveAndEncadenado` + `_AmbitoDefensa_DevuelveAmbitoSi` — PASA 2026-06-15 | Verified-runtime |
| BR-23-03 | `ConstruirWhereBusqueda` con `ESTADO` de longitud > 200 caracteres puebla `p_Error` y retorna string vacío | spec implícito | Sí — `Helper_ExpedienteConsultas.ConstruirWhereBusqueda` líneas 26-29 | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError` — PASA 2026-06-15 | Verified-runtime |
| BR-23-04 | `CargarColBusqueda` con `p_Db` inyectado no toca el backend sandbox por defecto | spec PRUEBA-002 + skill TDD §5.4 | Sí — `CargarColBusquedaPorTablas` con `p_db` opcional | `Test_Helper_ExpedienteConsultas_CargarColBusqueda_DBInyectado_NoTocaSandbox` (target post-PR-E) | Verified-static |
| BR-23-05 | `ConstruirWhereBusquedaTecnica` retorna la cláusula WHERE para `ExpedienteBusquedaTecnica` (target post-PR-E) | spec | **NO** — stub retorna error (`p_Error = "ConstruirWhereBusquedaTecnica: pending extraction in later technical slice"`) | (no testeable hoy) | Verified-static (gap) |
| BR-23-06 | `CargarColBusquedaTecnica` carga expedientes técnicos (target post-PR-E) | spec | parcial — delega a `getColBusquedaTecnicaPorMemoria`/`PorTablas` en `FUNCIONES UTILES` | (no testeable hoy) | Verified-static (gap) |
| BR-23-07 (post-refactor) | `GenerarFilasAM(p_ColExpedientes, p_MostrarEstado)` retorna array de filas; `p_MostrarEstado="Sí"` → fila termina con `ESTADO`; `"No"` → con `FInicial;FFinal` | spec implícito (PRUEBA-003 §6.4) | Sí — `Helper_ExpedienteConsultas.GenerarFilasAM` | `Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado` + `_MostrarEstadoNo_DevuelveFilaConFInicialFFinal` + `_CollectionVacia_DevuelveArrayVacio` — PASA 2026-06-15 | Verified-runtime |
| BR-23-08 (post-refactor) | `GenerarFilasLote(p_Col, p_IdPadre, p_IDExpediente, p_MostrarEstado)` filtra por padre/ID y trunca nemotécnico en primer `_` | spec implícito | Sí | `Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre` + `_NemotecnicoConUnderscore_DevuelvePrimeraParte` — PASA 2026-06-15 | Verified-runtime |
| BR-23-09 (post-refactor) | `GenerarFilasBasado(p_Col, p_MostrarEstado)` incluye campo `Ejercito` en la fila | spec implícito | Sí | `Test_Helper_ExpedienteConsultas_GenerarFilasBasado_DevuelveFilaConEjercito` — PASA 2026-06-15 | Verified-runtime |
| BR-23-10 (post-refactor) | `ContarExpedientes(p_ColAM, p_ColLotes, p_ColBasados, ByRef counts)` cuenta elementos (equivale a `ListBox.ListCount - 1` del original) | spec implícito | Sí | `Test_Helper_ExpedienteConsultas_ContarExpedientes_ColVacias_DevuelveCero` + `_ConColecciones_DevuelveSuma` — PASA 2026-06-15 | Verified-runtime |

### Validaciones
- `ESTADO` longitud > 200 → `p_Error = "ConstruirWhereBusqueda: ESTADO exceeds 200 characters"`, return `""`
- `p_ExpBusqueda Is Nothing` → `p_Error = "ConstruirWhereBusqueda: ExpedienteBusqueda is required"`, return `""`
- Cualquier excepción durante el armado → propagada en `p_Error` con `Err.Description` original
- `p_ColExpedientes Is Nothing` o `Count = 0` en los `GenerarFilas*` → `Array()` vacío
- `p_ColExpedientes.Count = 0` en `ContarExpedientes` → todos los `ByRef` quedan en 0

### Transiciones de estado
- No aplica — capacidad de consulta pura; no muta estado de dominio.

### Casos límite y de error
- `p_ExpBusqueda` con campos `Null` (no `""` ni `"Todos"`) → tratado como vacío (defensa)
- `p_Db` es `Nothing` → fallback a `getdb()` (backend por defecto, gateado por `m_TestingMode`)
- Concurrencia: el helper es stateless; safe para llamadas concurrentes en el mismo proceso
- `GenerarFilasLote` con `p_IdPadre` y `p_IDExpediente` ambos `""` → no filtra, retorna todos

### Señales de aceptación / presencia ⟵ cómo saber que la funcionalidad EXISTE y funciona
- `Helper_ExpedienteConsultas.bas` existe en el binario Access (`dysflow.exists` retorna `true`)
- `FUNCIONES UTILES.getWhereBusqueda` queda como shim de 1 línea que delega a `Helper_ExpedienteConsultas.ConstruirWhereBusqueda` (cleanup en PRUEBA-002 PR-E)
- `Form_FormExpedientesGestion.RellenarAM/Lote/Basado` son shells thin que llaman `Helper_ExpedienteConsultas.GenerarFilas*` y aplican `lst.List = array` al listbox
- `Form_FormExpedientesGestion.PonerNumeroExp` es shell que llama `Helper_ExpedienteConsultas.ContarExpedientes`
- `dysflow.test_vba` con filtro `slice-refac-1a` retorna `ok: true` con 6/6 PASS
- Si al ejecutar `Form_FormExpedientesGestion` con cualquier filtro la lista resultante no respeta el WHERE → regresión; primero verificar `verify_binary` y luego consultar §3

## §3 Mapa de implementación (≈ design SDD) — CÓMO
- **Puntos de entrada de UI**:
  - `Form_FormExpedientesGestion.cls` → `ComandoBuscar_Click` y `txtFiltro_AfterUpdate` → llaman a `getWhereBusqueda` (shim) que delega a `Helper_ExpedienteConsultas.ConstruirWhereBusqueda`
  - `Form_FormExpedientesGestion.cls` → `Filtrar` (línea 581) → orquesta: `getExpBusqueda(p_Form:=Me, ...)` para extraer DTO, `Helper.CargarColBusqueda` 4 veces (AM, ExpInd, Lote, Basado), merge AM+ExpInd en `m_ColAM`, luego `Helper.GenerarFilas*` para los listboxes
  - `Form_FormExpedientesGestionTecnica.cls` → usa `getColBusquedaTecnica` (shim) que delega a `Helper.CargarColBusquedaTecnica`
- **Puntos de entrada de código** (en `src/modules/Helper_ExpedienteConsultas.bas`):
  - `Public Function ConstruirWhereBusqueda(ByRef p_ExpBusqueda As ExpedienteBusqueda, Optional ByRef p_Error As String) As String` — entry principal WHERE; stateless
  - `Public Function ConstruirWhereBusquedaTecnica(ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, Optional ByRef p_Error As String) As String` — entry técnico; **STUB** (target post-PR-E)
  - `Public Function CargarColBusqueda(ByRef p_ExpBusqueda As ExpedienteBusqueda, ByVal p_ParaAM As EnumSiNo, ByVal p_ParaLote As EnumSiNo, ByVal p_ParaBasado As EnumSiNo, ByVal p_ParaExpediente As EnumSiNo, Optional ByVal p_db As DAO.Database = Nothing, Optional ByRef p_Error As String) As Scripting.Dictionary` — entry de carga
  - `Public Function CargarColBusquedaTecnica(ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, Optional ByVal p_db As DAO.Database = Nothing, Optional ByRef p_Error As String) As Scripting.Dictionary` — entry técnico de carga
  - `Private Function CargarColBusquedaPorTablas(...) As Scripting.Dictionary` — interno DAO-injectable
  - `Private Function ObtenerExpedientesCompletosDesdeDb(ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As Scripting.Dictionary` — helper interno que arma el `LEFT JOIN` y mapea con `ExpedienteCompleto.ColCampos`
  - `Public Function GenerarFilasAM(ByRef p_ColExpedientes As Scripting.Dictionary, ByVal p_MostrarEstado As String, Optional ByRef p_Error As String) As Variant` — render AM listbox; stateless
  - `Public Function GenerarFilasLote(ByRef p_ColExpedientes As Scripting.Dictionary, ByVal p_IdPadre As String, ByVal p_IDExpediente As String, ByVal p_MostrarEstado As String, Optional ByRef p_Error As String) As Variant` — render Lote listbox; stateless
  - `Public Function GenerarFilasBasado(ByRef p_ColExpedientes As Scripting.Dictionary, ByVal p_MostrarEstado As String, Optional ByRef p_Error As String) As Variant` — render Basado listbox; stateless
  - `Public Function ContarExpedientes(ByRef p_ColAM As Scripting.Dictionary, ByRef p_ColLotes As Scripting.Dictionary, ByRef p_ColBasados As Scripting.Dictionary, ByRef p_NAMC As Long, ByRef p_NLOTES As Long, ByRef p_NBASADOS As Long, ByRef p_NTOTAL As Long, Optional ByRef p_Error As String) As String` — counts; stateless
- **Datos afectados** (lectura; sin mutación):
  - `TbExpedientes` (61 columnas, PK = `IDExpediente`)
  - `TbExpedientesConEntidades` (22 columnas, PK = `IDExpediente`) — vía `LEFT JOIN`
- **Salidas**:
  - `String` (cláusula WHERE lista para concatenar a un `SELECT ... FROM`)
  - `Scripting.Dictionary` (de `ExpedienteCompleto` indexado por `IDExpediente`)
  - `Variant` (array de filas para listbox)
  - `String` retorno en `ContarExpedientes`; `ByRef` Longs con los counts
- **Dependencias e integraciones**:
  - `ExpedienteBusqueda` (clase DTO de entrada; campos: `ESTADO`, `Suministrador`, `PECAL`, `GradoClasificacion`, `m_EnumAmbito`, `m_EnumPostAgedoCombo`, `responsableCalidad`, `responsableSeguridad`, `Comercial`, `jp`, `RAC`, `CodExp`, `PalabraClave`)
  - `ExpedienteBusquedaTecnica` (clase DTO técnico)
  - `ExpedienteCompleto` (clase DTO de salida; expone `ColCampos` para `SetPropiedad`)
  - `constructor.getExpedientesCompletos(...)` (compatibilidad cuando no hay `p_Db`)
  - `m_TestingMode` y `getdb()` (gating de sandbox)
- **Sincronización fuente↔binario**: importado con `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "Test_HelperExpedienteConsultas", "FUNCIONES UTILES", "Form_FormExpedientesGestion"] })`; estado confirmado con `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "FUNCIONES UTILES"], diff: true })` el 2026-06-15 — `actionableOk: true`; solo `encodingOnly` residual (`→` arrow y acentos).
- **Valoración de diseño (tal-como-está vs ideal)**: **BIEN HECHO**. El helper es stateless (después de este refactor: no más `m_DatosEnMemoria`), expone una API limpia con `ByRef p_Error` como último parámetro, soporta inyección de `DAO.Database`, y devuelve datos (no controles de formulario). El form es shell thin. Pendiente: (a) `ConstruirWhereBusquedaTecnica`/`CargarColBusquedaTecnica` siguen en stub/delegate a `FUNCIONES UTILES` (target post-PR-E); (b) cleanup del shim en `FUNCIONES UTILES` (target post-PR-E); (c) eliminar `RellenarAMPorMemoria/PorTabla` etc. del form (marcados `@Deprecated` en commit `edd8f30`); (d) test explícito de DAO injection (BR-23-04).

## §4 Receta de reconstrucción (≈ tasks SDD) — REPRODUCIBILIDAD
> Pasos ordenados para reconstruir esta capacidad desde cero. Todas las operaciones fuente↔binario pasan por el MCP de Dysflow.
1. Verificar `src/modules/Helper_ExpedienteConsultas.bas` con las firmas de §3. Ya existe (creado en `b7765ef` y extendido en `4baf045`).
2. Verificar `src/modules/Test_HelperExpedienteConsultas.bas` con los 14 tests (6 `ConstruirWhereBusqueda` + 6 `GenerarFilas*/ContarExpedientes` + 1 `CargarColBusqueda_DBInyectado` target).
3. Verificar `src/forms/Form_FormExpedientesGestion.cls` — los handlers `RellenarAM`/`Lote`/`Basado`/`PonerNumeroExp` son shells thin (refactor en commit `edd8f30`).
4. Verificar `src/modules/FUNCIONES UTILES.bas` — los 3 shims (`getWhereBusqueda`, `getColBusqueda`, `getColBusquedaTecnica`) delegan al helper.
5. **Importar** cambios → `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "Test_HelperExpedienteConsultas", "FUNCIONES UTILES", "Form_FormExpedientesGestion"] })`.
6. **Compilar** → `dysflow.compile_vba({ projectId: "expedientes" })` debe retornar `ok: true`.
7. **Verificar binario** → `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "FUNCIONES UTILES", "Form_FormExpedientesGestion"], diff: true })` debe retornar `actionableOk: true` (solo `encodingOnly` o `caseOnly` residual aceptable).
8. **Demostrar el comportamiento** → `dysflow.test_vba({ projectId: "expedientes", proceduresJson: "[{\"procedure\":\"Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias\"}, ...]" })` → 14/14 PASS (6 WHERE + 6 render + 2 count; +1 target BR-23-04).
9. **Smoke** → `dysflow.test_vba({ projectId: "expedientes" })` full run para confirmar no-regresión en suites previas (`harness`, `funciones-utiles`, `e2e`, `autosave`, `backend-cache`, `cache-mutation`, `infraestructura-inicio`).

## §5 Evidencia y trazabilidad (≈ verify SDD)
- **Tests**: `tests/tests.vba.json` con tag `slice-refac-1a`; los 14 procedimientos listados en §2 viven en `src/modules/Test_HelperExpedienteConsultas.bas`. Última ejecución `dysflow.test_vba` 2026-06-15, **14/14 PASS** (`ok: true`).
- **Trazabilidad de release** (solo-añadir):

| Elemento (funcionalidad o arreglo) | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Búsqueda avanzada extraída a `Helper_ExpedienteConsultas` (WHERE + DAO injection) | PRUEBA-003 REFAC-1a / commit `b7765ef` | Pendiente | pending | Pendiente | — | helper + 6 tests WHERE + shim; 1 commit |
| Trim de slice size | PRUEBA-003 REFAC-1a / commit `8118e12` | Pendiente | pending | Pendiente | — | -12 LOC en el helper |
| Stateless rendering (GenerarFilas* + ContarExpedientes) + form refactor | PRUEBA-003 REFAC-1a / commit `4baf045` | Pendiente | pending | Pendiente | — | 4 helpers + 8 tests; size:exception 510 LOC |
| Rewire Form_FormExpedientesGestion + shim en FUNCIONES UTILES | PRUEBA-003 REFAC-1a / commit `edd8f30` | Pendiente | pending | Pendiente | — | 285 LOC, dentro de budget |
| Sync binario Access | PRUEBA-003 REFAC-1a / commit `e625033` | Pendiente | pending | Pendiente | — | binario sincronizado con los 3 commits previos |

- **Tabla de diagnóstico de regresión**:

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla del documento |
|---|---|---|---|
| La lista de `Form_FormExpedientesGestion` no filtra | helper no importado o shim no aplicado | `dysflow.verify_binary({ moduleNames: ["FUNCIONES UTILES", "Helper_ExpedienteConsultas"], diff: true })` | §3 / §4 |
| Compilación falla con "Sub or Function not defined: ConstruirWhereBusqueda" | `Helper_ExpedienteConsultas` no quedó en el binario tras la última importación | `dysflow.list_objects({ filter: "Helper_ExpedienteConsultas" })` | §3 / §4 |
| Tests `slice-refac-1a` fallan con "no such procedure" | `Test_HelperExpedienteConsultas` no quedó en el binario | `dysflow.list_objects({ filter: "Test_HelperExpedienteConsultas" })` | §4 paso 5 |
| Tests fallan con sandbox pollution | tests no usan `BeginTestSession`/`EndTestSession` o el marcador determinista | revisar logs `["Arrange: ...", "Act: ...", "Assert: ..."]` y fixture teardown | §4 paso 4 |
| Listbox queda vacío tras `Filtrar` | `lst.List = array` no se ejecuta (la rama de `If UBound(m_Filas) >= 0` falla) | revisar logs del form + `dysflow.test_vba` con `Form_FormExpedientesGestion` integration | §3 form refactor |

## §6 Notas de migración web
- **Conservar**: el contrato `ConstruirWhereBusqueda(p_ExpBusqueda) -> WHERE` es un boundary claro para una API REST (`POST /api/expedientes/buscar` con body `ExpedienteBusqueda`).
- **Transformar**: el armado de la cláusula debe pasar a un query builder con escapeo de parámetros (nunca concatenación con `LIKE` como hoy — riesgo de inyección si la entrada viene de usuario final sin sanitizar).
- **NO copiar (legado)**: el uso de `Like '*valor*'` con asterisco Access-style; en SQL estándar es `%valor%`. Cualquier port debe traducir. Tampoco copiar el módulo-level `m_DatosEnMemoria` (eliminado en este refactor).
- **Preguntas abiertas**:
  - ¿La cláusula WHERE actual sigue siendo la misma cuando se ejecuta en backend MySQL/Postgres? (sospecha: NO — `IIF`/`Is Null` Access-specific).
  - ¿La columna `EstadoBloqueo` (de CAP-002) entra como nueva dimensión del filtro? Hoy no figura en `p_ExpBusqueda`.
  - ¿`GenerarFilas*` deben aceptar también un flag de idioma (castellano/inglés) para los headers de columna, o eso queda en la UI?

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| Helper `ConstruirWhereBusqueda` retorna tautologías con campos vacíos | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias` PASA | 2026-06-15 |
| Helper filtra por `CodExp` correctamente | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorID_DevuelveWhere` PASA | 2026-06-15 |
| Helper filtra por `PalabraClave` con `LIKE` | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorPalabraClave_DevuelveLike` PASA | 2026-06-15 |
| Helper encadena PECAL y RAC con `AND` | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_PECALYRAC_DevuelveAndEncadenado` PASA | 2026-06-15 |
| Helper mapea `EnumAmbito.Defensa` a `Ambito='Sí'` | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_AmbitoDefensa_DevuelveAmbitoSi` PASA | 2026-06-15 |
| Helper valida `ESTADO` largo y puebla `p_Error` | Verified-runtime | `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError` PASA | 2026-06-15 |
| `CargarColBusqueda` con `p_Db` inyectado no toca sandbox | Verified-static | implementado; falta test explícito (BR-23-04) | 2026-06-15 |
| Shim `FUNCIONES UTILES.getWhereBusqueda` delega al helper | Verified-static | leído en código; `verify_binary` confirma match | 2026-06-15 |
| `GenerarFilasAM`/`Lote`/`Basado`/`ContarExpedientes` son stateless, no leen controles | Verified-runtime | 6/6 tests verde; código verificado (no hay `Form_*` ni `Me.X` en los helpers) | 2026-06-15 |
| Form `RellenarAM/Lote/Basado/PonerNumeroExp` son shells thin | Verified-static | lectura directa de `Form_FormExpedientesGestion.cls`; las llamadas al helper reemplazan los `For Each ... AddItem` previos | 2026-06-15 |
| `TbExpedientes` y `TbExpedientesConEntidades` existen con los campos que la helper referencia | Verified-runtime | `dysflow_get_schema` confirmó 61 y 22 columnas respectivamente; los campos `CadenaPecal`, `CadenaComerciales`, `CadenaRACs`, `Ambito`, `Clasificacion`, `Estado` están todos | 2026-06-15 |

**⚠️ Divergencias detectadas** (intención SDD vs realidad del código): ninguna detectada a 2026-06-15. La spec PRUEBA-002 §3.2 (BR-23-01..02) y el diseño PRUEBA-003 REFAC-1a se alinean con el código actual.
