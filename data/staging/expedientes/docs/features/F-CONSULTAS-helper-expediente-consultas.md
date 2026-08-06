<!--
Hoja de feature — Helper_ExpedienteConsultas.
Capa: docs/features/ (segundo nivel; el primero es docs/capabilities/).
Representa el artefacto de código (helper) introducido en PRUEBA-003 REFAC-1a, ya enlazado a CAP-023.
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Feature: `Helper_ExpedienteConsultas`

## §0 Identidad
- **ID de feature**: F-CONSULTAS
- **Helper introducido en**: PRUEBA-003 REFAC-1a (PR-REFAC-1a)
- **Capability destino**: [CAP-023 Búsqueda avanzada](../capabilities/CAP-023-busqueda-avanzada.md)
- **Módulo**: `src/modules/Helper_ExpedienteConsultas.bas` (stateless, `.bas`)
- **Estado**: active (en `staging` desde commit `b7765ef`)
- **Source**: sdd
- **Confianza**: mixta — ver §6 (BR-23-01..06 `Verified-runtime`; BR-23-07..08 `Verified-static`)

## §1 Propósito y motivo
- **Propósito técnico**: extraer las funciones de query `getWhereBusqueda`, `getColBusqueda`, `getColBusquedaTecnica` (originalmente en `FUNCIONES UTILES.bas` líneas 254–847) a un módulo dedicado, stateless, testeable de forma aislada, y con inyección opcional de `DAO.Database` para tests deterministas contra backend sandbox.
- **Motivo de existencia**: PRUEBA-002 §3.2 BR-23-01..08 requiere tests de la búsqueda avanzada; la lógica embebida en el form (3,474 líneas) y en `FUNCIONES UTILES.bas` no es testeable sin abrir UI. Convención `access-vba-tdd` §1.6 prohíbe `módulo = función pública`; este helper expone `ConstruirWhereBusqueda` (no `Helper_ExpedienteConsultas`) y cumple la convención.

## §2 Contrato público (firmas)
```vb
' Construcción de WHERE
Public Function ConstruirWhereBusqueda( _
    ByVal p_ExpBusqueda As ExpedienteBusqueda, _
    Optional ByRef p_Error As String) As String

Public Function ConstruirWhereBusquedaTecnica( _
    ByVal p_ExpBusqueda As ExpedienteBusqueda, _
    Optional ByRef p_Error As String) As String

' Carga de colecciones
Public Function CargarColBusqueda( _
    ByVal p_ExpBusqueda As ExpedienteBusqueda, _
    ByVal p_ParaAM As Boolean, _
    ByVal p_ParaLote As Boolean, _
    ByVal p_ParaBasado As Boolean, _
    ByVal p_ParaExpediente As Boolean, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Scripting.Dictionary

Public Function CargarColBusquedaTecnica( _
    ByVal p_ExpBusqueda As ExpedienteBusqueda, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Scripting.Dictionary

' Helper interno DAO-injectable
Private Function ObtenerExpedientesCompletosDesdeDb( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
```

Convenciones aplicadas:
- `Option Explicit` + `Option Compare Database` (skill `vba-access`).
- `ByRef p_Error` como último parámetro (AGENTS.md del proyecto).
- DAO inyectable como `Optional ByVal p_Db As DAO.Database = Nothing` con fallback a `getdb()` (skill `access-vba-tdd` §5.4).
- Sin `MsgBox`/`DoCmd`/`Me.` (helper callable sin abrir form).

## §3 Callers actuales
- `FUNCIONES UTILES.getWhereBusqueda` (shim): delega a `Helper_ExpedienteConsultas.ConstruirWhereBusqueda`. Mantenido por 1 release; cleanup en PRUEBA-002 PR-E.
- `FUNCIONES UTILES.getColBusqueda` (shim): delega a `Helper_ExpedienteConsultas.CargarColBusqueda`.
- `FUNCIONES UTILES.getColBusquedaTecnica` (shim): delega a `Helper_ExpedienteConsultas.CargarColBusquedaTecnica`.
- `Form_FormExpedientesGestion.ComandoBuscar_Click` y `txtFiltro_AfterUpdate` (vía el shim).
- Tests: `src/modules/Test_HelperExpedienteConsultas.bas` (6 tests atómicos en `tests/tests.vba.json` con tag `slice-refac-1a`).

## §4 Receta de reconstrucción
1. Crear `src/modules/Helper_ExpedienteConsultas.bas` con las firmas de §2.
2. Implementar `ConstruirWhereBusqueda` con la lógica de tautologías/dimensiones.
3. Implementar `ObtenerExpedientesCompletosDesdeDb` con `DAO.Recordset` + `dbOpenSnapshot`, iterando `m_ExpC.ColCampos` para `SetPropiedad`.
4. Implementar `CargarColBusqueda` con la rama `If p_Db Is Nothing Then ... Else ... End If`.
5. **Importar** → `dysflow.import_modules({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "Test_HelperExpedienteConsultas"] })`.
6. **Compilar** → `dysflow.compile_vba({ projectId: "expedientes" })` → `ok: true`.
7. **Verificar** → `dysflow.verify_binary({ projectId: "expedientes", moduleNames: ["Helper_ExpedienteConsultas", "FUNCIONES UTILES"], diff: true })` → `actionableOk: true`.
8. **Demostrar** → `dysflow.test_vba({ projectId: "expedientes", proceduresJson: "[{\"procedure\":\"Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias\"}, ...]" })` → 6/6 PASS.

## §5 Evidencia
- **Tests verdes (2026-06-15)**:
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias` — 3.3s
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorID_DevuelveWhere` — 3.1s
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorPalabraClave_DevuelveLike` — 3.3s
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_PECALYRAC_DevuelveAndEncadenado` — 3.0s
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_AmbitoDefensa_DevuelveAmbitoSi` — 3.3s
  - `Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError` — 3.0s
- **Sincronización fuente↔binario**: `verify_binary` clean el 2026-06-15; solo `caseOnly` residual en `.Value`/`.value` (no accionable).
- **Commits**: `b7765ef` (helper + 6 tests + form refactor), `8118e12` (trim slice size); slice T1a.7-T1a.9 pendiente de cierre por presupuesto 400 (siguiente paso).

## §6 Lagunas y confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| 6 tests de `slice-refac-1a` en verde | Verified-runtime | `dysflow.test_vba` 2026-06-15 | 2026-06-15 |
| Shim en `FUNCIONES UTILES` delega correctamente | Verified-static | leído en código; `verify_binary` confirma match | 2026-06-15 |
| `ObtenerExpedientesCompletosDesdeDb` no toca sandbox cuando `p_db` está inyectado | Verified-static | implementado; falta test explícito T1a.4 | 2026-06-15 |
| `ConstruirWhereBusqueda` con `p_Valor = ""` retorna tautología pura (sin OR con nada) | Verified-static | implícito en BR-23-01; falta test explícito | 2026-06-15 |
| Cleanup del shim en `FUNCIONES UTILES` | Intended | target PRUEBA-002 PR-E | 2026-06-15 |
| Forma-coupling: el form sigue llamando al shim en vez de al helper directo | Intended | cleanup diferido a PRUEBA-002 PR-E | 2026-06-15 |

## §7 Trazabilidad a capabilities

| Capability destino | Reglas cubiertas | Doc de capacidad |
|---|---|---|
| CAP-023 Búsqueda avanzada | BR-23-01..06 ✅ Verified-runtime, BR-23-07..08 ⚠️ Verified-static | [CAP-023](../capabilities/CAP-023-busqueda-avanzada.md) |
| CAP-024 Búsqueda técnica | BR-24-01 ⚠️ Verified-static (target post-PR-E) | [CAP-024](../capabilities/CAP-024-busqueda-tecnica.md) (pendiente) |
