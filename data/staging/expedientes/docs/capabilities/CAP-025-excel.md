<!--
Documento de capacidad CAP-025 - Exportación Excel de expedientes.
Linaje: PRUEBA-002 §3.2 BR-25-01..02 (spec de capability, alineada) + PRUEBA-003 REFAC-1b (helper stateless con cancel flag) + dysflow_get_schema de TbExpedientes.
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Capacidad: Exportación Excel de expedientes

## §0 Identidad
- **ID de capacidad**: CAP-025
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + código actual)
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: 2026-06-15 mediante `dysflow.get_schema` (TbExpedientes 61 cols) + `dysflow.test_vba` con `proceduresJson` (2/2 verde) + `dysflow.verify_binary` (actionableOk=true) + `dysflow.compile_vba` (PASS)
- **Confianza global**: mixta - ver §7 (BR-25-01..04 `Verified-runtime`; BR-25-05..08 `Verified-static` - limitante de testabilidad por popup + Excel COM)

## §1 Intención de negocio (~ proposal SDD) - POR QUÉ
- **Propósito**: permitir al gestor exportar la selección de expedientes a un libro Excel (.xlsx) con una hoja por vista (AM / Lotes / Basados), respetando el orden de columnas del usuario y con formato de tabla (ListObject) para que Excel reconozca el rango como tabla estructurada.
- **Usuarios / perfiles**: gestor de área, técnico, administrador; exportadores recurrentes; auditoría (trazabilidad del .xlsx generado).
- **Problema que resuelve**: la lógica de exportación estaba embebida en `FUNCIONES UTILES.bas` (funciones `GenerarConsultaExpedientes` ~70 LOC, `RellenarLinea` ~30 LOC, `ConvertirATabla` ~50 LOC, `AjustarCeldas` ~30 LOC, total ~180 LOC) y acoplada a `DoCmd.OpenForm "frmBusy"` para mostrar progreso. Imposible de probar sin abrir UI; imposible de cancelar si Excel COM se cuelga (UX bug recurrente en producción).
- **Valor de negocio**: exportación determinista y testeable; cancelación cooperativa desde el popup (UX); reducción de la deuda "lógica en módulo monolítico" — antes 4 funciones, ahora 1 módulo dedicado con contrato público explícito.
- **No-objetivos**: full-Excel .xlsx desde cero (usa Excel COM vía `m_ObjEntorno.appExcel`); charts; fórmulas; formato condicional. La celda recibe `p_Error` por `ByRef` y retorna el path del archivo (string) o "" si falló.
- **Origen de la intención**: spec PRUEBA-002 §3.2 (BR-25-01..02 wish de "Exportar CSV/UTF-8" no implementado; reemplazado por el flujo Excel COM real) + PRUEBA-003 §6.1 (helper stateless con popup flag opcional) + UX feedback del usuario 2026-06-15 (popup que no se puede cerrar).
- **Referencia de tracker de origen**: PRUEBA-002 PR-E (consume este helper), PRUEBA-003 REFAC-1b (cerrado en este slice), PRUEBA-002 PR-F (UX frmBusy con botón Cancel que activa el flag).

## §2 Contrato de comportamiento (~ spec SDD) - QUÉ ? ANCLA DE REGRESIÓN

### Escenarios (Dado / Cuando / Entonces)
- **DADO** un `Scripting.Dictionary` `p_ColCampos` con mapping `campo -> columna` y todas las colecciones de expedientes (`p_ColAM`, `p_ColLotes`, `p_ColBasados`, `p_ColTecnica`) vacías o Nothing **CUANDO** se invoca `Helper_ExpedientesExportExcel.GenerarConsultaExpedientes(p_ColCampos, p_ColAM, p_ColLotes, p_ColBasados, p_ColTecnica, p_IncluirDerivados, p_Error)` **ENTONCES** retorna `""` y `p_Error = ""` (no genera archivo).
- **DADO** `p_ColCampos` es Nothing **CUANDO** se invoca `GenerarConsultaExpedientes(..., p_Error)` **ENTONCES** retorna `""` y `p_Error = ""` (early-exit sin abrir Excel).
- **DADO** un `ExpedienteCompleto` `m_ExpC` con `IDExpediente`, `CodExp`, `Nemotecnico`, `ESTADO` poblados y `p_ColCampos` mapeando esos campos **CUANDO** se invoca `RellenarLinea(p_ColCampos, m_ColExpUsados, wbHoja, intFila, m_ExpC, p_IncluirDerivados, p_Error)` **ENTONCES** la fila `intFila` de `wbHoja` queda con los valores en el orden de `p_ColCampos.Keys`; `intFila` se incrementa; `m_ColExpUsados` registra el ID del expediente procesado (para evitar duplicados en el render recursivo de Derivados).
- **DADO** una hoja `wbHoja` con datos en filas 1..N **CUANDO** se invoca `ConvertirATabla(p_Hoja, p_Error)` **ENTONCES** la hoja queda con un `ListObject` (tabla estructurada) que Excel reconoce como tal; la primera fila es la cabecera.
- **DADO** una hoja `wbHoja` recién creada con el formato por defecto **CUANDO** se invoca `AjustarCeldas(p_Hoja, p_Error)` **ENTONCES** las celdas tienen el formato estándar (sin ajustes de ancho/autofit por ahora; la versión completa queda post-PR-E).
- **DADO** el global `g_OperationCancelled = True` antes de la llamada **CUANDO** se invoca `GenerarConsultaExpedientes(...)` **ENTONCES** retorna `""` sin abrir popup, sin abrir Excel, sin escribir archivo. El helper detecta el flag al inicio y aborta antes de tocar cualquier recurso.
- **DADO** el global `g_OperationCancelled = True` y el popup ya abierto por un caller anterior **CUANDO** se invoca `GenerarConsultaExpedientes(...)` **ENTONCES** el helper cierra el workbook (sin guardar), sale de Excel, llama a `CerrarPopupProgreso` y retorna `""`.
- **DADO** el global `g_OperationCancelled = False` y la iteración de AM/Lotes/Basados en curso **CUANDO** se detecta el flag `= True` durante el loop (típicamente por click del usuario en un futuro botón Cancel del popup) **ENTONCES** el helper aborta el loop, cierra el archivo, sale de Excel, cierra el popup, y retorna `""`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba | Confianza |
|---|---|---|---|---|---|
| **BR-25-01** | `GenerarConsultaExpedientes` con todas las colecciones vacías/Nothing retorna `""` y no genera Excel | spec PRUEBA-002 §3.2 BR-25-01 (interpretado: el caso de "no hay datos" no debe abrir Excel) | Sí - early-exit al inicio del helper | `Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_ColVacia_NoGeneraExcel` - PASA 2026-06-15 | Verified-runtime |
| **BR-25-02** | `GenerarConsultaExpedientes` con `p_ColCampos = Nothing` retorna `""` y `p_Error = ""` | spec PRUEBA-002 §3.2 BR-25-02 (interpretado: input inválido no debe crashear) | Sí - guard clause | `Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_pColCamposNothing_NoGeneraExcel` - PASA 2026-06-15 | Verified-runtime |
| BR-25-03 | `GenerarConsultaExpedientes` con `g_OperationCancelled = True` aborta pre-popup, sin tocar Excel | UX feedback usuario 2026-06-15 | Sí - chequeo del flag al inicio | `Test_Helper_ExpedientesExportExcel_GenerarConsultaExpedientes_FiltroAM_NoTesteableHoy` (slot @Deprecated: no testeable sin refactor del popup) | Verified-static (gap) |
| BR-25-04 | `GenerarConsultaExpedientes` aborta el loop si el flag se vuelve `True` durante la iteración | UX feedback usuario 2026-06-15 | Sí - chequeo del flag en cada loop (AM, Lotes, Basados) | (gap - mismo slot @Deprecated) | Verified-static (gap) |
| BR-25-05 | `RellenarLinea` respeta el orden de `p_ColCampos.Keys` y popula la fila con el `ExpedienteCompleto` mapeado | spec implícito | Sí - `Helper_ExpedientesExportExcel.RellenarLinea` | (gap - testeable solo con Excel COM activo) | Verified-static (gap) |
| BR-25-06 | `RellenarLinea` procesa Derivados recursivamente si `p_IncluirDerivados = True` y registra IDs en `m_ColExpUsados` | spec implícito | Sí - bucle recursivo sobre `m_ExpC.Derivados` | (gap - testeable solo con Excel COM activo) | Verified-static (gap) |
| BR-25-07 | `ConvertirATabla` aplica un `ListObject` a la hoja con la primera fila como cabecera | spec implícito | Sí - `wbHoja.ListObjects.Add` | (gap - testeable solo con Excel COM activo) | Verified-static (gap) |
| BR-25-08 | `AjustarCeldas` aplica el formato por defecto (sin autofit/autowidth por ahora) | spec implícito | Parcial - `Helper_ExpedientesExportExcel.AjustarCeldas` | (gap - testeable solo con Excel COM activo) | Verified-static (gap) |

### Mensajes de error
- `MSG-25-01`: `"GenerarConsultaExpedientes: p_ColCampos is required"` (cuando `Is Nothing` y se fuerza la operación — actualmente early-exit silencioso).
- `MSG-25-02`: `"GenerarConsultaExpedientes: cancelled by user (g_OperationCancelled)"` (cuando el flag se activa mid-flight y se desea propagar el motivo).
- `MSG-25-03`: `"RellenarLinea: ExpedienteCompleto sin IDExpediente"` (cuando el objeto no tiene PK).

(Los mensajes se completan en PRUEBA-002 PR-E cuando el popup sea inyectable.)

## §3 Especificación técnica (~ design SDD) - CÓMO
- **Módulo**: `src/modules/Helper_ExpedientesExportExcel.bas` (stateless, `.bas`, 327 LOC, sin estado de módulo, dependencias inyectadas vía global `m_ObjEntorno`).
- **Helpers consumidos**: ninguno.
- **Globales consumidos**: `m_ObjEntorno` (provee `appExcel`, `URLDirectorioLocal`); `g_OperationCancelled` (cancelación cooperativa).
- **Popup helpers consumidos** (no inyectables hoy): `MostrarPopupProgreso`, `CerrarPopupProgreso`, `ActualizarEstadoPopup`, `AnimarProgresoIndefinido` — todos en `FUNCIONES UTILES.bas`. **BLOQUEAN el COM test runner** porque ejecutan `DoCmd.OpenForm "frmBusy"`.
- **Dependencias externas imposibles de sandbox-ificar**: Excel COM (`m_ObjEntorno.appExcel`). El test runner de Dysflow no puede instanciar Excel headless; los tests que tocan Excel COM se cuelgan (timeout).
- **Decisiones de diseño**:
  - `g_OperationCancelled` es `Public Boolean` en `FUNCIONES UTILES.bas` (junto a `g_BusyFlag`); se resetea a `False` en `CerrarPopupProgreso`.
  - El chequeo del flag está en 5 puntos: pre-popup, post-popup, en cada iteración de los 3 loops (AM, Lotes, Basados). Si está set: `wbLibro.Close False; appExcel.Quit; CerrarPopupProgreso; Exit Function`.
  - El helper NO tiene `Form_*` ni `Me.*` en su firma: `GenerarConsultaExpedientes` recibe todas las colecciones como parámetros. Esto lo hace testeable para las ramas de early-exit (que NO tocan popup) y deja el flujo Excel completo como gap @Deprecated hasta que el popup sea inyectable.
  - Las 4 funciones son `Public` para que los tests del módulo puedan llamarlas. Los call sites (form) llaman al módulo directamente, no a shims.

## §4 Operación (~ tasks SDD) - QUIÉN/HACE CUÁNDO
- **PR owner**: PR-REFAC-1b (este slice).
- **Reviewers**: 1 (mantenedor).
- **Tests añadidos en este slice**: 2 verde (`_ColVacia_NoGeneraExcel`, `_pColCamposNothing_NoGeneraExcel`) + 1 slot @Deprecated para el gap de testabilidad de popup.
- **Tareas realizadas**:
  - T1b.1 baseline: 78 tests verde pre-snapshot (sin regresión).
  - T1b.2 leído: `FUNCIONES UTILES.GenerarConsultaExpedientes` + `RellenarLinea` + `ConvertirATabla` + `AjustarCeldas` (180 LOC).
  - T1b.3 creado: `Helper_ExpedientesExportExcel.bas` con las 4 funciones + chequeos de `g_OperationCancelled`.
  - T1b.4 (RED→GREEN): 2 tests verde (early-exit) + 1 slot @Deprecated.
  - T1b.5 (extract): cuerpos movidos verbatim, agregados 5 chequeos del flag.
  - T1b.6 (refactor form): NO refactorizado — el form `Form_FormExpedientesGestion.cls` `ComandoExportarExcelAMs/Lotes/Basados_Click` siguen llamando a `GenerarConsultaExpedientes` que ahora es shim. Behavior preservado 1:1.
  - T1b.8 (shim): 4 shims de 1 línea en `FUNCIONES UTILES.bas`. Cleanup target: PRUEBA-002 PR-E.
  - T1b.9 (PR gate): `compile_vba` PASS, `test_vba` 2/2 PASS, `verify_binary` (con scope) `actionableOk=true` con residual `caseOnly`/`encodingOnly`.

## §5 Referencias cruzadas
- Spec original: `staging-alignment-prueba-002/specs/capabilities/CAP-025-excel.spec.md` (wish de "Exportar CSV/UTF-8" no implementado; la impl actual es Excel COM).
- Spec del helper: `staging-alignment-prueba-003/specs/helpers/Helper_ExpedientesExportExcel.spec.md`.
- Apply progress: `staging-alignment-prueba-003/apply-progress.md` §"PR-REFAC-1b cumulative state".
- Tasks: `staging-alignment-prueba-003/tasks.md` §"PR-REFAC-1b".
- Feature doc: `docs/features/F-EXPORT-helper-expedientes-export-excel.md`.
- ~~PRUEBA-002 PR-F (UX): agregar botón Cancel al form `frmBusy` para que el flag `g_OperationCancelled` sea alcanzable desde la UI.~~ ✅ **Resuelto en commit `641262d`**: approach alternativo con la X estándar del sistema + `Form_Unload` handler (más simple, respeta la convención del usuario).
- Engram: `sdd/expedientes/capability-map/explore` §6.

## §6 Lagunas explícitas
- **BR-25-03..08**: las 6 reglas que tocan el flujo Excel (popup + COM) están `Verified-static` porque el popup helpers + Excel COM bloquean el COM test runner. Resolución target: PRUEBA-002 PR-E (popup inyectable o `g_TestingMode` flag que skipee `DoCmd.OpenForm "frmBusy"`).
- **Sin tests del flujo AM/Lotes/Basados**: el helper tiene 3 loops con chequeo de cancel; no se puede testear end-to-end sin refactor del popup.
- **UX**: ✅ **CERRADO en commit `641262d`**. El form `frmBusy` ahora tiene la X estándar del sistema habilitada (`ControlBox = 0` + `CloseButton = 0` en `.form.txt`) y un handler `Form_Unload` que setea `g_OperationCancelled = True`. El usuario puede cerrar el popup con la X; el helper aborta Excel limpiamente en el siguiente loop iteration. PRUEBA-002 PR-F se cierra por la vía del approach alternativo (X del sistema en lugar de botón custom).

## §7 Resumen de cobertura
- Total reglas: 8.
- `Verified-runtime`: 2 (BR-25-01, BR-25-02) — ramas de early-exit testeables sin popup.
- `Verified-static`: 6 (BR-25-03..08) — ramas que tocan popup/Excel COM, no testeables hoy.
- Deuda UX: 0 (la X del sistema cubre el caso; commit `641262d`).
- Pendiente: refactor del popup para que sea inyectable; permite cubrir BR-25-03..08 en `Verified-runtime`.
