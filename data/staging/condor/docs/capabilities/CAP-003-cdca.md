# Capacidad: Captura y ciclo de vida de CD/CA

## §0 Identidad

- **ID de capacidad**: CAP-003
- **Tier**: critical
- **Estado**: active con smoke runtime diagnóstico y deuda de harness strict TDD
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary(Test_CDCA)` con `actionableOk=true` (solo `whitespaceOnly`); smoke `tests/tests.cdca.smoke.json` ejecutado con 3/3 pruebas OK (`Test_GetDatosCDCA_Happy`, `Test_EsDatosGeneralesCompleta_True`, `Test_EsDictamenRACCompleta_True`).
- **Confianza global**: mayoritariamente `Verified-static`, con smoke runtime diagnóstico para tres rutas básicas. No se eleva a `Verified-runtime` strict porque la batería CDCA conserva deuda `access-vba-tdd` v2.4.2: harness v1.9, llamadas con `Nothing` como `DAO.Database` y cardinalidad incompleta.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). CDCA no debe promocionarse a `Verified-runtime` hasta eliminar `Nothing` como `DAO.Database`, migrar el harness y completar cardinalidad de mutaciones.

**Contrato TDD vigente**: las pruebas CDCA deben migrarse a `access-vba-tdd` v2.4.2 antes de sostener nuevas afirmaciones `Verified-runtime`: `BeginTestSession`/`EndTestSession`, fixture propio schema-first, `DAO.Database` explícito en todas las llamadas bajo prueba, cardinalidad antes/después de mutaciones, manifests atómicos separados de smoke y cero mutación de `TbConfiguracionBackends`.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: registrar y gestionar Concesiones/Desviaciones o Cambios de Diseño/Alcance vinculados a una solicitud y a su expediente, con trazabilidad técnica, impacto, aprobación, dictamen RAC y decisión final.
- **Usuarios / perfiles**: técnicos, calidad, RAC, autoridad de decisión y personal administrativo que formaliza la solicitud.
- **Problema que resuelve**: estructura la información de una no conformidad o cambio de alcance para que pueda revisarse, aprobarse, rechazarse, documentarse y cerrarse de forma trazable.
- **Valor de negocio**: reduce ambigüedad técnica, ordena el ciclo de validación y deja un rastro persistente en `tbDatosCDCA` asociado a `tbSolicitudes`.
- **No-objetivos**: no documenta en detalle el workflow global, la generación documental ni adjuntos; los vincula como capacidades externas.
- **Origen de la intención**: PRD `06_Formulario_Datos_CDCA.md`, código actual y manifest de pruebas CDCA existente.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios

- **DADO** una solicitud CD/CA en preregistro **CUANDO** se guardan datos generales válidos **ENTONCES** se crea o actualiza `tbDatosCDCA` y puede avanzar a registro mediante `WorkflowServicio`. **Estado**: `Verified-static`; smoke runtime solo confirma obtención/completitud básica, no transición.
- **DADO** un registro CD/CA base **CUANDO** se guarda propuesta técnica **ENTONCES** se exigen identificación de material, causa de no conformidad y descripción/impacto. **Estado**: `Verified-static`; prueba candidata existente.
- **DADO** un registro CD/CA base **CUANDO** se guarda impacto **ENTONCES** coste debe ser `AUMENTARÁ`, `DISMINUIRÁ` o `NO VARIARÁ`; clasificación debe ser `MAYOR` o `MENOR`; y debe indicarse si el suministrador es autoridad de diseño. **Estado**: `Verified-static`.
- **DADO** una CD/CA con firmas requeridas **CUANDO** se guarda aprobación de suministrador **ENTONCES** deben existir nombres de responsable de Ingeniería y Calidad. **Estado**: `Verified-static`.
- **DADO** una CD/CA con dictamen RAC **CUANDO** la decisión no es `RECHAZADO` **ENTONCES** son obligatorios código y nombre RAC. **Estado**: `Verified-static`.
- **DADO** una decisión final **CUANDO** se guarda **ENTONCES** `decisionFinal` y `NombreFirmanteFinal` son obligatorios para considerar completa la Decisión Final. **Estado**: `Verified-runtime` limitado a helper puro Slice 3.1; guardado/persistencia CDCA sigue `Verified-static` hasta prueba fixture-first propia.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | Datos generales requieren número de contrato, referencia del suministrador, nombre/dirección del suministrador e indicador de modificación de contrato. | Código + PRD | Sí: `GuardarDatosGenerales` y `EsDatosGeneralesCompleta`. | Smoke `Test_EsDatosGeneralesCompleta_True` OK; manifest strict candidato pendiente de migración. | Verified-runtime diagnóstico / Verified-static strict |
| BR-002 | Propuesta requiere identificación de material, causa de NC y descripción/impacto. | Código | Sí: `GuardarPropuesta` y `EsDetalleCompleto`. | Código + pruebas candidatas. | Verified-static |
| BR-003 | Impacto requiere valores válidos para coste, clasificación y autoridad de diseño. | Código | Sí: `GuardarImpacto` y `EsParteTecnicaCompleta`. | Código + pruebas candidatas. | Verified-static |
| BR-004 | Aprobación suministrador requiere nombres de Ingeniería y Calidad. | Código | Sí: `GuardarAprobacionSuministrador` y completitud. | Código + pruebas candidatas. | Verified-static |
| BR-005 | Dictamen RAC no rechazado requiere código y nombre RAC. Completitud además exige `racDecision` no vacío. | Código | Sí en guardado; `EsDictamenRACCompleto` valida `racCodigo` + `racNombre` + `racDecision`. | Smoke `Test_EsDictamenRACCompleta_True` OK para caso completo; átomo strict `Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse` en `tests.vba.json` (gemelo cross-proyecto). | Verified-runtime |
| BR-006 | Decisión final requiere `decisionFinal` y `NombreFirmanteFinal` no vacíos para completitud. | Código + invariante gemelos | Sí: `EsDecisionFinalCompleta` delega en `DecisionFinalHelper_EsCompleta`. | `tests/tests.decision-final.json` → 4/4 verde para contrato puro compartido; persistencia CDCA pendiente de prueba fixture-first. | Verified-runtime parcial |
| BR-007 | Las limpiezas de RAC, aprobación y decisión final blanquean solo su bloque. | Código | Sí en repositorio. | `DatosCDCARepositorio.bas`; pruebas candidatas sin reclamo runtime. | Verified-static |
| BR-008 | Las transiciones de estado deben pasar por `WorkflowServicio`, no por SQL directo. | Arquitectura + código | Parcial: datos generales invoca `WorkflowServicio`; UI/formulario no se auditó completo aquí. | Lectura estática parcial. | Verified-static parcial |

### Validaciones

- `idSolicitud <= 0` en obtención o guardados debe tratarse como error o no encontrado según método.
- Guardado de propuesta sin campos obligatorios produce error de validación.
- Guardado de impacto con coste/clasificación fuera de catálogo produce error de validación.
- Guardado de aprobación sin nombres requeridos produce error de validación.
- Guardado de decisión final sin decisión produce error de negocio.

### Transiciones de estado

- `GuardarDatosGenerales` consulta la solicitud y, si está en `estadoPreregistro`, ejecuta transición a `estadoRegistro` por `WorkflowServicio` dentro de la conexión de trabajo.
- Las transiciones posteriores están documentadas en PRD y probablemente orquestadas por formulario/workflow, pero requieren seam de helper/servicio/ViewModel para probarlas sin depender de controles Access.

### Señales de aceptación / presencia

- Existen entidad, ViewModel, servicio y repositorio: `DatosCDCA.cls`, `DatosCDCAViewModel.cls`, `DatosCDCAServicio.cls`, `DatosCDCARepositorio.bas`.
- Existe formulario principal `Form_frmDatosCDCA.cls` y subformularios por bloque: Generales, Propuesta, Impacto, Aprobación Suministrador, Dictamen RAC y Decisión Final.
- La tabla de dominio es `tbDatosCDCA`/`TbDatosCDCA`, vinculada por `idSolicitud` a `tbSolicitudes`.
- Existe manifest `tests/tests.cdca.json` con pruebas atómicas candidatas; no se declara en verde strict en este documento. Existe smoke `tests/tests.cdca.smoke.json` con 3 pruebas ejecutadas OK como diagnóstico.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**: `Form_frmDatosCDCA` y subformularios `Form_subfrmDatosCDCA_*`. El comportamiento de UI debe considerarse `Verified-static` hasta extraer o identificar seams testables y reconciliar evidencia.
- **Puntos de entrada de código**:
  - `DatosCDCAServicio.GetDatosCDCA`, `GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC`, `GuardarDecisionFinal`.
  - `DatosCDCAServicio.EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta`.
  - `DatosCDCARepositorio.Actualizar*`, `Limpiar*`, `Guardar`, `getPorIdSolicitud`.
- **Datos afectados**:
  - `tbSolicitudes`: contexto y estado de workflow.
  - `tbDatosCDCA`: datos específicos de CD/CA por `idSolicitud`.
  - `TbExpedientes`: contexto padre en fixture y datos iniciales.
- **Dependencias**: `SolicitudServicio`, `WorkflowServicio`, `RechazoRepositorio`, `JsonHelper`, `DocumentoServicio`, `FormulariosPadreAuxiliares`, `TestHelper`.
- **Valoración de diseño**: la capa servicio/repositorio está razonablemente alineada con MVVM y persistencia granular. Hay deuda importante en pruebas: mezcla de patrones legacy (`SuiteSetup`/`SuiteTeardown`, `ForceLocalBackend`/`RestoreBackend`), llamadas con `Nothing` como `DAO.Database`, precheck sandbox incompleto y divergencias entre código y expectativas de pruebas candidatas.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar entidad, ViewModel, servicio, repositorio, formulario principal y subformularios CDCA.
2. Confirmar `tbDatosCDCA` con `idDatosCDCA` como PK e `idSolicitud` como vínculo de dominio.
3. Mantener el patrón Formulario → ViewModel → Servicio → Repositorio; no introducir SQL directo de workflow desde UI.
4. Rehacer la batería CDCA al contrato `access-vba-tdd` v2.4.2 antes de usarla como evidencia runtime.
5. Crear seams para UI: helper de mapeo de pestaña/bloque, helper de decisión de navegación y servicio de transición testeable con `DAO.Database` explícito.
6. Antes de reclamar strict runtime, verificar `Test_CDCA` con `dysflow.verify_binary`, migrar el harness y ejecutar chunks atómicos v2.4.2. El smoke actual sirve solo como diagnóstico rápido.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/DatosCDCAServicio.cls`
  - `src/modules/DatosCDCARepositorio.bas`
  - `src/modules/Test_CDCA.bas`
  - `tests/tests.cdca.json`
  - `docs/PRD/06_Formulario_Datos_CDCA.md`
- **Evidencia runtime diagnóstica**:
  - `dysflow.verify_binary(Test_CDCA)`: `actionableOk=true`; diferencia `whitespaceOnly` no funcional.
  - `tests/tests.cdca.smoke.json`: 3/3 OK (`Test_GetDatosCDCA_Happy`, `Test_EsDatosGeneralesCompleta_True`, `Test_EsDictamenRACCompleta_True`).
- **Evidencia no reclamada**: no se declara que `tests/tests.cdca.json` completo cumpla strict TDD ni esté en verde como gate de capacidad.

### Deuda de pruebas v2.4.2

- El informe transversal [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md) define el orden de migración y prevalece sobre este resumen.
- `Test_CDCA.bas` declara `access-vba-tdd v1.9` y usa `SuiteSetup`/`SuiteTeardown`; debe migrar a `BeginTestSession`/`EndTestSession`/`ResetTestSession`.
- Varias llamadas pasan `Nothing` como `DAO.Database`; v2.4.2 exige inyección explícita para probar el camino correcto.
- Hay pruebas que actualizan filas y no demuestran siempre cardinalidad antes/después.
- Debe confirmarse que `ForceLocalBackend` no muta `TbConfiguracionBackends` y que el precheck sandbox rechaza UNC, fingerprints productivos y rutas inexistentes antes de tocar datos.
- El manifest atómico existe, pero debe separarse claramente de cualquier smoke/`RunAll` y usar procedimientos públicos globales únicos.
- **Resuelto 2026-06-18** (commit pendiente): las 3 funciones `Es*Completa` del gemelo CDCA ahora validan el tercer campo correspondiente (`racDecision` en `EsDictamenRACCompleto`, `decisionFinal` en `EsAprobacionSuministradorCompleta`, `NombreFirmanteFinal` en `EsDecisionFinalCompleta`). Patrón del commit `5453f25` aplicado en `DatosCDCAServicio.cls`.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación | Ancla |
|---|---|---|---|
| No se guardan datos generales | validación incompleta, conexión incorrecta o transición falla | prueba fixture-first de `GuardarDatosGenerales` con `db` explícito y cardinalidad | §2 / §5 |
| Completitud técnica no coincide con UI | reglas duplicadas o prueba desactualizada | comparar servicio, formulario y test candidato | §2 |
| Tests CDCA verdes pero no confiables | harness legacy o datos de sandbox contaminados | auditoría v2.4.2 del módulo de tests | §5 |
| Decisión final se considera completa sin firmante | contrato de código distinto al esperado por prueba/PRD | decisión de producto + ajuste de código o tests | §2 / §7 |

## §6 Notas de migración web

- **Conservar**: separación por bloques, validaciones de servicio, persistencia granular, transición por servicio de workflow, trazabilidad de rechazos.
- **Transformar**: subformularios Access a pasos/rutas web; mensajes modales a validaciones de formulario; ViewModel como DTO explícito.
- **NO copiar**: dependencia de `Nothing` para que el servicio llame a `getdb()`, lógica de UI no testeable en eventos de formulario, harness legacy de pruebas.

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| CDCA tiene entidad, ViewModel, servicio, repositorio, formulario principal y subformularios por bloque. | Verified-static | Archivos `DatosCDCA*` y `Form_*CDCA*`. | 2026-06-15 |
| `tbDatosCDCA` es la tabla de datos específicos CDCA vinculada por `idSolicitud`. | Verified-static | `DatosCDCARepositorio.bas`; PRD CDCA. | 2026-06-15 |
| Datos generales, propuesta, impacto, aprobación, RAC y decisión final se guardan por métodos de servicio. | Verified-static | `DatosCDCAServicio.cls`. | 2026-06-15 |
| Existe una batería CDCA amplia como candidata de evidencia. | Verified-static documental | `tests/tests.cdca.json`; `Test_CDCA.bas`. | 2026-06-15 |
| El smoke CDCA básico pasa en el binario actual. | Verified-runtime diagnóstico | `tests/tests.cdca.smoke.json`: 3/3 OK; `verify_binary(Test_CDCA)` actionableOk=true. | 2026-06-15 |
| La batería CDCA actual cumple `access-vba-tdd` v2.4.2. | Divergent / deuda | `Test_CDCA.bas` conserva patrón v1.9, `SuiteSetup` y llamadas con `Nothing`. | 2026-06-15 |
| Phase 0 clasifica `Form_frmDatosCDCA` como Tier 1 y los subformularios CDCA por bloques como Tier 2/Tier 3 según señales de lógica. | Verified-static / deuda Phase 0 | `audit-e2e-thin-forms-phase-0.md` §2-§4; no se modificó `src/` ni se ejecutó Access. | 2026-06-26 |
| El UAT CDCA queda bloqueado hasta migrar el harness strict, ejecutar átomo verde con `DAO.Database` explícito y firmar `ref`. | Divergent / pendiente | `audit-e2e-thin-forms-phase-0.md` §7-§8; la matriz exige `Test_CDCA_GuardarDatosGenerales_UsesInjectedDbAndCardinality`. | 2026-06-26 |
| `EsDecisionFinalCompleta` exige firmante final. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con átomo `Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse` en `tests.vba.json` (cross-proyecto PCSUB↔CDCA por invariante gemelos). | 2026-06-15 (original) → 2026-06-18 (resolución) |
| `EsDictamenRACCompleto` exige `racDecision`. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con átomo `Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse` en `tests.vba.json` (cross-proyecto PCSUB↔CDCA por invariante gemelos). | 2026-06-15 (original) → 2026-06-18 (resolución) |
| El contrato puro compartido de Decisión Final está alineado para PC, PCSUB, CDCA y CDCASUB. | Verified-runtime limitado a helper puro Slice 3.1 | `tests/tests.decision-final.json` → 4/4 verde; `DecisionFinalHelper_EsCompleta` usado por `DatosCDCAServicio.EsDecisionFinalCompleta`. | 2026-06-27 |

**Divergencias pendientes de revisión humana**:

- Confirmar si RAC completo debe exigir `racDecision`; si sí, crear SDD/fix y prueba v2.4.2.
- Confirmar si decisión final completa debe exigir `NombreFirmanteFinal`; si sí, crear SDD/fix y prueba v2.4.2.
- Decidir si se migra o se retira el `RunAll` CDCA como smoke separado para evitar mezclarlo con el manifest atómico.
