# Auditoría Phase 0 — formularios finos E2E

## 1. Alcance y contrato

- **Cambio SDD**: `phase-0-e2e-thin-forms-consolidation`
- **Work-unit**: WU1 + WU2 + WU3 + WU4 — inventario documental transversal, reconciliación PCSUB contra `staging`, matriz TDD↔UAT con ledger de capacidades y contratos para extracciones futuras.
- **Ámbito**: `Form_frmDatos*` y `Form_subfrmDatos*` de PC, CDCA, CDCASUB y PCSUB.
- **No incluido**: extracción de código VBA, cambios en `src/`, importación a Access, compilación o ejecución de `test_vba`.
- **Evidencia usada**: búsqueda estática focalizada en `src/forms`, `src/classes` y `src/modules` sobre `GuardarDesdeSubform`, `cmdGuardar_Click`, `cmdCargar*`, `MsgBox`, DAO, Workflow, `DoCmd` y `getdb()`.

## 2. `form_tiering`

| Objeto | Familia | Tier | Motivo | Handlers clave | Deuda / siguiente acción | Capability ref |
|---|---|---:|---|---|---|---|
| `Form_frmDatosPC` | PC | 1 | Orquestador de guardado y workflow con `MsgBox`, transacción DAO, `getdb()`, `DoCmd` y transición de estado. | `GuardarDesdeSubform`, `ProcesarCierreFormalizacion`, navegación de pestañas. | Extraer guardado/orquestación/prompt a helper testeable antes de cualquier UAT. | CAP-004, CAP-007 |
| `Form_frmDatosPCSUB` | PCSUB | 1 | Mismo rol que PC; baseline candidato todavía no reconciliado con `staging`. | `GuardarDesdeSubform`, `ProcesarCierreFormalizacion`, navegación de pestañas. | Reconciliar cobertura PCSUB en WU2 antes de tratarla como patrón. | CAP-001, CAP-007 |
| `Form_frmDatosCDCA` | CDCA | 1 | Orquestador equivalente; mantiene transacción DAO local y decisiones de workflow en el formulario. | `GuardarDesdeSubform`, `ProcesarCierreFormalizacion`, navegación de pestañas. | Extraer por gemelo manteniendo paridad con PC/CDCASUB/PCSUB. | CAP-003, CAP-007 |
| `Form_frmDatosCDCASUB` | CDCASUB | 1 | Orquestador equivalente; parte del guardado ya delega en servicio pero el prompt/workflow sigue inline. | `GuardarDesdeSubform`, `ProcesarCierreFormalizacion`, navegación de pestañas. | Separar decisión testeable de presentación UI y registrar átomos antes de tocar código. | CAP-005, CAP-007 |
| `Form_subfrmDatosPC_Generales` | PC | 2 | Subformulario con lectura/escritura de VM y defaults con `DoCmd.Hourglass`; guardado delegado al padre. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Extraer carga de defaults si contiene regla; mantener UI como adaptador fino. | CAP-004 |
| `Form_subfrmDatosPCSUB_Generales` | PCSUB | 2 | Gemelo de PC con defaults y guardado delegado al padre. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Alinear con PC/CDCA/CDCASUB en slice posterior. | CAP-001 |
| `Form_subfrmDatosCDCA_Generales` | CDCA | 2 | Defaults con `DoCmd.Hourglass`; guardado delegado. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Extraer helper de defaults si la carga usa datos de expediente. | CAP-003 |
| `Form_subfrmDatosCDCASUB_Generales` | CDCASUB | 2 | Defaults con `DoCmd.Hourglass`; guardado delegado. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Mantener como gemelo de CDCA y registrar divergencias. | CAP-005 |
| `Form_subfrmDatosPC_Propuesta` | PC | 3 | Handler principal solo delega `Me.Parent.GuardarDesdeSubform`. | `cmdGuardar_Click`. | Deuda Tier 3 documentada: conservar hasta que el padre se extraiga. | CAP-004 |
| `Form_subfrmDatosPCSUB_Propuesta` | PCSUB | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-001 |
| `Form_subfrmDatosCDCA_Propuesta` | CDCA | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-003 |
| `Form_subfrmDatosCDCASUB_Propuesta` | CDCASUB | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-005 |
| `Form_subfrmDatosPC_Impacto` | PC | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Conservar como adaptador fino mientras no tenga reglas propias. | CAP-004 |
| `Form_subfrmDatosPCSUB_Impacto` | PCSUB | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-001 |
| `Form_subfrmDatosCDCA_Impacto` | CDCA | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-003 |
| `Form_subfrmDatosCDCASUB_Impacto` | CDCASUB | 3 | Handler principal solo delega al padre. | `cmdGuardar_Click`. | Igual que PC. | CAP-005 |
| `Form_subfrmDatosPC_DictamenRAC` | PC | 2 | Carga de datos predeterminados con `MsgBox` si no hay RAC y `DoCmd.Hourglass`. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Extraer decisión “sin RAC” y carga de valores a helper con seam de prompt. | CAP-004 |
| `Form_subfrmDatosPCSUB_DictamenRAC` | PCSUB | 2 | Mismo patrón que PC. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Alinear con PC/CDCA/CDCASUB. | CAP-001 |
| `Form_subfrmDatosCDCA_DictamenRAC` | CDCA | 2 | Mismo patrón que PC, con comentario que el feedback vive en el padre. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Extraer carga/validación de RAC a helper compartido o por familia. | CAP-003 |
| `Form_subfrmDatosCDCASUB_DictamenRAC` | CDCASUB | 2 | Mismo patrón que CDCA. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Alinear con CDCA. | CAP-005 |
| `Form_subfrmDatosPC_AprobacionSuministrador` | PC | 2 | Defaults con `DoCmd.Hourglass`; guardado delegado. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Extraer carga de aprobadores si contiene regla de negocio. | CAP-004 |
| `Form_subfrmDatosPCSUB_AprobacionSuministrador` | PCSUB | 2 | Mismo patrón que PC. | `cmdGuardar_Click`, `cmdCargarDefaults_Click`. | Revisar baseline PCSUB en WU2. | CAP-001 |
| `Form_subfrmDatosCDCA_AprobacionSuministrador` | CDCA | 2 | Defaults con `DoCmd.Hourglass`; guardado delegado. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Extraer carga de aprobadores por gemelo. | CAP-003 |
| `Form_subfrmDatosCDCASUB_AprobacionSuministrador` | CDCASUB | 2 | Defaults con `DoCmd.Hourglass`; guardado delegado. | `cmdGuardar_Click`, `cmdCargarDatosPredeterminados_Click`. | Alinear con CDCA. | CAP-005 |
| `Form_subfrmDatosPC_DecisionFinal` | PC | 2 | Validación UI inline con `MsgBox` para decisión obligatoria antes de delegar. | `cmdGuardar_Click`. | Extraer validación a helper o servicio y dejar `MsgBox` en formulario. | CAP-004 |
| `Form_subfrmDatosPCSUB_DecisionFinal` | PCSUB | 2 | Mismo patrón que PC. | `cmdGuardar_Click`. | Alinear con PC/CDCA/CDCASUB. | CAP-001 |
| `Form_subfrmDatosCDCA_DecisionFinal` | CDCA | 2 | Mismo patrón que PC. | `cmdGuardar_Click`. | Alinear con PC/CDCASUB/PCSUB. | CAP-003 |
| `Form_subfrmDatosCDCASUB_DecisionFinal` | CDCASUB | 2 | Mismo patrón que CDCA. | `cmdGuardar_Click`. | Alinear con CDCA. | CAP-005 |

## 3. `inline_handler_audit`

| Formulario | Handler | Señales detectadas | Riesgo | Helper objetivo |
|---|---|---|---|---|
| `Form_frmDatosPC` | `GuardarDesdeSubform` | `MsgBox`, DAO, `BeginTrans`, `CommitTrans`, `Rollback`, Workflow, `DoCmd`, `getdb()` | Alto: mezcla lectura UI, persistencia, transacción, prompts y transición. | `DatosPCGuardarHelper.GuardarDesdeSubformPlan(...)` + servicios existentes. |
| `Form_frmDatosPCSUB` | `GuardarDesdeSubform` | `MsgBox`, DAO, `BeginTrans`, `CommitTrans`, `Rollback`, Workflow, `DoCmd`, `getdb()` | Alto: gemelo crítico y baseline candidato no reconciliado. | `DatosPCSUBGuardarHelper.GuardarDesdeSubformPlan(...)` o helper compartido parametrizado. |
| `Form_frmDatosCDCA` | `GuardarDesdeSubform` | `MsgBox`, DAO, `BeginTrans`, `CommitTrans`, `Rollback`, Workflow, `DoCmd`, `getdb()` | Alto: patrón equivalente con deuda de formulario grueso. | `DatosCDCAGuardarHelper.GuardarDesdeSubformPlan(...)`. |
| `Form_frmDatosCDCASUB` | `GuardarDesdeSubform` | `MsgBox`, Workflow, `DoCmd`, rollback comentado/local; servicio gestiona parte de la transacción. | Alto: la delegación parcial no elimina prompts ni workflow inline. | `DatosCDCASUBGuardarHelper.GuardarDesdeSubformPlan(...)`. |
| `Form_frmDatos*` | `ProcesarCierreFormalizacion` | `MsgBox`, Workflow, `DoCmd`, selección de documento/adjunto. | Alto: cierre formal con decisión de usuario y transición; no debe ir a UAT sin átomo. | `CierreFormalizacionHelper.ProcesarCierrePlan(...)`. |
| `Form_subfrmDatos*_DecisionFinal` | `cmdGuardar_Click` | `MsgBox`, `Me.Parent.GuardarDesdeSubform`. | Medio: validación obligatoria en UI; el guardado real está en el padre. | `DecisionFinalFormHelper.ValidarDecisionFinal(...)` con seam de prompt. |
| `Form_subfrmDatos*_DictamenRAC` | `cmdCargarDatosPredeterminados_Click` | `MsgBox`, `DoCmd.Hourglass`, lectura de datos iniciales. | Medio: regla “sin RAC asignado” y carga de defaults mezcladas con UI. | `DictamenRACDefaultsHelper.CargarDefaults(...)` con `p_PromptResult`. |
| `Form_subfrmDatos*_Generales` | `cmdCargarDefaults_Click` | `DoCmd.Hourglass`, carga de valores iniciales. | Medio: posible regla de datos iniciales según expediente/contrato. | `DatosGeneralesDefaultsHelper.CargarDefaults(...)`. |
| `Form_subfrmDatos*_AprobacionSuministrador` | `cmdCargar*` | `DoCmd.Hourglass`, carga de firmantes/defaults. | Medio: reglas de firmantes deben poder probarse sin abrir formulario. | `AprobacionSuministradorDefaultsHelper.CargarDefaults(...)`. |
| `Form_subfrmDatos*_Propuesta` / `*_Impacto` | `cmdGuardar_Click` | `Me.Parent.GuardarDesdeSubform`; sin señales fuertes propias en la búsqueda focalizada. | Bajo: adaptadores finos, dependen de extraer el padre. | Sin helper propio salvo que aparezca lógica en lectura posterior. |

## 4. `helper_map`

| Familia | Helper existente | Helper objetivo | Firma honesta propuesta | Seam | Cobertura por gemelo |
|---|---|---|---|---|---|
| PC | Servicios/repositorios `DatosPCServicio` y `DatosPCRepositorio` con métodos `Guardar*` y `Optional db As DAO.Database`. | `DatosPCGuardarHelper` | `Function GuardarDesdeSubformPlan(ByVal nombreSub As String, ByRef vm As DatosPCViewModel, ByVal prompt As VbMsgBoxResult, Optional ByRef db As DAO.Database = Nothing) As GuardarResultado` | `db` para persistencia real; `prompt` para decisiones `MsgBox`; servicio inyectable si se formaliza interfaz. | Parcial: servicio existe; falta helper de formulario y átomos focales. |
| PCSUB | `DatosPCSUBServicio` y `DatosPCSUBRepositorio`; existe `src/modules/DatosPCSUBGuardarHelper.bas` sin contenido en el árbol actual. | `DatosPCSUBGuardarHelper` | Misma forma que PC, ajustada a `DatosPCSUBViewModel`. | `db` + `prompt`; WU2 debe reconciliar si la rama candidata ya aporta contratos. | Pendiente: baseline no tratado como `Verified-runtime` hasta reconciliar. |
| CDCA | `DatosCDCAServicio` y `DatosCDCARepositorio` con `GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, etc. | `DatosCDCAGuardarHelper` | `Function GuardarDesdeSubformPlan(ByVal nombreSub As String, ByRef vm As DatosCDCAViewModel, ByVal prompt As VbMsgBoxResult, Optional ByRef db As DAO.Database = Nothing) As GuardarResultado` | `db` + `prompt`; posible interfaz para Workflow si se separa de servicio. | Parcial: servicios con `db`; falta helper de orquestación/prompt. |
| CDCASUB | `DatosCDCASUBServicio` y `DatosCDCASUBRepositorio` con métodos `Guardar*` y `Optional db As DAO.Database`. | `DatosCDCASUBGuardarHelper` | Misma forma que CDCA, ajustada a `DatosCDCASUBViewModel`. | `db` + `prompt`; mantener paridad con CDCA. | Parcial: servicios con `db`; falta helper de orquestación/prompt. |
| Transversal cierre | `WorkflowServicio`, `AdjuntosServicio`, formularios padre. | `CierreFormalizacionHelper` | `Function PrepararCierreFormalizacion(ByVal idSolicitud As Long, ByVal adjuntoSeleccionado As String, ByVal prompt As VbMsgBoxResult, Optional ByRef db As DAO.Database = Nothing) As CierreFormalizacionResultado` | `db` para workflow/adjuntos; `prompt` y ruta de archivo como seams; no abrir UI en helper. | Sin cobertura auditada en este slice. |
| Transversal defaults | Servicios `GetDatosGeneralesIniciales`, `GetDatosAprobacionIniciales` por familia. | `DatosDefaultsHelper` o helpers por bloque | `Function ObtenerDefaults(ByVal familia As String, ByVal bloque As enumBloqueFormulario, ByVal idSolicitud As Long, Optional ByRef db As DAO.Database = Nothing) As Object` | `db`; sin `prompt` salvo caso DictamenRAC sin RAC. | Parcial y desigual; requiere átomos por familia. |
| Dictamen RAC | Servicios/repositorios de expediente y RAC. | `DictamenRACDefaultsHelper` | `Function PrepararDictamenRACDefaults(ByVal idExpediente As Long, Optional ByRef db As DAO.Database = Nothing) As DictamenRACDefaultsResultado` | `db`; `prompt` solo como salida de resultado, no modal. | Requiere cuatro átomos gemelos o un contrato parametrizado. |
| Decisión final | Servicios `GuardarDecisionFinal` por familia. | `DecisionFinalFormHelper` | `Function ValidarDecisionFinal(ByVal decisionFinal As String) As ValidacionResultado` | Puro; sin `db`; el formulario muestra el mensaje. | Misma regla en cuatro gemelos; buen candidato para helper compartido. |

## 5. Evidencia estática relevante

- Los cuatro formularios padre contienen `GuardarDesdeSubform` y muestran señales de formulario grueso.
- `Form_frmDatosPC`, `Form_frmDatosPCSUB` y `Form_frmDatosCDCA` abren explícitamente `ws.OpenDatabase(getdb().name, ...)` y gestionan transacciones en el formulario.
- `Form_frmDatosCDCASUB` documenta que parte del `CommitTrans` vive en el servicio, pero mantiene prompts y transiciones en el formulario.
- Los subformularios `Propuesta` e `Impacto` aparecen como adaptadores finos en la búsqueda focalizada; se clasifican Tier 3 salvo evidencia posterior.
- `src/modules/DatosPCSUBGuardarHelper.bas` existe en el árbol de trabajo pero está vacío; además aparece como archivo no versionado, por lo que esta auditoría no lo toma como implementación existente.

## 6. Reconciliación PCSUB coverage — WU2

### 6.1 Ancestros `b946c15..59b4438` frente a `staging`

Comandos de solo lectura ejecutados desde `C:\00repos\codigo\00_CONDOR_staging`:

- `git rev-parse --verify staging`
- `git rev-parse --verify feature/pcsub-workflow-helper-coverage`
- `git log --oneline --reverse b946c15^..59b4438`
- `git merge-base --is-ancestor <commit> staging`
- `git diff --name-status staging..feature/pcsub-workflow-helper-coverage`
- `git diff --stat staging..feature/pcsub-workflow-helper-coverage`
- `git diff --numstat staging..feature/pcsub-workflow-helper-coverage`
- `git diff --check staging..feature/pcsub-workflow-helper-coverage`

Refs observadas:

| Ref | SHA |
|---|---|
| `staging` | `bd319a2f97bdd0124a04e04b4d27d461a15ea547` |
| `feature/pcsub-workflow-helper-coverage` | `59b44388078425562d4f0bdbeb3512a730908785` |
| `origin/feature/pcsub-workflow-helper-coverage` | `59b44388078425562d4f0bdbeb3512a730908785` |

Resultado de ancestros:

| Commit | Resumen | ¿Ancestro de `staging`? | Evidencia |
|---|---|---:|---|
| `b946c157d3f69b65a86c91ad3e23efaed580847e` | `docs(audit): pcsub form handler audit` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |
| `a9edf7377c94240ec8f80ac93a5dc0e3e9cbd30c` | `feat(pcsub): extract WizardHelper + Form_Load seam (1 atom)` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |
| `b9beff0c7a1ad48e422e36141c703e53a97ffc89` | `feat(pcsub): extract GuardarHelper + 6 subform wiring (4 atoms, 34/34 verde)` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |
| `7210788df300ba016ee70d61f9c16cc62af58cdc` | `feat(pcsub): extract CierreHelper + 8 atoms (3 verde, 5 sandbox-blocked)` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |
| `86f61193c61c4171d8e79a03df4500fd15170a4e` | `feat(pcsub): extract TipoSolicitudHelper + 5 atoms (5/5 verde)` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |
| `59b44388078425562d4f0bdbeb3512a730908785` | `docs(pcsub): cap-001 flip + archive traceability report` | No | `git merge-base --is-ancestor` devolvió salida distinta de cero. |

Conclusión: ningún commit del tramo `b946c15^..59b4438` está integrado en `staging`. La rama candidata existe localmente y en `origin`, pero su contenido no puede tratarse como baseline de `staging`.

### 6.2 Clasificación por archivo del diff `staging..feature/pcsub-workflow-helper-coverage`

Resumen del diff: `30 files changed, 3139 insertions(+), 352 deletions(-)`. La clasificación usa estas reglas:

- **Adoptable**: candidato reutilizable en una slice futura, sin afirmarlo como integrado ni validado en `staging`.
- **Conflictivo**: toca código VBA, formularios, harness de tests o rutas con solape local; requiere adopción explícita, importación, compilación y `test_vba` antes de cualquier cierre.
- **Obsoleto**: afirmación documental que ya no representa el estado real de `staging` o queda sustituida por esta reconciliación.

| Archivo | Estado diff | Clasificación | Motivo |
|---|---|---|---|
| `docs/ERD/condor_datos.md` | M | Adoptable | Cambio documental mínimo; debe verificarse contra ERD actual antes de incorporarlo. |
| `docs/capabilities/CAP-001-pcsub.md` | M | Obsoleto | Eleva PCSUB a `Verified-runtime` con evidencia de la rama candidata, no de `staging`; queda sustituido por el estado `Divergent` de esta WU2 hasta adopción real. |
| `docs/capabilities/audit-pcsub-form-helpers.md` | A | Adoptable | Inventario útil como referencia histórica de la rama, pero no sustituye el audit Phase 0 transversal. |
| `src/classes/SolicitudServicio.cls` | M | Conflictivo | Cambio VBA de servicio compartido; requiere slice de implementación y pruebas. |
| `src/classes/WorkflowServicio.cls` | M | Conflictivo | Cambio VBA de workflow; alto impacto y requiere pruebas de transición. |
| `src/forms/Form_frmAltaSolicitud.cls` | M | Conflictivo | Cambio en formulario; requiere importación y compilación manual de Access antes de confiar en tests. |
| `src/forms/Form_frmAltaSolicitud.form.txt` | M | Conflictivo | Cambio de layout/formulario; requiere sincronización de formulario y validación manual. |
| `src/forms/Form_frmDatosPCSUB.cls` | M | Conflictivo | Refactor del formulario padre PCSUB; no es ancestro de `staging`. |
| `src/forms/Form_frmDatosPCSUB.form.txt` | M | Conflictivo | Cambio de formulario no adoptado en `staging`. |
| `src/forms/Form_subfrmDatosPCSUB_AprobacionSuministrador.cls` | M | Conflictivo | Cambio de code-behind de subformulario. |
| `src/forms/Form_subfrmDatosPCSUB_AprobacionSuministrador.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/forms/Form_subfrmDatosPCSUB_DecisionFinal.cls` | M | Conflictivo | Cambio de code-behind con rama de cierre; requiere compile manual. |
| `src/forms/Form_subfrmDatosPCSUB_DecisionFinal.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/forms/Form_subfrmDatosPCSUB_DictamenRAC.cls` | M | Conflictivo | Cambio de code-behind de subformulario. |
| `src/forms/Form_subfrmDatosPCSUB_DictamenRAC.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/forms/Form_subfrmDatosPCSUB_Generales.cls` | M | Conflictivo | Cambio de code-behind de subformulario. |
| `src/forms/Form_subfrmDatosPCSUB_Generales.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/forms/Form_subfrmDatosPCSUB_Impacto.cls` | M | Conflictivo | Cambio de code-behind de subformulario. |
| `src/forms/Form_subfrmDatosPCSUB_Impacto.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/forms/Form_subfrmDatosPCSUB_Propuesta.cls` | M | Conflictivo | Cambio de code-behind de subformulario. |
| `src/forms/Form_subfrmDatosPCSUB_Propuesta.form.txt` | M | Conflictivo | Cambio de layout/formulario. |
| `src/modules/DatosPCSUBCierreHelper.bas` | A | Adoptable | Helper candidato nuevo; exige RED/import/compile/test antes de adopción. |
| `src/modules/DatosPCSUBGuardarHelper.bas` | A | Conflictivo | Helper candidato nuevo, pero en el árbol actual existe el mismo path como archivo no versionado; adoptarlo sobrescribiría trabajo local si no se reconcilia primero. |
| `src/modules/DatosPCSUBWizardHelper.bas` | A | Adoptable | Helper candidato nuevo; no integrado en `staging`. |
| `src/modules/DatosTipoSolicitudHelper.bas` | A | Adoptable | Helper candidato nuevo para filtrado de tipo; requiere validación gemela. |
| `src/modules/TestHelper.bas` | M | Conflictivo | Modifica harness compartido de tests; riesgo transversal. |
| `src/modules/Test_DatosTipoSolicitudHelper_PCSUB_Strict.bas` | A | Adoptable | Átomos candidatos para helper nuevo; dependen de adoptar el helper. |
| `src/modules/Test_PCSUB_Cierre_Strict.bas` | A | Adoptable | Átomos candidatos, aunque la propia rama declara 5 casos bloqueados por sandbox. |
| `src/modules/Test_PCSUB_Strict.bas` | M | Conflictivo | Modifica suite PCSUB existente; no puede mezclarse sin revalidar todo el manifest. |
| `tests/tests.pcsub.json` | M | Conflictivo | El manifest referencia procedimientos que no existen en `staging`; fallaría sin adoptar módulos de la rama. |

### 6.3 Estado de cobertura PCSUB según evidencia actual

Estado para Phase 0: **`Divergent`**.

Evidencia:

- Ninguno de los seis commits de `feature/pcsub-workflow-helper-coverage` es ancestro de `staging`.
- El diff frente a `staging` añade o modifica 30 archivos, incluidos formularios, helpers, tests y manifest.
- `tests/tests.pcsub.json` de la rama candidata referencia átomos que no existen en `staging` sin adoptar módulos nuevos.
- `CAP-001-pcsub.md` de la rama candidata contiene afirmaciones `Verified-runtime` apoyadas en ejecuciones de esa rama; no son evidencia runtime del `staging` actual.
- El árbol de trabajo actual contiene cambios no confirmados, incluido `src/modules/DatosPCSUBGuardarHelper.bas` como archivo no versionado, que solapa con un archivo añadido por la rama candidata.

Por tanto, PCSUB helper coverage puede usarse como **candidato de adopción futura**, pero no como cobertura integrada ni como `Verified-runtime` de `staging`. Tampoco se eleva a `Verified-static` global porque la fuente candidata difiere materialmente del código de `staging`; el estado correcto es `Divergent` hasta que una slice posterior adopte código, importe, compile y ejecute pruebas focales.

## 7. `tdd_uat_matrix` — WU3

Esta matriz no autoriza UAT HTML todavía. Cada caso queda bloqueado hasta tener un átomo verde ejecutado en el `staging` actual, un manifest válido y un `ref` firmado en el registro de aceptación.

| Regla / contrato | Átomo TDD requerido | Manifest | Escenario UAT | `ref` | Estado | Clase |
|---|---|---|---|---|---|---|
| PCSUB: guardar Datos Generales debe persistir el bloque y solo después permitir avance de workflow por servicio. | `Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow` | `tests/tests.pcsub.form.json` | DADO una PCSUB en preregistro, CUANDO se guardan Datos Generales completos desde el formulario, ENTONCES se persisten los datos y se ofrece el avance permitido. | Bloqueado: falta átomo verde en `staging` y registro firmado. | `Divergent`: helper/cobertura PCSUB existe solo como candidata no integrada. | happy |
| PCSUB: el formulario no debe tratar un bloque incompleto como listo para enviar a Calidad. | `Test_PCSUB_GuardarDesdeSubform_IncompleteTechnicalData_NavigatesToPendingBlock` | `tests/tests.pcsub.form.json` | DADO una PCSUB en Desarrollo Técnico con Detalle o Motivos incompletos, CUANDO se guarda, ENTONCES se mantiene al usuario en el bloque pendiente sin transición indebida. | Bloqueado: falta átomo verde en `staging` y registro firmado. | `Divergent`: requiere adoptar helper/seam antes de probar. | sad |
| PCSUB: cierre de formalización exige PDF de cierre antes de aprobar definitivamente. | `Test_PCSUB_CierreFormalizacion_RequiresSignedPdf` | `tests/tests.pcsub.form.json` | DADO una PCSUB en formalización sin PDF firmado, CUANDO se intenta cerrar, ENTONCES el sistema bloquea el cierre y explica la acción pendiente. | Bloqueado: falta átomo verde en `staging` y registro firmado. | `Divergent`: los átomos candidatos de cierre no están integrados; la rama candidata declara casos bloqueados por sandbox. | edge |
| PC: `GuardarDesdeSubform` debe quedar detrás de helper testeable antes de UAT de navegación/transacción. | `Test_PC_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow` | `tests/tests.pc.form.json` | DADO una PC en preregistro, CUANDO se guarda Datos Generales, ENTONCES se persiste y la transición se canaliza por `WorkflowServicio`. | Bloqueado: falta átomo verde y `ref` firmado. | `Verified-static`: deuda de seam documentada; sin manifest PC strict. | happy |
| CDCA: harness strict debe migrar antes de usar smoke como evidencia de aceptación. | `Test_CDCA_GuardarDatosGenerales_UsesInjectedDbAndCardinality` | `tests/tests.cdca.strict.json` | DADO una CD/CA con fixture controlado, CUANDO se guarda Datos Generales, ENTONCES la escritura ocurre en la `DAO.Database` inyectada y se verifica cardinalidad. | Bloqueado: falta átomo migrado a v2.4.2 y `ref` firmado. | `Verified-static` con smoke diagnóstico; harness actual `Divergent` frente al contrato strict. | adversarial |
| CDCASUB: los átomos parciales no cubren todavía guardados ni transiciones completas. | `Test_CDCASUB_GuardarDatosGenerales_PromotesWithInjectedDb` | `tests/tests.cdcasub.json` | DADO una CDCASUB en preregistro, CUANDO se guarda Datos Generales completos, ENTONCES se persiste el bloque y se verifica el avance con `db` explícito. | Bloqueado: falta átomo focal y `ref` firmado. | `Verified-runtime` parcial para 3 átomos; `Verified-static` para guardados/transiciones. | happy |
| Workflow: una transición debe registrar `tbLogEstados` y rechazar precondiciones fallidas sin SQL directo. | `Test_Workflow_EjecutarTransicion_PreconditionFailureDoesNotMutateState` | `tests/tests.workflow.json` | DADO una solicitud que no cumple precondiciones, CUANDO se intenta avanzar de estado, ENTONCES no cambia `idEstadoInterno` y queda un error de validación comprensible. | Bloqueado: falta átomo sobre `EjecutarTransicion` y `ref` firmado. | `Verified-runtime` parcial para repositorio de log; `Verified-static` para motor de transición. | sad |
| Transversal: subformularios Tier 3 (`Propuesta`/`Impacto`) son adaptadores finos, no cierre de cobertura. | `Test_FormTier3_SubformsDelegateOnly_NoBusinessRule` | `tests/tests.form-tiering.json` | DADO un subformulario Tier 3, CUANDO el usuario pulsa guardar, ENTONCES solo delega al padre y no ejecuta reglas propias ocultas. | Bloqueado: falta átomo o revisión estática firmada por caso. | `Verified-static`: deuda aceptada solo si el padre tiene helper y átomo verde. | edge |

## 8. Bloqueo de UAT HTML — WU3

No se debe generar ni entregar ningún HTML UAT para estas reglas mientras la fila correspondiente de `tdd_uat_matrix` no tenga:

1. átomo TDD verde ejecutado contra el `staging` actual;
2. manifest Dysflow focal con procedimiento público global único;
3. `ref` firmado que enlace átomo, commit de implementación y caso UAT;
4. estado del ledger actualizado en el documento de capacidad afectado.

Hasta cumplir esos cuatro puntos, cualquier web UAT queda en estado **bloqueado**, aunque el caso sea claro desde producto.

## 9. Contratos para extracciones futuras — WU4

Phase 0 es una auditoría documental y contractual. No extrae ni adopta código de `src/forms/*.cls`, `src/modules/*.bas` ni `src/classes/*.cls`. Cualquier extracción posterior debe abrir una slice propia con RED primero, importación controlada y verificación Dysflow antes de reclamar cobertura runtime.

### 9.1 Gate RED antes de producción

Antes de tocar producción para una extracción futura, debe existir una tarea RED explícita y visible con estos mínimos:

| Elemento | Contrato obligatorio |
|---|---|
| Módulo de test | Crear o actualizar `src/modules/Test_*.bas` con un átomo focal que falle contra el comportamiento actual o contra la ausencia del helper. |
| Manifest focal | Crear o actualizar `tests/*.json` con un procedimiento público global único, sin calificar por módulo y sin agregadores `RunAll` como evidencia atómica. |
| Escenario | El átomo debe mapear una fila de `tdd_uat_matrix` o una regla de capacidad con clase happy/sad/edge/adversarial. |
| Fixture | Si toca datos, usar fixture propio, schema-first, `DAO.Database` explícita y cardinalidad antes/después para mutaciones. |
| Bloqueo | Si el RED requiere implementar primero, la slice debe detenerse: no se toca producción sin contrato de test previo. |

### 9.2 Gate para cambios en módulos y clases

Para cambios que solo afecten `src/modules/*.bas` o `src/classes/*.cls`, el cierre mínimo de la slice posterior es:

1. editar fuente en `src/` sin tocar formularios;
2. ejecutar `import_modules` o `import_all` con el conjunto mínimo necesario;
3. ejecutar compilación Dysflow (`compile_vba` o importación con `compile: true`) y resolver cualquier `VBA_COMPILE_ERROR` antes de continuar;
4. ejecutar `test_vba` contra el manifest focal `tests/*.json` definido en el RED;
5. ejecutar `verify_code` para confirmar que fuente y binario quedan sincronizados;
6. actualizar el ledger de capacidad afectado (`Verified-runtime`, `Verified-static` o `Divergent`) con evidencia, fecha, manifest y `ref`.

Si cualquiera de estos pasos falla, no se puede elevar cobertura a `Verified-runtime` ni desbloquear el caso UAT relacionado.

### 9.3 Gate para cambios en formularios y reportes

Para cambios que afecten `src/forms/*.cls`, `src/forms/*.form.txt` o `src/reports/*`, el cierre mínimo añade la limitación de compilación manual de Access:

1. editar código de formulario solo en `.cls` y layout solo en `.form.txt` / `.report.txt`;
2. importar el formulario o reporte con Dysflow;
3. pedir al usuario compilación manual en VBE (`Debug -> Compile`) y esperar confirmación explícita de OK;
4. no ejecutar `test_vba` dependiente de ese formulario antes de la confirmación manual;
5. tras el OK, ejecutar `test_vba` contra el manifest focal;
6. ejecutar `verify_code` para comprobar sincronía fuente↔binario;
7. actualizar el ledger de capacidad y mantener el caso UAT bloqueado si falta átomo verde, manifest focal o `ref` firmado.

La compilación headless puede verificar módulos estándar y clases, pero no sustituye la confirmación manual para document modules de formularios/reportes.

### 9.4 Separación de Phase 0 frente a futuras extracciones

Este documento no autoriza adopción automática de helpers candidatos ni refactor de formularios. En particular:

- `src/modules/DatosPCSUBGuardarHelper.bas` sigue siendo un solape local no versionado y debe reconciliarse antes de cualquier adopción.
- La cobertura PCSUB de la rama `feature/pcsub-workflow-helper-coverage` permanece `Divergent` para el `staging` actual hasta que una slice futura adopte código, importe, compile y ejecute pruebas focales.
- Los formularios Tier 1 y Tier 2 identificados aquí son cola de extracción futura; Phase 0 solo define el contrato de entrada y salida de esas slices.
- Los subformularios Tier 3 siguen como deuda documentada y no equivalen a cierre de cobertura.

## 10. Riesgos y deuda para los siguientes slices

| Riesgo | Impacto | Siguiente acción |
|---|---|---|
| PCSUB helper candidate no integrado | Puede inducir a adoptar cobertura inexistente en `staging`. | Tratarlo como `Divergent`; una slice posterior debe reconciliar `src/modules/DatosPCSUBGuardarHelper.bas` local antes de adoptar código. |
| Prompts dentro de formularios padre | Bloquea tests headless si se extrae sin seam. | Diseñar resultado de helper con `promptKind`, `message` y `suggestedAction`. |
| Transacciones mezcladas formulario/servicio | Riesgo de doble transacción o rollback incompleto. | Cada helper futuro debe decidir si recibe `db` externa o gestiona transacción, nunca ambas sin contrato. |
| Subformularios Tier 3 | Deuda aceptada solo por este inventario; no es cierre de cobertura. | Registrar en ledger de capacidades en WU3. |
| UAT HTML prematuro | Calidad podría validar flujos que no tienen átomo verde ni `ref` firmado. | Mantener bloqueo explícito hasta que cada fila de `tdd_uat_matrix` alcance evidencia runtime. |
| Extracción sin RED previo | Una slice futura podría implementar helpers sin contrato ejecutable y producir falso verde. | Exigir `src/modules/Test_*.bas` + manifest focal `tests/*.json` antes de tocar producción. |
| Cambio de formulario sin compile manual | Access no garantiza document modules compilados con verificación headless. | Importar, pedir compilación manual en VBE, esperar OK y solo después ejecutar `test_vba`. |

## 11. Verificación de este slice

- Phase 0 no incluye modificaciones de código en `src/`; el workspace ya contiene `src/modules/DatosPCSUBGuardarHelper.bas` como archivo no versionado y queda fuera de alcance hasta una slice de reconciliación/adopción.
- Phase 0 no incluye cambios en `CONDOR.accdb`; el workspace muestra el binario modificado fuera del alcance de esta auditoría, por lo que no se usa como evidencia de Phase 0 ni se declara verificado.
- No se ejecutaron escrituras Dysflow ni `import_modules`.
- Validación WU1: revisión estática focalizada y actualización de tareas SDD 1.1-1.3.
- Validación WU2: comandos Git de solo lectura para ancestros, diff, `--stat`, `--numstat` y `--check`; clasificación por archivo documentada; PCSUB coverage marcado como `Divergent` para `staging`.
- Validación WU3: matriz TDD↔UAT documentada; capability ledger actualizado para CAP-001, CAP-003, CAP-004, CAP-005 y CAP-007; HTML UAT bloqueado hasta átomo verde y `ref` firmado.
- Validación WU4: contratos RED/import/compile/test/verify documentados; separación explícita entre Phase 0 documental y futuras extracciones de `src/forms/*.cls`, `src/modules/*.bas` y `src/classes/*.cls`.
