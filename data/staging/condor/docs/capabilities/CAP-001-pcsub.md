# Capacidad: Gestión de Propuestas de Cambio de Subcontratista (PCSUB)

## §0 Identidad

- **ID de capacidad**: CAP-001
- **Nivel**: crítico
- **Estado**: activa con verificación runtime histórica para reglas BR-001..BR-011 de servicio/repositorio; cobertura helper/formulario Phase 0 en estado `Divergent` para el `staging` actual.
- **Fuente**: híbrida
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `tests/tests.pcsub.json` y chunks `tests/sequences/pcsub-br*.json`: 29 pruebas atómicas BR-001..BR-011 superadas mediante `dysflow.test_vba`; `dysflow.compile_vba` OK; `dysflow.verify_binary(Test_PCSUB_Strict)` con `actionableOk=true` y solo diferencia no funcional `caseOnly`.
- **Confianza global**: `Verified-runtime` histórico para las reglas de servicio/repositorio BR-001..BR-011 cubiertas por `Test_PCSUB_Strict`. Las afirmaciones dependientes de helpers/formularios Phase 0 son `Divergent` frente al `staging` actual porque la rama candidata `feature/pcsub-workflow-helper-coverage` no está integrada; la navegación UI sigue como deuda hasta extraer seams y ejecutar verificación específica.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). PCSUB ya tiene manifest atómico estricto; la deuda restante se concentra en reglas UI/formulario y decisiones de producto sobre gaps de completitud/RAC delegado.

**Justificación del nivel**: crítico, porque PCSUB es un tipo de solicitud completo con interfaz propia, persistencia en `tbDatosPCSUB`, reglas de workflow, formalización documental y cobertura de pruebas fixture-first. Una regresión puede bloquear la gestión de solicitudes de subcontratista.

**Contrato TDD vigente**: cualquier prueba nueva o migrada de PCSUB debe alinearse con `access-vba-tdd` v2.4.2: función pública que devuelve JSON canónico, fixture propio en sandbox, inspección schema-first, `DAO.Database` inyectado explícitamente, cardinalidad de mutaciones, manifest Dysflow con procedimiento global único sin calificación, sin UI/`Debug.Print` y sin mutar `TbConfiguracionBackends`. Las reglas hoy acopladas al formulario deben moverse a helper, servicio o ViewModel antes de tratarlas como pruebas serias de comportamiento; el formulario debe conservar solo orquestación visual.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: permitir gestionar una Propuesta de Cambio de Subcontratista como gemelo funcional de PC, adaptada al rol y a los datos del subcontratista.
- **Usuarios / perfiles**: subcontratistas, técnicos, calidad y perfiles de aprobación/formalización involucrados en el ciclo de vida de solicitudes.
- **Problema que resuelve**: separar las propuestas de cambio iniciadas o tratadas por subcontratistas de las PC principales, conservando workflow, validaciones y documentación propias.
- **Valor de negocio**: trazabilidad de cambios técnicos de subcontratista, control de calidad y formalización documental dentro del flujo CONDOR.
- **No objetivos**: este documento no cubre en detalle todas las capacidades transversales de alta de solicitud, generación Word/PDF, adjuntos, seguridad o workflow global; las vincula como dependencias.
- **Origen de la intención**: SDD `openspec/changes/pcsub-nueva-solicitud`, `pcsub-staging-recovery`, `pcsub-guardar-phase-advancement`, `pcsub-full-coverage-dpcdcf`, más lectura del código actual.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** una solicitud PCSUB válida **CUANDO** se abre `frmDatosPCSUB` con un `idSolicitud` válido **ENTONCES** el formulario carga un `DatosPCSUBViewModel`, muestra la pestaña activa según `WorkflowServicio.getPaginaActivaPCSUB` y puebla el subformulario actual. **Estado**: `Verified-static`; requiere reconciliar deriva de formularios antes de elevarlo a `Verified-runtime`.
- **DADO** una PCSUB en captura de datos generales **CUANDO** el usuario guarda campos obligatorios completos **ENTONCES** se persisten en `tbDatosPCSUB`; si aplica, el flujo puede avanzar de preregistro a registro y luego preguntar por asignación a Desarrollo Técnico. **Estado**: reglas de servicio/repositorio `Verified-runtime` por `Test_PCSUB_Strict`; el alcance Slice 2 `Bloque_Generales` PCSUB queda `Verified-runtime` por `Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow`; el resto de bloques UI sigue condicionado a seams/verificación específica.
- **DADO** una PCSUB en Desarrollo Técnico **CUANDO** se guardan propuesta e impacto **ENTONCES** el servicio valida y persiste los bloques correspondientes. **Estado**: `Verified-runtime` por pruebas atómicas BR-002..BR-006.
- **DADO** una PCSUB con datos técnicos de impacto **CUANDO** no se declara ningún motivo **ENTONCES** el servicio rechaza el guardado con error de validación y no debe mutar el estado persistido cubierto por pruebas. **Estado**: `Verified-runtime` por `Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo`.
- **DADO** una PCSUB en Dictamen RAC **CUANDO** la decisión no es `RECHAZADO` **ENTONCES** `racCodigo` es obligatorio según el código actual. **Estado**: `Verified-runtime` por BR-008; mantiene divergencia de intención sobre RAC delegado.
- **DADO** una PCSUB en cierre de formalización **CUANDO** se aprueba definitivamente **ENTONCES** el formulario exige seleccionar un PDF de cierre y delega en `WorkflowServicio.EjecutarCierreFormalizacion`. **Estado**: `Verified-static`; requiere prueba específica y reconciliación de formularios.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | Los datos generales requieren referencia de contrato/inspección, referencia de sub-suministrador, nombre/dirección de sub-suministrador, denominación de contrato y objeto de contrato. | Código + intención PCSUB | Sí: `DatosPCSUBServicio.GuardarDatosGenerales` y `EsDatosGeneralesCompleta`. | `Test_PCSUB_BR001_*` en `tests/tests.pcsub.json` / `pcsub-br001-br002.json`. | Verified-runtime |
| BR-002 | La propuesta requiere descripción del material afectado y descripción de propuesta de cambio. | Código | Sí: `GuardarPropuesta`. | `Test_PCSUB_BR002_*` en `tests/tests.pcsub.json` / `pcsub-br001-br002.json`. | Verified-runtime |
| BR-003 | Impacto requiere al menos un motivo; si `Otros` está marcado, requiere detalle; si hay detalle de otros, debe estar marcado `Otros`. | Código | Sí: `GuardarImpacto`. | `Test_PCSUB_BR003_*` en `tests/tests.pcsub.json` / `pcsub-br003-br004.json`. | Verified-runtime |
| BR-004 | Incidencia de coste y plazo solo aceptan `AUMENTARÁ`, `DISMINUIRÁ` o `NO VARIARÁ`. | Código | Sí: `GuardarImpacto`. | `Test_PCSUB_BR004_*` en `tests/tests.pcsub.json` / `pcsub-br003-br004.json`. | Verified-runtime |
| BR-005 | Clasificación de impacto solo acepta `MAYOR` o `MENOR`. | Código | Sí: `GuardarImpacto`; completitud técnica comprueba clasificación. | `Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida` en `pcsub-br005-br006.json`. | Verified-runtime |
| BR-006 | El cambio debe indicar si afecta a `Material ya entregado` o `Material por entregar`. | Código | Sí: `GuardarImpacto`. | `Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido` en `pcsub-br005-br006.json`. | Verified-runtime |
| BR-007 | Aprobación de sub-suministrador requiere nombre de firma de Oficina Técnica, nombre de firma del representante del sub-suministrador y `decisionFinal` no vacío. | Código | Sí: `GuardarAprobacionSuministrador` y `EsAprobacionSuministradorCompleta` (completitud valida las 2 firmas + `decisionFinal`). | `Test_PCSUB_BR007_*` en `tests/tests.pcsub.json` / `pcsub-br007-br008.json`; `Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse` en `tests.vba.json`. | Verified-runtime |
| BR-008 | En Dictamen RAC, si la decisión no es `RECHAZADO`, el código RAC es obligatorio. Completitud además exige `racDecision` no vacío. | Código | Sí: `GuardarDictamenRAC`; `EsDictamenRACCompleto` valida `racCodigo` + `racNombre` + `racDecision`. | `Test_PCSUB_BR008_*` en `tests/tests.pcsub.json` / `pcsub-br007-br008.json`; `Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse` en `tests.vba.json`. | Verified-runtime; mantiene divergencia RAC delegado |
| BR-009 | La decisión final requiere `decisionFinal` y `NombreFirmanteFinal` no vacíos. | Código | Sí: `GuardarDecisionFinal`; `EsDecisionFinalCompleta` valida ambos. | `Test_PCSUB_BR009_*` en `tests/tests.pcsub.json` / `pcsub-br009-br010.json`; `Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse` en `tests.vba.json`. | Verified-runtime |
| BR-010 | Las operaciones de limpieza deben borrar solo el bloque objetivo y conservar el resto de la fila/parentales. | Código + tests | Sí en repositorio para RAC, aprobación, decisión y eliminación PCSUB. | `Test_PCSUB_BR010_*` en `tests/tests.pcsub.json` / `pcsub-br009-br010.json`. | Verified-runtime |
| BR-011 | La conexión `DAO.Database` inyectada debe conservarse y seguir siendo el contexto usado por servicio/repositorio. Las transiciones UI siguen debiendo pasar por `WorkflowServicio`. | Regla de arquitectura + código | Sí para preservación de conexión en `GuardarDictamenRAC` y `ActualizarAprobacionSuministrador`; workflow UI sigue como deuda. | `Test_PCSUB_BR011_*` en `tests/tests.pcsub.json` / `pcsub-br011.json`; reglas UI sin prueba focal. | Verified-runtime para inyección; Verified-static/deuda para UI |

La deuda de BR-011 debe cerrarse con pruebas compatibles con `access-vba-tdd` v2.4.2: extraer la decisión de navegación/transición a helper, servicio o ViewModel, sembrar fixture propio en sandbox, inyectar `DAO.Database` y verificar efectos observables sin depender de `Debug.Print`, `MsgBox` ni manipulación directa de controles como contrato principal.

### Validaciones observadas

- Datos generales: campos obligatorios de referencia, sub-suministrador, contrato y objeto.
- Propuesta: material afectado y descripción de propuesta obligatorios.
- Impacto: motivos, detalle de otros, incidencias de coste/plazo, clasificación y afectación a material.
- Aprobación suministrador: dos nombres de firma obligatorios + `decisionFinal`.
- RAC: código obligatorio salvo decisión `RECHAZADO` + `racDecision`.
- Decisión final: decisión obligatoria + firmante obligatorio.
- Apertura del formulario: `OpenArgs` debe contener `idSolicitud` válido. Estado documental condicionado por deriva de formulario.
- Cierre formalización: selección obligatoria de PDF firmado antes de ejecutar cierre. Estado documental condicionado por deriva de formulario y falta de prueba focal.

### Transiciones de estado y navegación

- `Bloque_Generales` puede activar avance desde registro hacia Desarrollo Técnico mediante `WorkflowServicio.EjecutarTransicion` si el usuario acepta la asignación. Estado: `Verified-runtime` solo para el camino PCSUB Slice 2 (`estadoRegistro -> estadoDesarrolloTecnico`) probado por `Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow`; el resto de navegación/transiciones UI sigue `Verified-static`.
- En `estadoDesarrolloTecnico`, si faltan motivos se navega a `tabImpacto`. Estado: `Verified-static` hasta reconciliar `Form_frmDatosPCSUB.cls`.
- En `estadoDesarrolloTecnico`, si se guarda Impacto pero falta Detalle, se navega a `tabPropuesta`. Estado: `Verified-static` hasta reconciliar `Form_frmDatosPCSUB.cls`.
- Cuando Detalle y Motivos están completos, el formulario pregunta si se envía a Calidad y transiciona a `estadoModificacion` si el usuario acepta. Estado: `Verified-static` hasta reconciliar `Form_frmDatosPCSUB.cls`.
- En formalización, `ProcesarCierreFormalizacion` delega el cierre y archivo del PDF en `WorkflowServicio.EjecutarCierreFormalizacion`. Estado: `Verified-static` hasta reconciliar `Form_frmDatosPCSUB.cls`.

### Casos límite y hallazgos

- La spec inicial pretendía campos `racDelegado*`; el código actual usa `racCodigo`, `racDecision`, `racNombre`, `observacionesRACDelegador` y `racNombreDelegador`. Esto queda como divergencia para revisión.
- La cobertura histórica caracterizaba gaps de completitud: por ejemplo, funciones que podían devolver completo aunque `racDecision`, `decisionFinal` o `NombreFirmanteFinal` estuvieran en blanco. **Resuelto 2026-06-18**: las 3 funciones `Es*Completa` ahora validan el tercer campo correspondiente. Atomos `Test_PCSUB_Es*Blank*_ReturnsFalse` en `tests.vba.json`.
- No hubo fallos de prueba en la evidencia indicada: 29 pruebas atómicas BR-001..BR-011 pasaron mediante Dysflow. Esto no cierra las reglas UI/formulario ni las decisiones de producto abiertas sobre RAC delegado y completitud.

### Señales de aceptación / presencia

- Existe `Form_frmDatosPCSUB.cls` y los subformularios PCSUB reales descubiertos: `AprobacionSuministrador`, `DecisionFinal`, `DictamenRAC`, `Generales`, `Impacto` y `Propuesta`.
- Existe la capa MVVM/servicio/repositorio: `DatosPCSUB.cls`, `DatosPCSUBViewModel.cls`, `DatosPCSUBServicio.cls`, `DatosPCSUBRepositorio.bas`.
- La persistencia usa `tbDatosPCSUB` con `idSolicitud` como clave de dominio y `idDatosPCSUB` como identificador de fila.
- Los tests PCSUB fixture-first usan el grafo `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` con IDs deterministas `>= 900000`.
- Las pruebas futuras deben mantener el fixture propio y revisar el esquema real antes de insertar; cualquier mutación debe demostrar cardinalidad antes/después.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmDatosPCSUB.Form_Load` carga ViewModel, determina pestaña activa y navega. Evidencia condicionada por deriva `bothChanged` en `Form_frmDatosPCSUB.cls`.
  - `Form_frmDatosPCSUB.GuardarDesdeSubform` recoge valores desde el subformulario activo, exporta VM→Entidad y llama al servicio por bloque. Evidencia condicionada por deriva `bothChanged`.
  - `Form_frmDatosPCSUB.ProcesarCierreFormalizacion` exige PDF y delega el cierre. Evidencia condicionada por deriva `bothChanged`.
- **Subformularios PCSUB reales**:
  - `Form_subfrmDatosPCSUB_AprobacionSuministrador`
  - `Form_subfrmDatosPCSUB_DecisionFinal` — deriva `sourceNewer`; requiere importar a binario o reconciliar explícitamente.
  - `Form_subfrmDatosPCSUB_DictamenRAC`
  - `Form_subfrmDatosPCSUB_Generales`
  - `Form_subfrmDatosPCSUB_Impacto`
  - `Form_subfrmDatosPCSUB_Propuesta`
- **Puntos de entrada de código**:
  - `DatosPCSUBServicio.GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC`, `GuardarDecisionFinal`.
  - `DatosPCSUBServicio.EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta`.
  - `DatosPCSUBRepositorio.Guardar`, `getPorIdSolicitud`, `Actualizar*`, `Limpiar*`, `EliminarPorIdSolicitud`.
- **Datos afectados**:
  - `TbExpedientes`: padre de dominio en fixtures y contexto de expediente.
  - `tbSolicitudes`: solicitud y estado de workflow.
  - `tbDatosPCSUB`: datos específicos de PCSUB.
- **Salidas**: persistencia de bloques PCSUB, mensajes de validación/navegación, transición de workflow y archivo de PDF de cierre mediante servicio transversal.
- **Dependencias**: `WorkflowServicio`, `SolicitudServicio`, `ExpedienteServicio`, `FormulariosPadreAuxiliares`, `RepositorioComun`, `RechazoRepositorio`, `JsonHelper`, `TestHelper`.
- **Valoración de diseño**: la estructura principal respeta Formulario → ViewModel → Servicio → Repositorio y usa `getdb()` genérico. Hay deuda documentada en comportamiento de completitud, divergencia histórica de RAC delegado y reconciliación fuente↔binario de formularios.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar/validar las clases y módulos PCSUB: entidad, ViewModel, servicio, repositorio y formulario principal.
2. Restaurar/validar subformularios PCSUB por bloque: Generales, Propuesta, Impacto, Aprobación Suministrador, Dictamen RAC y Decisión Final.
3. Confirmar que `FormulariosPadreAuxiliares` mapea `TipoForm_PCSUB` a los subformularios y bloques correctos.
4. Confirmar que `WorkflowServicio` expone las páginas/guardas de PCSUB y que las transiciones se ejecutan por servicio.
5. Confirmar esquema y vínculos del backend seleccionado: `TbExpedientes`, `tbSolicitudes`, `tbDatosPCSUB`.
6. Reconciliar deriva fuente↔binario:
   - `Form_frmDatosPCSUB.cls`: `bothChanged`; requiere merge manual antes de importar o exportar.
   - `Form_subfrmDatosPCSUB_DecisionFinal.cls`: `sourceNewer`; recomendación: importar fuente al binario después de revisión.
7. Importar fuentes al binario con `dysflow.import_modules` / `dysflow.import_all` solo cuando la reconciliación esté decidida.
8. Compilar con `dysflow.compile_vba`.
9. Verificar fuente↔binario con `dysflow.verify_binary` hasta no tener deriva accionable.
10. Ejecutar pruebas PCSUB con `dysflow.test_vba` usando los chunks `tests/sequences/pcsub-br*.json` o el manifest atómico `tests/tests.pcsub.json` si el timeout disponible lo permite.

## §5 Evidencia y trazabilidad

- **Evidencia Dysflow incorporada**:
  - `tests/tests.pcsub.json` + `tests/sequences/pcsub-br*.json`: 29 pruebas atómicas BR-001..BR-011 superadas.
  - `dysflow.compile_vba`: OK tras importar `Test_PCSUB_Strict`.
  - `dysflow.verify_binary(Test_PCSUB_Strict)`: `actionableOk=true`; solo diferencia `caseOnly` no funcional.
- **Lecturas documentales y de código referenciadas**:
  - `src/forms/Form_frmDatosPCSUB.cls`
  - `src/classes/DatosPCSUBServicio.cls`
  - `src/modules/DatosPCSUBRepositorio.bas`
  - `tests/tests.vba.json`
  - `tests/tests.pcsub.json`
  - `docs/testing/pcsub-fixture-graph.md`
  - `docs/testing/pcsub-coverage-map.md`
- **SDD/intención consultada**:
  - `openspec/changes/pcsub-nueva-solicitud/spec/pcsub.md`
  - `openspec/changes/pcsub-staging-recovery/proposal.md`
  - `openspec/changes/pcsub-guardar-phase-advancement/proposal.md`
  - `changes/pcsub-full-coverage-dpcdcf/design.md`
- **Estado honesto de evidencia**: no hay fallos de prueba PCSUB en la evidencia recibida, pero no se declara cierre completo de regresión mientras persista deriva accionable fuente↔binario.
- **Política de actualización de pruebas**: la evidencia histórica se conserva como contexto, pero las pruebas nuevas o modificadas deben cumplir `access-vba-tdd` v2.4.2 y registrar en el manifest un procedimiento público global único sin calificación.
- **Informe de deuda v2.4.2**: [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md).

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| PCSUB BR-001..BR-011 servicio/repositorio | Pendiente | `f27f693` en `staging` | pending | Pendiente | Pendiente | 29 atómicos strict TDD v2.4.2 en verde; falta release/UAT y reglas UI/formulario. |

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| El formulario PCSUB no abre o abre pestaña incorrecta | formulario no importado, ViewModel incompleto, página activa de workflow incorrecta o deriva `Form_frmDatosPCSUB.cls` | `dysflow.verify_binary`; prueba focal de `Form_Load`/navegación | §2 / §3 / §5 |
| Guardar bloque no persiste | repositorio no actualizado, backend vinculado incorrecto o validación falla antes de persistir | `dysflow.test_vba` con `tests/tests.pcsub.json`; consulta `tbDatosPCSUB` en sandbox | §2 / §3 |
| Flujo técnico no promueve a Calidad | lógica de `GuardarDesdeSubform` o `WorkflowServicio` divergente | prueba de workflow/UI enfocada después de reconciliar formularios | §2 / §5 |
| Tests pasan por datos existentes | fixture no determinista | revisar `docs/testing/pcsub-fixture-graph.md` y rehacer fixture-first | §5 |

## §6 Notas de migración web

- **Conservar**: reglas de validación por bloque, transiciones por servicio de workflow, separación entre datos generales/propuesta/impacto/aprobación/RAC/decisión final, grafo de datos `Expediente -> Solicitud -> DatosPCSUB`.
- **Transformar**: navegación por subformularios de Access a rutas/pasos web; mensajes modales a validaciones y acciones explícitas; `SendKeys`/`BrowseTo` a navegación controlada.
- **NO copiar**: dependencia de estado global de Access, controles `NavigationSubform`, `MsgBox` como contrato de negocio único, y cualquier test que dependa de datos preexistentes.
- **Preguntas abiertas**: confirmar con producto el contrato real de RAC delegado (los gaps de completitud de `Es*Completa` en los 4 gemelos PCSUB/PC/CDCA/CDCASUB quedaron resueltos 2026-06-18 con el patrón del commit `5453f25` aplicado en los 4 servicios).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| PCSUB tiene formulario principal y subformularios por bloque: `AprobacionSuministrador`, `DecisionFinal`, `DictamenRAC`, `Generales`, `Impacto` y `Propuesta`. | Verified-static | Descubrimiento de subformularios PCSUB y referencias documentales. | 2026-06-15 |
| PCSUB persiste en `tbDatosPCSUB` por `idSolicitud`. | Verified-runtime para rutas ejercitadas por la suite; Verified-static para el contrato estructural completo. | `tests/tests.pcsub.json`, `tests/sequences/pcsub-br*.json`, `DatosPCSUBRepositorio.bas`. | 2026-06-15 |
| Datos generales, propuesta, impacto, aprobación, RAC y decisión final tienen métodos de guardado en servicio. | Verified-runtime para rutas ejercitadas por la suite; Verified-static para UI/formularios. | `Test_PCSUB_Strict.bas`, 29 pruebas atómicas BR-001..BR-011. | 2026-06-15 |
| Las pruebas PCSUB estrictas están en verde en el binario actual. | Verified-runtime | `compile_vba` OK; chunks `pcsub-br001-br002`, `pcsub-br003-br004`, `pcsub-br005-br006`, `pcsub-br007-br008`, `pcsub-br009-br010`, `pcsub-br011` OK; commit `f27f693`. | 2026-06-15 |
| La fuente y el binario Access están reconciliados para `Test_PCSUB_Strict`. | Verified-runtime | `verify_binary(Test_PCSUB_Strict)`: `actionableOk=true`, `caseOnly` no funcional. | 2026-06-15 |
| Phase 0 WU2 marca la cobertura helper/formulario PCSUB como no integrada en el `staging` actual. | Divergent | `audit-e2e-thin-forms-phase-0.md` §6.1-§6.3: ningún commit `b946c15^..59b4438` de `feature/pcsub-workflow-helper-coverage` es ancestro de `staging`; diff de 30 archivos con formularios, helpers, tests y manifest. | 2026-06-26 |
| Phase 0 WU3 bloquea UAT HTML de PCSUB hasta tener átomo verde, manifest focal y `ref` firmado por caso. | Divergent / pendiente | `audit-e2e-thin-forms-phase-0.md` §7-§8; matriz TDD↔UAT con casos PCSUB de guardado, navegación incompleta y cierre de formalización. | 2026-06-26 |
| Los subformularios PCSUB `Propuesta` e `Impacto` quedan como Tier 3: adaptadores finos documentados, no cierre de cobertura. | Verified-static / deuda Tier 3 | `audit-e2e-thin-forms-phase-0.md` §2 y §7; dependen de extraer `Form_frmDatosPCSUB.GuardarDesdeSubform` a helper testeable. | 2026-06-26 |
| PCSUB `Bloque_Generales` guardado desde subformulario persiste los cinco campos generales y solo planifica/ejecuta la transición `estadoRegistro -> estadoDesarrolloTecnico` vía helper y `WorkflowServicio` después de validación/persistencia correcta. | Verified-runtime limitado a Slice 2 | `tests/tests.pcsub.form.json` → 2/2 verde tras compile manual VBE de `Form_frmDatosPCSUB`: `Test_PCSUB_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow` y `Test_PCSUB_GuardarDesdeSubform_Generales_InvalidDoesNotPlanPromptOrWorkflow`; `verify_code` focal `actionableOk=true`, `recommendedAction=no_action`. | 2026-06-27 |
| El contrato puro compartido de Decisión Final exige `decisionFinal` y `NombreFirmanteFinal` no vacíos, con trim y soporte de `Null`. | Verified-runtime limitado a helper puro Slice 3.1 | `tests/tests.decision-final.json` → 4/4 verde: matriz PC/PCSUB/CDCA/CDCASUB, decisión ausente, firmante ausente y espacios; `DecisionFinalHelper_EsCompleta` importado y compilado con Dysflow. | 2026-06-27 |
| SDD inicial pretendía RAC delegado completo. | Intended | `openspec/changes/pcsub-nueva-solicitud/spec/pcsub.md`. | 2026-06-15 |
| El código actual de RAC usa campos RAC estándar y campos delegador parciales, no el contrato completo `racDelegado*` de la spec inicial. | Divergent | `DatosPCSUBServicio.GuardarDictamenRAC`; `DatosPCSUBRepositorio.ActualizarDictamenRAC`; spec PCSUB. | 2026-06-15 |
| Algunas funciones de completitud tienen gaps caracterizados. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con átomos `Test_PCSUB_Es*Blank*_ReturnsFalse` en `tests.vba.json` que ahora assertan `False` (en lugar de `True` que era el comportamiento de gap previo). | 2026-06-15 (original) → 2026-06-18 (resolución) |

**Divergencias pendientes de revisión humana**:

- RAC delegado: confirmar si debe implementarse el contrato SDD original o si el comportamiento actual con campos RAC estándar es el nuevo contrato aceptado.
- Completitud de RAC/aprobación/decisión final: decidir si los gaps caracterizados son comportamiento aceptado o bugs a corregir.- Deriva fuente↔binario PCSUB: reconciliar `Form_frmDatosPCSUB.cls` y `Form_subfrmDatosPCSUB_DecisionFinal.cls` antes de declarar cierre completo de regresión.
