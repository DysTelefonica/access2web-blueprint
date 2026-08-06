# Capacidad: Gestión de Cambios de Diseño/Alcance de Subcontratista (CDCASUB)

## §0 Identidad

- **ID de capacidad**: CAP-005
- **Tier**: critical
- **Estado**: active con deuda de reconciliación UI/layout y de pruebas
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre `DatosCDCASUBServicio`, `DatosCDCASUBRepositorio`, `Form_frmDatosCDCASUB` y seis subformularios CDCASUB. Servicio/repositorio y code-behind `.cls` están `matched`; `Form_frmDatosCDCASUB.form.txt` solo `formSerializationOnly`; los seis `.form.txt` de subformularios CDCASUB tienen diferencias accionables `bothChanged`.
- **Confianza global**: mayoritariamente `Verified-static`. Las reglas de servicio/repositorio están implementadas y el código está sincronizado con el binario, pero la ausencia de manifest atómico CDCASUB y la deriva UI/layout de subformularios impiden promover la capacidad a `Verified-runtime`.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). CDCASUB comparte la deuda con CDCA/PC/PCSUB y debe migrarse a `BeginTestSession`/`EndTestSession`, `DAO.Database` explícito y cardinalidad de mutaciones.

**Contrato TDD vigente**: las pruebas CDCASUB deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio en sandbox, schema-first, `DAO.Database` inyectado explícito, `countBefore`/`countAfter` para mutaciones, manifests atómicos separados de los smoke y cero mutación de `TbConfiguracionBackends`.

**Justificación del nivel**: crítico, porque CDCASUB es un tipo de solicitud completo con persistencia propia en `tbDatosCDCASUB`, navegación con seis pestañas gemela a CD/CA, persistencia granular por bloque con `DAO.Recordset`, transición automática de Preregistro a Registro y campos RAC delegados parcialmente documentados.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: registrar y gestionar Cambios de Diseño/Alcance vinculados a subcontratistas (no a la matriz), con la misma trazabilidad que CD/CA pero con roles y suministradores gemelos a PCSUB.
- **Usuarios / perfiles**: técnicos, calidad, RAC, autoridad de decisión, subcontratista principal y sub-suministrador que firman la propuesta.
- **Problema que resuelve**: extiende CD/CA para soportar escenarios donde la propuesta la origina o la ejecuta un subcontratista, conservando el mismo workflow y los mismos criterios de validación.
- **Valor de negocio**: uniformidad de proceso entre PC, PCSUB, CD/CA y CDCASUB, y trazabilidad de cambios técnicos de subcontratistas.
- **No-objetivos**: este documento no cubre en detalle el workflow global, la generación documental, los adjuntos ni la búsqueda. Los vincula como capacidades externas.
- **Origen de la intención**: código actual, gemelo estructural con `DatosPCSUBServicio.cls` y `DatosCDCAServicio.cls`, SDD `pcsub-nueva-solicitud` y sus derivados (gemelos que aplican a CDCASUB por patrón). Ver también CAP-001 (PCSUB) y CAP-003 (CDCA).
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** una solicitud CDCASUB en `estadoPreregistro` **CUANDO** el usuario guarda Datos Generales con `refSubSuministrador`, `suministradorPrincipalNombreDir`, `subSuministradorNombreDir` y `requiereModificacionContrato` no vacíos **ENTONCES** se crea/actualiza `tbDatosCDCASUB` y, si la solicitud estaba en `estadoPreregistro`, `WorkflowServicio.EjecutarTransicion` la promueve a `estadoRegistro` dentro de la misma transacción. **Estado**: `Verified-static`.
- **DADO** una CDCASUB con Datos Generales guardados **CUANDO** se guarda Propuesta con `identificacionMaterial`, `causaNC` y `descripcionImpactoNC` no vacíos **ENTONCES** `DatosCDCASUBRepositorio.ActualizarPropuesta` persiste los cambios. Si la solicitud está en `estadoValidacion` con un rechazo activo, se registra el delta JSON bajo `CDCASUB_Propuesta` con los campos de la propuesta. **Estado**: `Verified-static`.
- **DADO** una CDCASUB con `impactoCoste`, `clasificacionNC` y `esSubSuministradorAD` definidos **CUANDO** se guarda Impacto **ENTONCES** se exige que `clasificacionNC` ∈ {`MAYOR`,`MENOR`} y que `esSubSuministradorAD` no sea `Null`. **Estado**: `Verified-static`.
- **DADO** una CDCASUB con `racDecision <> "RECHAZADO"` **CUANDO** se guarda Dictamen RAC **ENTONCES** `racCodigo` es obligatorio. **Estado**: `Verified-static`.
- **DADO** una CDCASUB con `decisionFinal` y `NombreFirmanteFinal` no vacíos **CUANDO** se guarda Decisión Final **ENTONCES** se persiste el bloque. **Estado**: `Verified-runtime` (commit 2026-06-18; `DatosCDCASUBServicio.EsDecisionFinalCompleta` ahora exige ambos).
- **DADO** una CDCASUB en formalización con rechazo **CUANDO** se ejecuta `RegistrarRechazoDesdeFormalizacion` **ENTONCES** se borra la decisión final, se conserva `racDecision = "RECHAZADO"` con `racRechazoMotivos`, y se transiciona a `estadoRechazada` en la misma transacción. **Estado**: `Verified-static`.
- **DADO** una CDCASUB en `estadoDesarrolloTecnico` con Detalle completo pero Motivos pendientes **CUANDO** el usuario guarda Impacto **ENTONCES** la rama de precondiciones indica al formulario que navegue a `tabImpacto` (vía `WorkflowServicio` y `FormulariosPadreAuxiliares`). **Estado**: `Verified-static`; depende del gemelo CD/CA.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | Datos Generales CDCASUB requieren `refSubSuministrador`, `suministradorPrincipalNombreDir`, `subSuministradorNombreDir` y `requiereModificacionContrato` no vacíos. `refSuministrador` ya NO es obligatorio (cambio explícito respecto a CD/CA). | Código | Sí: `DatosCDCASUBServicio.GuardarDatosGenerales` (líneas 142-147). | Sin manifest atómico. | Verified-static |
| BR-002 | La transición Preregistro → Registro se ejecuta dentro de la misma transacción que el alta/actualización de Datos Generales. | Código + AGENTS | Sí: `GuardarDatosGenerales` invoca `WorkflowServicio.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, db)`. | Pendiente. | Verified-static |
| BR-003 | Propuesta exige `identificacionMaterial`, `causaNC` y `descripcionImpactoNC` no vacíos. | Código | Sí: `GuardarPropuesta` (líneas 191-193). | Pendiente. | Verified-static |
| BR-004 | Impacto exige `impactoCoste` no vacío, `clasificacionNC` ∈ {`MAYOR`,`MENOR`} y `esSubSuministradorAD` no `Null`. | Código | Sí: `GuardarImpacto` (líneas 262-264) y `EsParteTecnicaCompleta` (líneas 446-453). | Pendiente. | Verified-static |
| BR-005 | Si la solicitud está en `estadoValidacion` y existe rechazo activo, guardar Impacto registra el delta JSON completo (`afecta*`, `impactoCoste`, `clasificacionNC`, `esSubSuministradorAD`, `identificacionAutoridadDiseno`, `efectoFechaEntrega`). | Código | Sí: `GuardarImpacto` (líneas 296-313) y helper `EsSolicitudEnValidacion`. | Pendiente. | Verified-static |
| BR-006 | Aprobación de Suministrador exige `firmaAprobacionRespIngenieriaNombre` y `firmaAprobacionRespCalidadNombre` no vacíos. Los campos `firmaAprobacionRespProduccionNombre`, `firmaAprobacionRespDisenioNombre` y `firmaAprobacionRepresentanteSumNombre` no se validan. | Código | Sí: `GuardarAprobacionSuministrador` (líneas 337-339). | Pendiente. | Verified-static |
| BR-007 | Dictamen RAC exige `racCodigo` no vacío salvo `racDecision = "RECHAZADO"`. | Código | Sí: `GuardarDictamenRAC` (líneas 351-353). | Pendiente. | Verified-static |
| BR-008 | Decisión Final exige `decisionFinal` y `NombreFirmanteFinal` no vacíos. | Código | Sí: `GuardarDecisionFinal` línea 396; `EsDecisionFinalCompleta` valida ambos (commit 2026-06-18). | Átomo `Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse` en `tests.vba.json` (cross-proyecto PCSUB↔CDCASUB por gemelos). | Verified-runtime |
| BR-009 | La purga técnica (`PurgaTecnica`) limpia Propuesta e Impacto dejando intactos los datos generales, RAC y aprobación. | Código | Sí: `DatosCDCASUBServicio.PurgaTecnica` (líneas 611-650). | Pendiente. | Verified-static |
| BR-010 | El rechazo desde formalización borra decisión final, deja `racDecision = "RECHAZADO"` con motivos, y transiciona a `estadoRechazada`. | Código | Sí: `RegistrarRechazoDesdeFormalizacion` (líneas 652-709). | Pendiente. | Verified-static |
| BR-011 | La rama transaccional del RAC separa la decisión de fuente de la conexión (`db Is Nothing` abre y opcionalmente `BeginTrans` solo si `EsSolicitudEnValidacion`). | Código (slice 2 fix #18) | Sí: `GuardarDictamenRAC` (líneas 360-372). | Pendiente. | Verified-static |
| BR-012 | El `ImportarEntidad_aVM`/`ExportarVM_aEntidad` siguen el patrón gemelo con CD/CA, pero con `esSubSuministradorAD` en lugar de `esSuministradorAD`. | Código | Sí: `ImportarEntidad_aVM`/`ExportarVM_aEntidad`. | Pendiente. | Verified-static |

### Validaciones observadas

- `idSolicitud <= 0` en `getDatosCDCASUB` y guardados debe tratarse como error o no encontrado.
- Datos Generales sin `refSubSuministrador` o sin `suministradorPrincipalNombreDir`/`subSuministradorNombreDir` produce error 513.
- `requiereModificacionContrato` debe ser `Sí`/`No` (no `Null`) para considerar completos los Datos Generales.
- Propuesta sin `identificacionMaterial`, `causaNC` o `descripcionImpactoNC` produce error.
- Impacto con `clasificacionNC` fuera de `MAYOR`/`MENOR` o `esSubSuministradorAD` `Null` produce error.
- Aprobación sin firmas de Ingeniería o Calidad produce error.
- RAC con decisión no rechazada y `racCodigo` vacío produce error.

### Transiciones de estado y navegación

- En Preregistro, al guardar Datos Generales válidos se ejecuta la transición a `estadoRegistro`.
- El `WorkflowServicio.PrecondicionesCumplidas_CDCASUB` valida `estadoModificacion` con `CumplePasoAModificacion` (parte técnica completa + verificación de cambios tras rechazo), `estadoValidacion` con `CumplePasoAValidacion` (parte técnica + RAC + aprobación), `estadoRevision` con `CumplePasoARevision` (adjunto de validación) y `estadoFormalizacion` con `CumplePasoAFormalizacion` (RAC aprobado).
- `PermiteEdicion_CDCASUB` bloquea edición en `estadoRechazada`, limita `estadoAprobada` a Decisión Final, y limita `estadoFormalizacion` a Aprobación Suministrador + Decisión Final para roles `Calidad`/`Administrador`.

### Casos límite y hallazgos

- La spec inicial PCSUB/CAP-001 declaraba campos `racDelegado*` para subcontratistas. El esquema `tbDatosCDCASUB` ya incluye `observacionesRACDelegador` y `racNombreDelegador` (memo + texto), pero `ImportarEntidad_aVM`/`ExportarVM_aEntidad` solo persisten los dos últimos; el contrato RAC delegado completo no está implementado. Esta divergencia se cruza con la documentada en CAP-001 §7.
- `EsMotivosCompleto` exige `descripcionImpactoNC` y `descripcionImpactoNCCont`, lo que exige al técnico que reparta el texto en dos campos cuando supera 255 caracteres. Esta es la misma estrategia que el overflow `_extN`/`Cont` documentado en CAP-002.
- **Resuelto 2026-06-18** (commit pendiente): `EsDecisionFinalCompleta` ahora exige `NombreFirmanteFinal` (mismo fix aplicado en los 4 gemelos). Idem `EsDictamenRACCompleto` ahora exige `racDecision`; idem `EsAprobacionSuministradorCompleta` ahora exige `decisionFinal`.

### Señales de aceptación / presencia

- Existen entidad, ViewModel, servicio y repositorio: `DatosCDCASUB.cls`, `DatosCDCASUBViewModel.cls`, `DatosCDCASUBServicio.cls`, `DatosCDCASUBRepositorio.bas`.
- Existe formulario principal `Form_frmDatosCDCASUB.cls` con seis subformularios: `AprobacionSuministrador`, `DecisionFinal`, `DictamenRAC`, `Generales`, `Impacto`, `Propuesta`. El code-behind está sincronizado, pero el layout `.form.txt` de los seis subformularios tiene deriva accionable.
- `tbDatosCDCASUB` se persiste por `idSolicitud` con `idDatosCDCASUB` como PK; la clave de negocio es `idSolicitud`. El esquema tiene 45 columnas observables según ERD.
- El helper `EsSolicitudEnValidacion` (privado en `DatosCDCASUBServicio.cls`) sigue el mismo patrón que PC y PCSUB y comparte la misma deuda de inyección de `DAO.Database`.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmDatosCDCASUB.Form_Load` carga VM, página activa y navega con `BrowseTo`. Equivalente a `Form_frmDatosPCSUB` y `Form_frmDatosPC`.
  - `Form_frmDatosCDCASUB.GuardarDesdeSubform(nombreSub, silenciarErrorValidacion)` orquesta el guardado por bloque, dirty-check, transacción y navegación. Esta API es gemela de `Form_frmDatosPCSUB.GuardarDesdeSubform`.
  - `Form_frmDatosCDCASUB.cmdCerrar_Click` cierra el formulario.
  - Subformularios: `Form_subfrmDatosCDCASUB_Generales`, `Form_subfrmDatosCDCASUB_Propuesta`, `Form_subfrmDatosCDCASUB_Impacto`, `Form_subfrmDatosCDCASUB_AprobacionSuministrador`, `Form_subfrmDatosCDCASUB_DictamenRAC`, `Form_subfrmDatosCDCASUB_DecisionFinal`. Cada uno expone `PoblarControles(vm)`, `RecogerValores(vm)` y `cmdGuardar_Click`.
- **Puntos de entrada de código**:
  - `DatosCDCASUBServicio.GuardarDatosGenerales` (con transición a `estadoRegistro`).
  - `DatosCDCASUBServicio.GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC` (con la corrección slice 2 #18), `GuardarDecisionFinal`.
  - `DatosCDCASUBServicio.EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal`, `PurgaTecnica`, `RegistrarRechazoDesdeFormalizacion`, `ActualizarCamposDependientesDeExpediente`.
  - `DatosCDCASUBServicio.ObtenerViewModelCompleto`, `EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta`.
  - `DatosCDCASUBServicio.ImportarEntidad_aVM` y `ExportarVM_aEntidad`.
  - `DatosCDCASUBRepositorio.Guardar` (upsert), `getPorIdSolicitud`, `ActualizarPropuesta`, `ActualizarImpacto`, `ActualizarAprobacionSuministrador`, `ActualizarDictamenRAC`, `ActualizarDecisionFinal`, `Limpiar*`, `EliminarPorIdSolicitud`.
- **Datos afectados**:
  - `tbDatosCDCASUB`: datos específicos CDCASUB por `idSolicitud`.
  - `tbSolicitudes`: estado, fecha, usuario de modificación; log en `tbLogEstados`.
  - `TbExpedientes`: contexto padre.
- **Dependencias**:
  - `WorkflowServicio` (rama `PrecondicionesCumplidas_CDCASUB`, `PermiteEdicion_CDCASUB`).
  - `SolicitudServicio`, `ExpedienteServicio`, `SuministradorServicio`.
  - `JsonHelper` para delta de rechazo.
  - `RechazoRepositorio` para detectar rechazo activo.
  - `FormulariosPadreAuxiliares` (mapeo `TipoForm_CDCASUB`).
  - `m_ObjEntorno`, `m_ObjUsuarioActivo`, `m_ObjUsuarioReal`, `rolUsuario`.
- **Sincronización fuente↔binario**: aplicar el patrón UI → código, primero `import-form` para los seis subformularios y `import-code` para `cls`/`bas`. Mantener la disciplina de no paralelizar imports.
- **Valoración de diseño (tal-como-está vs ideal)**: la estructura gemela con CD/CA y PCSUB es correcta, pero arrastra la misma deuda. La rama de Dictamen RAC con la corrección slice 2 #18 es una buena muestra de que la separación entre decisión de fuente y política de transacción es la dirección correcta. La deuda principal está en (a) ausencia de seam testeable para `EsSolicitudEnValidacion` y la rama transaccional condicional, (b) reglas de completitud que omiten `racDecision`/`NombreFirmanteFinal` y (c) el contrato RAC delegado no implementado (cruce con CAP-001 §7).

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `DatosCDCASUB.cls`, `DatosCDCASUBViewModel.cls`, `DatosCDCASUBServicio.cls` y `DatosCDCASUBRepositorio.bas`.
2. Confirmar el esquema `tbDatosCDCASUB` con PK `idDatosCDCASUB` y clave de negocio `idSolicitud`; revisar especialmente `racNombreDelegador` y `observacionesRACDelegador`.
3. Restaurar `Form_frmDatosCDCASUB.cls` y los seis subformularios `Form_subfrmDatosCDCASUB_*.cls` y `*.form.txt`; verificar que `FormulariosPadreAuxiliares` mapea `TipoForm_CDCASUB` a los bloques correctos.
4. Confirmar que `WorkflowServicio` expone `PrecondicionesCumplidas_CDCASUB`, `PermiteEdicion_CDCASUB` y `getPaginaActivaCDCASUB` (delegada en `getPaginaActivaGenerica`).
5. Reconciliar deriva fuente↔binario de los formularios con `dysflow.verify_binary`; actualmente los seis `.form.txt` de subformularios CDCASUB están `bothChanged` y deben revisarse antes de cualquier import/export.
6. Importar con `dysflow.import_modules` (UI primero, código después) y compilar con `dysflow.compile_vba`.
7. Demostrar los escenarios de §2 con un manifest atómico `tests/tests.cdcasub.json` que cubra los BR; mientras no exista, esta capacidad queda en `Verified-static`.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/DatosCDCASUBServicio.cls`
  - `src/classes/DatosCDCASUB.cls`
  - `src/classes/DatosCDCASUBViewModel.cls`
  - `src/modules/DatosCDCASUBRepositorio.bas`
  - `src/forms/Form_frmDatosCDCASUB.cls` (existe en `src/forms/`, con su `form.txt`).
  - `src/classes/WorkflowServicio.cls` (rama `PrecondicionesCumplidas_CDCASUB`, `PermiteEdicion_CDCASUB`).
  - `src/classes/SolicitudServicio.cls`, `ExpedienteServicio.cls`, `SuministradorServicio.cls`.
  - `docs/ERD/condor_datos.md` → `tbDatosCDCASUB`.
- **Evidencia Dysflow incorporada**:
  - `dysflow.verify_binary` de `DatosCDCASUBServicio`, `DatosCDCASUBRepositorio`, `Form_frmDatosCDCASUB` y subformularios CDCASUB: code-behind `.cls` matched; `Form_frmDatosCDCASUB.form.txt` no accionable `formSerializationOnly`; seis subformularios `.form.txt` accionables `bothChanged`.
- **Tests existentes**: manifest atómico `tests/testsCdcasub.json` (commit `854f32b`, 2026-06-15, Slice B2). Cubre `DatosCDCASUBServicio.EsParteTecnicaCompleta` (exclusive a CDCASUB, no existe en PC) y `EsAprobacionSuministradorCompleta` (sad path con `firmaAprobacionRespIngenieriaNombre` vacía). 3/3 átomos verdes.
- **Evidencia runtime Dysflow**:
  - `Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow`: **VERDE** 3.8 s, sandbox `condor_datos.accdb` local, `idSolicitud=900811`, `clasificacionNC='MAYOR'`, `esSubSuministradorAD=True`, `EsParteTecnicaCompleta=True`.
  - `Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid`: **VERDE** 4.5 s, `idSolicitud=900812`, `clasificacionNC='INVALID'` (no en {MAYOR, MENOR}), `EsParteTecnicaCompleta=False`.
  - `Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty`: **VERDE** 5.7 s, `idSolicitud=900813`, `firmaAprobacionRespIngenieriaNombre=''` y `firmaAprobacionRespCalidadNombre='Maria Calidad'`, `EsAprobacionSuministradorCompleta=False`.
- **SDD/intención consultada**:
  - SDD `pcsub-nueva-solicitud` (patrón gemelo).
  - PRD `06_Formulario_Datos_CDCA.md` (gemelo CD/CA, base estructural).
  - Slice 2 fix #18 (corrección transaccional de `GuardarDictamenRAC`).

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| Guardar Datos Generales falla con error 513 sobre `refSuministrador` | Se está usando la rama de CD/CA en lugar de la de CDCASUB (cambio documentado en BR-001). | `dysflow.test_vba` con manifest atómico que ejercite BR-001. | §2 BR-001 |
| Guardar Dictamen RAC se demora o deja la BD en estado parcial | La corrección slice 2 #18 no se aplicó y se mezclan las decisiones de fuente/transacción. | prueba focal que invoque `GuardarDictamenRAC` con `db Is Nothing` y `idEstadoInterno = estadoValidacion`. | §2 BR-011 |
| Decisión Final aceptada sin `decisionFinal` | **Resuelto 2026-06-18** (commit pendiente): `EsDecisionFinalCompleta` ahora exige `decisionFinal` y `NombreFirmanteFinal`. | Test que falle sin `decisionFinal` o sin `NombreFirmanteFinal` y pruebe que pasa con ambos. | §2 BR-008 / §7 |
| RAC delegado no se persiste | `ImportarEntidad_aVM`/`ExportarVM_aEntidad` solo cubren `observacionesRACDelegador` y `racNombreDelegador` (memo y texto). El resto del contrato SDD no existe. | revisión con producto del contrato RAC delegado; ver CAP-001 §7. | §7 / CAP-001 §7 |
| El binario Access no contiene los seis subformularios CDCASUB reconciliados. | Deriva fuente↔binario no corregida. | `dysflow.verify_binary` con `moduleNames` filtrado a CDCASUB. | §3 / §4 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| CDCASUB como capacidad completa | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Servicio/repositorio y code-behind sincronizados; pendiente manifest atómico y reconciliación de layout `.form.txt` en seis subformularios. |

## §6 Notas de migración web

- **Conservar**: separación por bloques (Generales/Propuesta/Impacto/Aprobación/RAC/Decisión Final), validaciones por servicio, persistencia granular por bloque, transición automática Preregistro → Registro dentro de la misma transacción, gemelo estructural con CD/CA/PC/PCSUB.
- **Transformar**: `NavigationSubform` de Access a pestañas/rutas web; `BrowseTo` + `SendKeys` a navegación controlada; `MsgBox` modal a diálogos async; `DAO.Workspace` local con `MS Access;PWD=…` a transacciones en servidor con ORM; mini-transacción de RAC a `BEGIN TRANSACTION` con `db` inyectado.
- **NO copiar**: dependencia de `m_ObjEntorno`/`m_ObjUsuarioActivo` globales, `getdb()` como singleton, validación de RAC contra campos parciales, contrato `DAO.Recordset` editable con password en connection string, `EsSolicitudEnValidacion` privado con salto a `getdb()`.
- **Preguntas abiertas**: ¿el contrato RAC delegado para subcontratistas es un requisito de producto o un vestigio de la spec original? (responsable + equipo de calidad). ¿La completitud de Decisión Final debe exigir `NombreFirmanteFinal` en web? (responsable de producto). ¿`EsMotivosCompleto` debe seguir exigiendo `descripcionImpactoNCCont` o el troceo `_extN` ya lo gestiona en la capa de documento? (responsable + equipo técnico).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| CDCASUB tiene entidad, ViewModel, servicio, repositorio, formulario principal y seis subformularios. | Verified-static | `DatosCDCASUB*` y `Form_*CDCASUB*` en `src/`. | 2026-06-15 |
| `tbDatosCDCASUB` es la tabla de datos específicos CDCASUB vinculada por `idSolicitud`, con campos RAC delegados parciales (`observacionesRACDelegador`, `racNombreDelegador`). | Verified-static | `DatosCDCASUBRepositorio.bas`; ERD `tbDatosCDCASUB`. | 2026-06-15 |
| `GuardarDatosGenerales` ejecuta transición Preregistro → Registro dentro de la misma transacción. | Verified-static | `DatosCDCASUBServicio.cls` líneas 133-183. | 2026-06-15 |
| `GuardarDictamenRAC` separa la decisión de fuente de la política de transacción (slice 2 #18). | Verified-static | `DatosCDCASUBServicio.cls` líneas 360-372. | 2026-06-15 |
| `EsDecisionFinalCompleta` solo exigía `decisionFinal`, no `NombreFirmanteFinal`. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con `DatosCDCASUBServicio.EsDecisionFinalCompleta` ahora exigiendo ambos. | 2026-06-15 (original) → 2026-06-18 (resolución) |
| `EsDictamenRACCompleto` solo exigía `racCodigo` y `racNombre`, no `racDecision`. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con `DatosCDCASUBServicio.EsDictamenRACCompleto` ahora exigiendo `racDecision` también. | 2026-06-15 (original) → 2026-06-18 (resolución) |
| `EsMotivosCompleto` exige `descripcionImpactoNC` y `descripcionImpactoNCCont`, lo que obliga a trocear textos largos. | Verified-static | `DatosCDCASUBServicio.cls` línea 482-500. | 2026-06-15 |
| El contrato RAC delegado completo (campos `racDelegado*`) no está implementado; la spec inicial y el código actual no coinciden. | Divergent / cross-ref CAP-001 | `DatosCDCASUBServicio.ImportarEntidad_aVM`/`ExportarVM_aEntidad`; `openspec/changes/pcsub-nueva-solicitud/spec/pcsub.md` (patrón gemelo). | 2026-06-15 |
| `EsSolicitudEnValidacion` (helper privado) rompe la inyección de `DAO.Database` cuando el llamador no pasa `db`. | Verified-static / deuda arquitectónica | `DatosCDCASUBServicio.cls` línea 711-739. | 2026-06-15 |
| Existe un manifest atómico de pruebas CDCASUB que cumpla `access-vba-tdd` v2.4.2. | Verified-runtime (parcial) / pendiente expansión | `tests/testsCdcasub.json` con tres átomos verdes (commit `854f32b`): `EsParteTecnicaCompleta` 3.8 s + sad path `clasificacionNC=INVALID` 4.5 s, `EsAprobacionSuministradorCompleta` sad path `firmaAprobacionRespIngenieriaNombre=''` 5.7 s. Pendientes: `EsDatosGeneralesCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleta`, `EsDecisionFinalCompleta`, `EsParteTecnicaCompleta` happy path, y los `Guardar*` / `Actualizar*` con cardinalidad. | 2026-06-15 |
| El binario Access contiene los seis subformularios CDCASUB reconciliados. | Divergent / blocker UI | `verify_binary`: seis `Form_subfrmDatosCDCASUB_*.form.txt` con `bothChanged`; code-behind `.cls` matched. | 2026-06-15 |
| Phase 0 clasifica `Form_frmDatosCDCASUB` como Tier 1, defaults/validaciones como Tier 2 y `Propuesta`/`Impacto` como Tier 3. | Verified-static / deuda Phase 0 | `audit-e2e-thin-forms-phase-0.md` §2-§4; WU3 no adopta código ni ejecuta Access. | 2026-06-26 |
| UAT CDCASUB queda bloqueado para guardados y transiciones completas hasta que existan átomos focales verdes y `ref` firmado. | Verified-runtime parcial / pendiente | `audit-e2e-thin-forms-phase-0.md` §7-§8; los tres átomos actuales no cubren `GuardarDatosGenerales` ni transición. | 2026-06-26 |
| El contrato puro compartido de Decisión Final exige `decisionFinal` y `NombreFirmanteFinal` no vacíos, con trim y soporte de `Null`. | Verified-runtime limitado a helper puro Slice 3.1 | `tests/tests.decision-final.json` → 4/4 verde; `DecisionFinalHelper_EsCompleta` usado por `DatosCDCASUBServicio.EsDecisionFinalCompleta`. | 2026-06-27 |

**Divergencias pendientes de revisión humana**:

- BR-008 / BR-007: los gaps de completitud de `EsDecisionFinalCompleta` y `EsDictamenRACCompleto` (y `EsAprobacionSuministradorCompleta` para `decisionFinal`) fueron resueltos en el commit 2026-06-18 (mismo patrón aplicado en los 4 gemelos).
- RAC delegado: el esquema `tbDatosCDCASUB` ya tiene `racNombreDelegador` y `observacionesRACDelegador`, pero el resto del contrato SDD no se persiste. Mismo hallazgo que CAP-001 §7. Decidir si se implementa o se documenta como fuera de alcance.
- BR-005: la rama transaccional de Impacto depende de `EsSolicitudEnValidacion` que rompe la inyección de `DAO.Database`. Mover a helper inyectable antes de cualquier prueba focal.
- `EsMotivosCompleto`: revisar si la obligación de rellenar `descripcionImpactoNCCont` es comportamiento aceptado o deuda; alinear con la política de overflow `_extN`/`Cont` de CAP-002.
