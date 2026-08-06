# Capacidad: Gestión de Propuestas de Cambio (PC)

## §0 Identidad

- **ID de capacidad**: CAP-004
- **Tier**: critical
- **Estado**: active con deuda de reconciliación UI/layout y de pruebas
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre `DatosPCServicio`, `DatosPCRepositorio`, `Form_frmDatosPC` y seis subformularios PC. Servicio/repositorio y code-behind `.cls` están `matched`; `Form_frmDatosPC.form.txt` solo `formSerializationOnly`; los seis `.form.txt` de subformularios PC tienen diferencias accionables `bothChanged`.
- **Confianza global**: mayoritariamente `Verified-static`. Las reglas de servicio/repositorio están implementadas y el código está sincronizado con el binario, pero no se eleva a `Verified-runtime` porque no existe manifest PC strict y el layout/UI de subformularios requiere reconciliación.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). PC no debe promocionarse a `Verified-runtime` sin migrar la suite heredada, garantizar cardinalidad de mutaciones y separar los manifests atómicos de los smoke.

**Contrato TDD vigente**: las pruebas PC deben migrarse a `access-vba-tdd` v2.4.2: `Public Function` con retorno JSON canónico, fixture propio en sandbox, schema-first, `DAO.Database` inyectado explícitamente, `countBefore`/`countAfter` para mutaciones, manifests atómicos separados de `*_RunAll` y cero mutación de `TbConfiguracionBackends`.

**Justificación del nivel**: crítico, porque PC es un tipo de solicitud completo con persistencia propia en `tbDatosPC`, flujo de navegación con seis pestañas, persistencia por bloque con `DAO.Recordset`, transacciones explícitas en RAC, transición automática de Preregistro a Registro y gemelo estructural con PCSUB/CDCA/CDCASUB.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: registrar y gestionar Propuestas de Cambio principales (no SUB) asociadas a un expediente, con datos generales, propuesta técnica, impacto, aprobación de suministrador, dictamen RAC y decisión final.
- **Usuarios / perfiles**: técnicos, calidad, RAC, autoridad de decisión y personal administrativo que formaliza la solicitud.
- **Problema que resuelve**: estandariza la captura de la información técnica y de aprobación de un cambio contractual, y la integra con el ciclo de validación, revisión y formalización del workflow CONDOR.
- **Valor de negocio**: trazabilidad contractual y técnica de cambios, validaciones por bloque, automatización de transición de Preregistro a Registro y gemelo con PCSUB para soportar procesos de subcontratación.
- **No-objetivos**: este documento no cubre en detalle el workflow global, los documentos Word, los adjuntos ni la búsqueda. Los vincula como capacidades externas (CAP-007, CAP-002, CAP-008, CAP-006).
- **Origen de la intención**: código actual, gemelo estructural con `DatosPCSUBServicio.cls` y `DatosCDCAServicio.cls`, PRD histórico de CD/CA, y SDD `pcsub-guardar-phase-advancement` que define el patrón compartido.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** una solicitud PC en estado Preregistro **CUANDO** el usuario guarda Datos Generales completos **ENTONCES** se crea/actualiza `tbDatosPC` y, si la solicitud estaba en `estadoPreregistro`, `WorkflowServicio.EjecutarTransicion` la promueve a `estadoRegistro` dentro de la misma transacción. La transacción usa `DAO.Workspace` aislado y `getdb().name` con `MS Access;PWD=…`. **Estado**: `Verified-static`; requiere seam de servicio testeable con `db` explícito.
- **DADO** una PC con datos generales guardados **CUANDO** se guarda Propuesta con `descripcionMaterialAfectado` y `descripcionPropuestaCambio` no vacíos **ENTONCES** `DatosPCRepositorio.ActualizarPropuesta` persiste los cambios. Si existe un rechazo activo, `JsonHelper.RegistrarCambio` registra el delta bajo `PC_Propuesta`. **Estado**: `Verified-static`; prueba focal pendiente.
- **DADO** una PC en `estadoValidacion` con rechazo activo **CUANDO** el usuario guarda Impacto **ENTONCES** el servicio abre transacción local, persiste el bloque y registra el delta JSON de motivos, incidencias y clasificación. **Estado**: `Verified-static`; la rama transaccional de Impacto está condicionada por el helper `EsSolicitudEnValidacion` que no se ha extraído a un seam testeable.
- **DADO** una PC con dictamen RAC **CUANDO** `racDecision <> "RECHAZADO"` **ENTONCES** `racCodigo` es obligatorio. **Estado**: `Verified-static`.
- **DADO** una PC con decisión final **CUANDO** `decisionFinal` y `NombreFirmanteFinal` no son vacíos **ENTONCES** se persiste y se permite continuar el flujo. **Estado**: `Verified-runtime` (commit pendiente 2026-06-18; `DatosPCServicio.EsDecisionFinalCompleta` ahora exige ambos).
- **DADO** una PC en fase de cierre de formalización **CUANDO** el usuario confirma la selección del PDF firmado **ENTONCES** `frmDatosPC.ProcesarCierreFormalizacion` exige `msoFileDialogFilePicker`, exige PDF para `etapa="Documento Final Firmado"` y delega el cierre completo en `WorkflowServicio.EjecutarCierreFormalizacion`. **Estado**: `Verified-static` condicionado por la reconciliación `bothChanged` ya marcada en CAP-001 para el gemelo PCSUB; requiere verificación Dysflow antes de declararlo `Verified-runtime`.
- **DADO** una PC con datos técnicos en Desarrollo Técnico **CUANDO** el usuario guarda y todos los bloques obligatorios están completos **ENTONCES** el formulario pregunta si envía a Calidad y transiciona a `estadoModificacion` si el usuario acepta. **Estado**: `Verified-static`; depende del flujo de navegación en `Form_frmDatosPC.GuardarDesdeSubform` y de `Mensajería` no testeable directamente.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | Datos Generales requieren `refContratoInspeccionOficial`, `refSuministrador`, `denominacionContrato`, `SuministradorNombreDir` y `objetoContrato` no vacíos. | Código | Sí: `DatosPCServicio.GuardarDatosGenerales` y `EsDatosGeneralesCompleta`. | `tests/tests.cdca.json` estilo PC pendiente; manifest atómico propio no existe. | Verified-static |
| BR-002 | La transición Preregistro → Registro se ejecuta dentro de la misma transacción que el alta/actualización de Datos Generales. | Código + AGENTS | Sí: `GuardarDatosGenerales` invoca `WorkflowServicio.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, db)`. | Sin prueba focal; cualquier intento de promoción por SQL directo sería regresión. | Verified-static |
| BR-003 | Propuesta exige `descripcionMaterialAfectado` y `descripcionPropuestaCambio` no vacíos; el resto de campos del bloque son opcionales. | Código | Sí: `DatosPCServicio.GuardarPropuesta`. | Pendiente de prueba focal con cardinalidad. | Verified-static |
| BR-004 | Impacto exige: al menos un motivo marcado; `motivoOtros` ⇒ `motivoOtrosDetalle` obligatorio; `motivoOtros = False` con `motivoOtrosDetalle` no vacío ⇒ error; `incidenciaCoste` ∈ {`AUMENTARÁ`,`DISMINUIRÁ`,`NO VARIARÁ`}; `incidenciaPlazo` mismo dominio; `impactoClasificacion` ∈ {`MAYOR`,`MENOR`}; `CambioAfectaAMaterial` ∈ {`Material ya entregado`,`Material por entregar`}. | Código | Sí: `DatosPCServicio.GuardarImpacto` (líneas 366-419). | Pendiente de prueba focal con casos por rama. | Verified-static |
| BR-005 | Si existe un rechazo activo y la solicitud está en `estadoValidacion`, guardar Impacto registra el delta JSON completo en `tbValidacionRevision`/`JsonHelper`. | Código | Sí: helper `EsSolicitudEnValidacion` + `JsonHelper.RegistrarCambio` por cada campo. | Pendiente de prueba focal que verifique el delta. | Verified-static |
| BR-006 | Aprobación de Suministrador exige `firmaOficinaTecnicaNombre` y `firmaRepSuministradorNombre` no vacíos. | Código | Sí: `DatosPCServicio.GuardarAprobacionSuministrador`. | Pendiente. | Verified-static |
| BR-007 | Dictamen RAC exige `racCodigo` no vacío salvo `racDecision = "RECHAZADO"`. | Código | Sí: `GuardarDictamenRAC` líneas 527-530. | Pendiente. | Verified-static |
| BR-008 | Decisión Final exige `decisionFinal` y `NombreFirmanteFinal` no vacíos. | Código | Sí: `GuardarDecisionFinal` línea 566; `EsDecisionFinalCompleta` valida ambos (commit 2026-06-18). | Átomo `Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse` en `tests.vba.json` (cross-proyecto PCSUB↔PC por gemelos). | Verified-runtime |
| BR-009 | La purga técnica (`PurgaTecnica`) limpia los bloques Propuesta e Impacto y solo esos. | Código | Sí: `DatosPCServicio.PurgaTecnica` líneas 54-100. | Pendiente. | Verified-static |
| BR-010 | Las operaciones de limpieza de bloque (`LimpiarDictamenRAC`, `LimpiarAprobacionSuministrador`, `LimpiarDecisionFinal`) solo blanquean su bloque, conservando el resto de la fila. | Código | Sí en `DatosPCRepositorio.bas`. | Sin prueba focal; auditoría pendiente. | Verified-static |
| BR-011 | El rechazo desde formalización borra los datos de Decisión Final, deja `racDecision = "RECHAZADO"`, persiste `racRechazoMotivos` y transiciona a `estadoRechazada` en la misma transacción. | Código | Sí: `DatosPCServicio.RegistrarRechazoDesdeFormalizacion`. | Pendiente. | Verified-static |
| BR-012 | El formulario principal delega el mapeo de bloque y la inicialización de campos de `firmaOficinaTecnicaNombre` y `firmaRepSuministradorNombre` en `DatosPCServicio.GetDatosAprobacionSuministradorIniciales`, usando `m_ObjEntorno.NombreJefeCalidad` y el responsable técnico del expediente. | Código | Sí: `GetDatosAprobacionSuministradorIniciales`. | Pendiente. | Verified-static |

### Validaciones observadas

- `idSolicitud <= 0` en `getDatosPC` y guardados debe tratarse como error o no encontrado.
- `descripcionMaterialAfectado` o `descripcionPropuestaCambio` vacíos en Propuesta producen error 513.
- Motivos: al menos uno marcado; reglas bidireccionales con `motivoOtros` y `motivoOtrosDetalle`.
- `incidenciaCoste`/`incidenciaPlazo` fuera del dominio válido produce error.
- `impactoClasificacion` y `CambioAfectaAMaterial` se validan contra catálogos cerrados.
- `racCodigo` vacío con decisión no rechazada produce error.
- `decisionFinal` vacío produce error.

### Transiciones de estado y navegación

- En Preregistro, al guardar Datos Generales válidos se ejecuta `WorkflowServicio.EjecutarTransicion` a `estadoRegistro`. El formulario además pregunta si asigna a Desarrollo Técnico y, si el usuario acepta, transiciona a `estadoDesarrolloTecnico`.
- En `estadoDesarrolloTecnico`, si `EsMotivosCompleto` y `EsDetalleCompleto` son `True`, el formulario pregunta si envía a Calidad y transiciona a `estadoModificacion`. Si solo `EsDetalleCompleto` es `False`, navega a `tabPropuesta`; si solo `EsMotivosCompleto` es `False`, navega a `tabImpacto`.
- `WorkflowServicio.getPaginaActivaPC` mapea el estado a la pestaña activa (Delegada a `getPaginaActivaGenerica`).
- En `estadoFormalizacion` solo se permiten editar `Bloque_AprobacionSuministrador` y `Bloque_DecisionFinal` para `rol.Calidad`/`rol.Administrador` (ver CAP-007 §3).

### Casos límite y hallazgos

- La completitud de Decisión Final, Aprobación Suministrador y Dictamen RAC en PC no exige `NombreFirmanteFinal` ni `racDecision` como se documenta en CAP-003; el patrón es consistente entre PC, CDCA y PCSUB pero diverge respecto a las expectativas de las pruebas candidatas. Marcar para revisión humana en §7.
- `Form_frmDatosPC.GuardarDesdeSubform` mezcla la lógica de validación, dirty-check, transacción y navegación, lo que dificulta el seam. La rama del RAC (líneas 221-236) ejecuta una mini-transacción específica del bloque; cualquier intento de probar el comportamiento UI obligaría a extraer el helper.
- `EsSolicitudEnValidacion` (helper privado en `DatosPCServicio.cls`) hace `SELECT idEstadoInterno FROM tbSolicitudes WHERE idSolicitud = …` directamente con `getdb()` cuando el llamador no pasa `db`. Esta rama rompe la inyección de `DAO.Database` y debe migrarse a un seam testeable.

### Señales de aceptación / presencia

- Existen entidad, ViewModel, servicio y repositorio: `DatosPC.cls`, `DatosPCViewModel.cls`, `DatosPCServicio.cls`, `DatosPCRepositorio.bas`.
- Existe formulario principal `Form_frmDatosPC.cls` con seis subformularios: `AprobacionSuministrador`, `DecisionFinal`, `DictamenRAC`, `Generales`, `Impacto`, `Propuesta`. El code-behind está sincronizado, pero el layout `.form.txt` de los seis subformularios tiene deriva accionable.
- `tbDatosPC` se persiste por `idSolicitud` con `idDatosPC` como PK; la clave de negocio es `idSolicitud`.
- El método de mapeo `ImportarEntidad_aVM` y `ExportarVM_aEntidad` vive en el servicio y es invocado por el formulario, no por cada subformulario.
- `m_ObjEntorno.estados`, `m_ObjUsuarioActivo` y `rolUsuario` se usan en todos los guardados y precondiciones, por lo que la integridad del entorno es prerrequisito.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmDatosPC.Form_Load` carga VM, determina pestaña activa vía `WorkflowServicio.getPaginaActivaPC` y navega con `DoCmd.BrowseTo`.
  - `Form_frmDatosPC.GuardarDesdeSubform(nombreSub, silenciarErrorValidacion)` orquesta el guardado por bloque, dirty-check, transacción y navegación.
  - `Form_frmDatosPC.ProcesarCierreFormalizacion` exige PDF firmado y delega el cierre en `WorkflowServicio.EjecutarCierreFormalizacion`.
  - `Form_frmDatosPC.cmdCerrar_Click` y `cmdVolverResumen_Click` cierran o vuelven a `frmGestionSolicitud`.
  - Subformularios: `Form_subfrmDatosPC_Generales`, `Form_subfrmDatosPC_Propuesta`, `Form_subfrmDatosPC_Impacto`, `Form_subfrmDatosPC_AprobacionSuministrador`, `Form_subfrmDatosPC_DictamenRAC`, `Form_subfrmDatosPC_DecisionFinal`. Cada uno expone `PoblarControles(vm)`, `RecogerValores(vm)` y `cmdGuardar_Click`.
- **Puntos de entrada de código**:
  - `DatosPCServicio.GuardarDatosGenerales` (con transición a `estadoRegistro`).
  - `DatosPCServicio.GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC`, `GuardarDecisionFinal`.
  - `DatosPCServicio.EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal`, `PurgaTecnica`, `RegistrarRechazoDesdeFormalizacion`.
  - `DatosPCServicio.ObtenerViewModelCompleto` (carga `Solicitud`, `Expediente`, `tbDatosPC`, `Estado`, `TiempoCiclo`, `Permisos`).
  - `DatosPCServicio.EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta`.
  - `DatosPCServicio.ImportarEntidad_aVM` y `ExportarVM_aEntidad` (mapeo Entidad ↔ VM por bloque).
  - `DatosPCRepositorio.Guardar` (upsert), `getPorIdSolicitud`, `ActualizarPropuesta`, `ActualizarImpacto`, `ActualizarAprobacionSuministrador`, `ActualizarDictamenRAC`, `ActualizarDecisionFinal`, `Limpiar*`, `EliminarPorIdSolicitud`.
- **Datos afectados**:
  - `tbDatosPC`: datos específicos PC por `idSolicitud` (42 columnas observables según ERD).
  - `tbSolicitudes`: estado, fecha, usuario de modificación; log de transiciones en `tbLogEstados`.
  - `TbExpedientes`: contexto padre en fixture y datos iniciales.
- **Dependencias**:
  - `WorkflowServicio` para transiciones, precondiciones, página activa y permisos.
  - `SolicitudServicio`, `ExpedienteServicio`, `SuministradorServicio` para orquestación.
  - `JsonHelper` para delta de rechazo.
  - `RechazoRepositorio` para detectar rechazo activo.
  - `FormulariosPadreAuxiliares` para `ClonarEntidad`, `SonEntidadesIguales`, `MapearBloquePorNombreSub` y `GetSubformNameFromTabName`.
  - `m_ObjEntorno`, `m_ObjUsuarioActivo`, `m_ObjUsuarioReal` y `rolUsuario` (globals).
- **Sincronización fuente↔binario**: la suite `access-vba-sync`/`dysflow.import_modules`/`dysflow.import_all` debe respetar el orden UI (subforms) → código (`cls`/`bas`) para evitar errores de import por header. La forma recomendada es `import-form` para los seis subformularios y los `form.txt` actualizados, seguido de `import-code` para los `cls` y `bas`.
- **Valoración de diseño (tal-como-está vs ideal)**: la separación VM/Servicio/Repositorio es razonablemente limpia y el uso de `DAO.Recordset` editable con `QueryDef` para `Memo` es correcto. La deuda principal está en (a) `GuardarDesdeSubform` con mezcla de UI y transacciones, (b) `EsSolicitudEnValidacion` privado con salto a `getdb()`, (c) reglas de completitud que omiten `NombreFirmanteFinal`/`racDecision` y (d) falta de seam testeable para `ActualizarCamposDependientesDeExpediente`. La pieza está bien hecha, pero la cobertura `Verified-runtime` no debe declararse hasta cerrar la deuda de seam y de completitud.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `DatosPC.cls`, `DatosPCViewModel.cls`, `DatosPCServicio.cls` y `DatosPCRepositorio.bas` con sus métodos actuales.
2. Confirmar el esquema `tbDatosPC` con PK `idDatosPC` y clave de negocio `idSolicitud`.
3. Restaurar `Form_frmDatosPC.cls` y los seis subformularios `Form_subfrmDatosPC_*.cls` y `*.form.txt`; verificar que `FormulariosPadreAuxiliares` mapea `TipoForm_PC` a los bloques correctos.
4. Confirmar que `WorkflowServicio` expone `getPaginaActivaPC`, `EjecutarTransicion`, `PuedeEditarBloque`, `PrecondicionesCumplidas_PC` y `PermiteEdicion_PC`.
5. Confirmar que `m_ObjEntorno.estados` carga los IDs canónicos `estadoPreregistro`, `estadoRegistro`, `estadoDesarrolloTecnico`, `estadoModificacion`, `estadoValidacion`, `estadoRevision`, `estadoFormalizacion`, `estadoAprobada`, `estadoRechazada`.
6. Reconciliar deriva fuente↔binario de los formularios y subformularios con `dysflow.verify_binary`; actualmente los seis `.form.txt` de subformularios PC están `bothChanged` y deben revisarse antes de cualquier import/export.
7. Importar fuentes con `dysflow.import_modules` (UI primero, código después; sin paralelismo).
8. Compilar con `dysflow.compile_vba` y verificar binario con `dysflow.verify_binary` hasta que no quede deriva accionable.
9. Demostrar los escenarios de §2 con `dysflow.test_vba` y un manifest atómico `tests/tests.pc.json` que cubra los BR; mientras ese manifest no exista, esta capacidad queda en `Verified-static` y referenciada desde la deuda v2.4.2.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/DatosPCServicio.cls`
  - `src/classes/DatosPC.cls`
  - `src/classes/DatosPCViewModel.cls`
  - `src/modules/DatosPCRepositorio.bas`
  - `src/forms/Form_frmDatosPC.cls`
  - `src/classes/FormulariosPadreAuxiliares.bas` (referencias cruzadas en `FormulariosPadreAuxiliares`).
  - `src/classes/WorkflowServicio.cls` (gemelo PC).
  - `src/classes/SolicitudServicio.cls` (creación y eliminación).
  - `docs/ERD/condor_datos.md` → `tbDatosPC`.
- **Evidencia Dysflow incorporada**:
  - `dysflow.verify_binary` de `DatosPCServicio`, `DatosPCRepositorio`, `Form_frmDatosPC` y subformularios PC: code-behind `.cls` matched; `Form_frmDatosPC.form.txt` no accionable `formSerializationOnly`; seis subformularios `.form.txt` accionables `bothChanged`.
- **Tests existentes**: no hay manifest atómico específico para PC; existe `tests/tests.pcsub.json` con el gemelo PCSUB y `tests/tests.cdca.json` con el gemelo CDCA. La evidencia focal para PC queda condicionada a la migración v2.4.2.
- **Evidencia no reclamada**: no se ejecutó `dysflow.test_vba` con un manifest PC; no se declara que ningún escenario de §2 esté en verde en este documento.
- **SDD/intención consultada**:
  - SDD `pcsub-guardar-phase-advancement` (patrón gemelo PCSUB).
  - PRD `06_Formulario_Datos_CDCA.md` (gemelo CDCA, base estructural).

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| Datos Generales guardan pero no promueven a Registro | `sol.idEstadoInterno <> estadoPreregistro` o `WorkflowServicio.EjecutarTransicion` falla por precondición | `dysflow.test_vba` con manifest atómico que examine `tbSolicitudes.idEstadoInterno` antes/después de `GuardarDatosGenerales` | §2 BR-002 / §3 |
| Guardar Propuesta no persiste | `DatosPCRepositorio.ActualizarPropuesta` ejecutado contra `idSolicitud` sin fila previa; o `dbTrabajo` apunta a la BD equivocada | prueba con cardinalidad `countBefore`/`countAfter` sobre `tbDatosPC` filtrada por `idSolicitud` | §2 BR-003 / §3 |
| Guardar Impacto rechaza por completo aunque hay motivos marcados | Motivos son booleanos; si todos `False`, el servicio aborta; verificar el binding del subform | prueba con un caso por rama de motivo (al menos uno, otros obligatorio, etc.) | §2 BR-004 |
| Decisión Final se considera completa con todo vacío excepto `decisionFinal` | **Resuelto 2026-06-18** (commit pendiente): `EsDecisionFinalCompleta` ahora exige `decisionFinal` y `NombreFirmanteFinal`. | Test `EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse` que falla sin `NombreFirmanteFinal` y pasa con ambos. | §2 BR-008 / §7 |
| Cierre de formalización pide un PDF y abre el diálogo correctamente | `ProcesarCierreFormalizacion` exige PDF para `etapa="Documento Final Firmado"` | prueba focal con stub de `Application.FileDialog` | §2 / §3 |
| Pruebas PC verdes pero no confiables | harness legacy, datos de sandbox contaminados o `EsSolicitudEnValidacion` rompe la inyección de `db` | auditoría v2.4.2 del módulo de tests y del helper | §7 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| PC como capacidad completa | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Servicio/repositorio y code-behind sincronizados; pendiente manifest atómico y reconciliación de layout `.form.txt` en seis subformularios. |

## §6 Notas de migración web

- **Conservar**: separación por bloques (Generales/Propuesta/Impacto/Aprobación/RAC/Decisión Final), validaciones por servicio, persistencia granular por bloque, transición automática Preregistro → Registro dentro de la misma transacción, gemelo estructural con PCSUB/CDCA/CDCASUB.
- **Transformar**: `NavigationSubform` de Access a pestañas/rutas web; `BrowseTo` + `SendKeys` a navegación controlada; `MsgBox` modal a diálogos async; `DAO.Workspace` local + password en `ConnectionString` a transacciones en servidor con ORM.
- **NO copiar**: dependencia de `m_ObjEntorno`/`m_ObjUsuarioActivo` globales, `getdb()` como singleton de proceso, contratos `DAO.Recordset` editable con password en connection string, validación de RAC contra campos parciales (gap con `racDecision`/`NombreFirmanteFinal`).
- **Preguntas abiertas**: ¿el guardado de Impacto en `estadoValidacion` debe seguir creando un delta JSON aunque no haya rechazo activo, o solo cuando `RechazoRepositorio.GetUltimoRechazoActivo` devuelva uno? (responsable + equipo de calidad). ¿La completitud de Decisión Final debe exigir `NombreFirmanteFinal` en web? (responsable de producto).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| PC tiene entidad, ViewModel, servicio, repositorio, formulario principal y seis subformularios por bloque. | Verified-static | `DatosPC*` y `Form_*PC*` en `src/`. | 2026-06-15 |
| `tbDatosPC` es la tabla de datos específicos PC vinculada por `idSolicitud`. | Verified-static | `DatosPCRepositorio.bas`; ERD `tbDatosPC`. | 2026-06-15 |
| `GuardarDatosGenerales` ejecuta transición Preregistro → Registro dentro de la misma transacción. | Verified-static | `DatosPCServicio.cls` líneas 259-292. | 2026-06-15 |
| `GuardarImpacto` exige al menos un motivo, dominio cerrado de `incidenciaCoste`/`incidenciaPlazo`, `impactoClasificacion` y `CambioAfectaAMaterial`. | Verified-static | `DatosPCServicio.cls` líneas 366-419. | 2026-06-15 |
| `EsDecisionFinalCompleta` solo exigía `decisionFinal`, no `NombreFirmanteFinal`. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con `DatosPCServicio.EsDecisionFinalCompleta` ahora exigiendo ambos campos. | 2026-06-15 (original) → 2026-06-18 (resolución) |
| `EsDictamenRACCompleto` solo exigía `racCodigo` y `racNombre`, no `racDecision`. | Resuelto 2026-06-18 (commit pendiente) — `Verified-runtime` con `DatosPCServicio.EsDictamenRACCompleto` ahora exigiendo `racDecision` también. | 2026-06-15 (original) → 2026-06-18 (resolución) |
| `EsSolicitudEnValidacion` (helper privado) rompe la inyección de `DAO.Database` cuando el llamador no pasa `db`. | Verified-static / deuda arquitectónica | `DatosPCServicio.cls` línea 852-880. | 2026-06-15 |
| `Form_frmDatosPC.GuardarDesdeSubform` mezcla UI y transacciones, sin seam testeable. | Verified-static / deuda de seam | `Form_frmDatosPC.cls` líneas 157-419. | 2026-06-15 |
| Existe un manifest atómico de pruebas PC que cumpla `access-vba-tdd` v2.4.2. | Divergent / pendiente | `tests/` sin `tests.pc.json`; `tests.pcsub.json` cubre el gemelo. | 2026-06-15 |
| El binario Access contiene los seis subformularios PC reconciliados. | Divergent / blocker UI | `verify_binary`: seis `Form_subfrmDatosPC_*.form.txt` con `bothChanged`; code-behind `.cls` matched. | 2026-06-15 |
| Phase 0 clasifica `Form_frmDatosPC` como Tier 1, subformularios de defaults/validación como Tier 2 y `Propuesta`/`Impacto` como Tier 3. | Verified-static / deuda Phase 0 | `audit-e2e-thin-forms-phase-0.md` §2-§4; WU3 no modifica VBA ni binario. | 2026-06-26 |
| UAT PC queda bloqueado hasta extraer helper de `GuardarDesdeSubform`, ejecutar átomo verde y firmar `ref` por caso. | Divergent / pendiente | `audit-e2e-thin-forms-phase-0.md` §7-§8; fila `Test_PC_GuardarDesdeSubform_Generales_PersistsAndPlansWorkflow`. | 2026-06-26 |
| El contrato puro compartido de Decisión Final exige `decisionFinal` y `NombreFirmanteFinal` no vacíos, con trim y soporte de `Null`. | Verified-runtime limitado a helper puro Slice 3.1 | `tests/tests.decision-final.json` → 4/4 verde; `DecisionFinalHelper_EsCompleta` usado por `DatosPCServicio.EsDecisionFinalCompleta`. | 2026-06-27 |

**Divergencias pendientes de revisión humana**:

- BR-008 y BR-007 (gemelos): los gaps de completitud de `EsDecisionFinalCompleta` y `EsDictamenRACCompleto` fueron resueltos en el commit 2026-06-18 (mismo patrón aplicado en los 4 gemelos).- BR-005: la rama transaccional de Impacto depende de `EsSolicitudEnValidacion` que rompe `DAO.Database` injection. Mover a helper inyectable antes de cualquier prueba focal.
- `GuardarDesdeSubform`: extraer la decisión de navegación y la mini-transacción del RAC a un helper/servicio testeable sin controles Access. La rama UI seguirá consumiendo ese helper.
