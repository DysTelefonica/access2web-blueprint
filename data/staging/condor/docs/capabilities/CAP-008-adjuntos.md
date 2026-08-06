# Capacidad: Gestión de adjuntos y archivo de documentación

## §0 Identidad

- **ID de capacidad**: CAP-008
- **Tier**: critical
- **Estado**: active con deuda de seams, cierre de Word y reconciliación UI/layout
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre `AdjuntosServicio`, `Adjunto*`, `AdjuntoRepositorio`, formularios de adjuntos, `Correo*` y `modSimuladorNotificaciones`. Servicio/repositorio/code-behind `.cls` están `matched`; `CorreoRepositorio` solo `caseOnly`; `Form_frmGestionAdjuntos.form.txt` y `Form_frmElegirEtapaAdjunto.form.txt` tienen diferencias accionables `bothChanged`.
- **Confianza global**: `Verified-runtime` parcial para la slice strict de servicio/repositorio (`GuardarAdjuntoDesdeArchivo` con happy-path + sad-paths idSolicitud<=0 + sad-paths ruta vacía + sad-paths ruta inexistente + sad-paths PDF-only final firmado + dedup nombre destino, `ActualizarFicheroAdjunto` con happy-path + sad-paths idAdjunto<=0 + sad-paths id inexistente, `EliminarAdjunto`, `EliminarAdjuntosPorEtapas`) y mayoritariamente `Verified-static` para el resto. La callback atómica en `frmGestionAdjuntos` y la deriva UI/layout siguen impidiendo promover toda la capacidad.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). Adjuntos ya tiene slice focal strict para servicio/repositorio incluyendo `ActualizarFicheroAdjunto` y dos sad paths de `Validar`; todavía faltan rollback de persistencia, BR-004 reglas 3-5 (256+ chars, etapaWF, usuarioSubida, fechaSubida) y callbacks UI.

**Contrato TDD vigente**: las pruebas de adjuntos deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio, schema-first, `DAO.Database` inyectado, cardinalidad `countBefore`/`countAfter`, manifests atómicos.

**Justificación del nivel**: crítico, porque los adjuntos son la fuente de verdad documental de cada fase y soporte del PDF de cierre que aprueba la solicitud.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: capturar, versionar, sustituir y eliminar los documentos asociados a una solicitud por etapa de workflow (Preregistro, Registro, Desarrollo Técnico, Modificación, Validación, Revisión, Formalización, Aprobada, Cierre) y conservar la trazabilidad de quién, cuándo y por qué se subió.
- **Usuarios / perfiles**: técnicos, calidad, RAC, administrador y todos los roles con permiso sobre la etapa correspondiente.
- **Problema que resuelve**: centraliza el archivo físico (`m_ObjEntorno.URLDirectorioDocumentacion`) y su registro en `tbAdjuntos`, y soporta el contrato de cierre de formalización (PDF firmado → aprobación).
- **Valor de negocio**: trazabilidad documental, sustitución del PDF de cierre, deduplicación por nombre, callback atómico de aprobación al subir el PDF de cierre.
- **No-objetivos**: este documento no cubre la generación del Word (CAP-002) ni la notificación de correo (CAP-007/NotificacionServicio), pero las vincula.
- **Origen de la intención**: código actual, gemelo con los formularios PC/PCSUB/CDCA/CDCASUB y con el visualizador web.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** una solicitud en cualquier estado **CUANDO** el usuario abre `frmGestionAdjuntos` con un `idSolicitud` válido **ENTONCES** el formulario muestra el listado de adjuntos con la traducción de `etapaWF` (puede ser ID numérico o nombre) a nombre legible. **Estado**: `Verified-static`; el botón "Añadir" está deshabilitado para `rol.Tecnico`.
- **DADO** una solicitud **CUANDO** el usuario selecciona una etapa vía `frmElegirEtapaAdjunto` y elige un archivo **ENTONCES** `AdjuntosServicio.GuardarAdjuntoDesdeArchivo` valida, copia al directorio de documentación, renombra si colisiona y persiste `tbAdjuntos` con `etapaWF = nombre_etapa`, `fechaSubida = Now`, `usuarioSubida` y `descripcion` por defecto. **Estado**: `Verified-static`.
- **DADO** una solicitud con etapa `Documento Final Firmado` **CUANDO** se sube un archivo que no es `.pdf` **ENTONCES** `GuardarAdjuntoDesdeArchivo` aborta con `Err.Raise 513`. **Estado**: `Verified-static`.
- **DADO** una solicitud con un `Documento Final Firmado` previo **CUANDO** se sube otro PDF de cierre **ENTONCES** `frmGestionAdjuntos.EjecutarFlujoSubida` pregunta si se sustituye; si el usuario acepta, elimina el existente y sube el nuevo. **Estado**: `Verified-static`.
- **DADO** una solicitud con PDF de cierre recién subido **CUANDO** la etapa es `Documento Final Firmado` **ENTONCES** la callback atómica persiste `decisionFinal = "APROBADO"` en la tabla del tipo correspondiente y transiciona a `estadoAprobada` en la misma transacción. **Estado**: `Verified-static`; la callback está acoplada al formulario.
- **DADO** una solicitud en `estadoAprobada` con un `Documento Final Firmado` **CUANDO** el usuario elimina ese adjunto **ENTONCES** `frmGestionAdjuntos` muestra una advertencia, borra la decisión final del tipo y revierte a `estadoFormalizacion` mediante `WorkflowServicio.RevertirAFaseAnterior`. **Estado**: `Verified-static`.
- **DADO** un adjunto existente **CUANDO** el usuario abre el doble clic **ENTONCES** se invoca `AbrirEnLocal` con la ruta `m_ObjEntorno.URLDirectorioDocumentacion & nombreArchivo`. **Estado**: `Verified-static`.
- **DADO** un adjunto de tipo `Generado` (creado por `WorkflowServicio.GenerarYAdjuntarDocumentoBorrador`) **CUANDO** el usuario intenta eliminarlo **ENTONCES** el botón "Eliminar" está deshabilitado. **Estado**: `Verified-static`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | `idSolicitud` debe ser positivo; `rutaArchivoOrigen` debe existir; el archivo es obligatorio. | Código | Sí: `AdjuntosServicio.GuardarAdjuntoDesdeArchivo` líneas 49-52. | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaRutaInexistente`, verde con Dysflow. | Verified-runtime parcial |
| BR-002 | Para `etapaWF = "Documento Final Firmado"` la extensión debe ser `pdf`. | Código | Sí: `GuardarAdjuntoDesdeArchivo` líneas 54-60. | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaNoPdfFinalFirmado`, verde con Dysflow. | Verified-runtime parcial |
| BR-003 | Si el archivo destino ya existe, se renombra con sufijo `_YYYYMMDD_HHMMSS`. | Código | Sí: `GuardarAdjuntoDesdeArchivo` líneas 67-72. | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_DeduplicaNombreDestinoExistente`, verde con Dysflow. | Verified-runtime parcial |
| BR-004 | `Validar` exige `idSolicitud > 0`, `etapaWF` no vacío, `nombreArchivo` no vacío y ≤255 caracteres, `usuarioSubida` no vacío, `fechaSubida` válida. | Código | Sí: `AdjuntosServicio.Validar` líneas 278-310 y guard temprano en `GuardarAdjuntoDesdeArchivo` línea 49 (idSolicitud) y línea 50 (rutaOrigen). | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_RechazaIdSolicitudInvalido` (cubre regla 1) y `..._RechazaRutaOrigenVacia` (cubre regla de entrada), ambos verdes con Dysflow. La regla 3 (≤255 chars) y la cascada de `Validar` para etapaWF/usuarioSubida/fechaSubida no cubiertas en runtime (deuda documentada: caso 256+ chars requiere crear archivo físico con nombre 256+ chars en Windows, costo/beneficio bajo para este pase; rama de `Validar` es `Private Sub` no testeable directamente sin refactor). | Verified-runtime parcial |
| BR-005 | La transacción se gestiona con `getWorkspace()` aislado y `ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())`. | Código + AGENTS | Sí: `GuardarAdjuntoDesdeArchivo` líneas 33-99. | Pendiente. | Verified-static |
| BR-006 | Si la persistencia falla tras la copia, se borra el archivo destino (`fso.DeleteFile destinoFull, True`). | Código | Sí: `GuardarAdjuntoDesdeArchivo` líneas 121-124. | **Deuda explícita (2026-06-15)**: dos intentos de test focal fallaron por restricciones del entorno. (a) PK collision via dummy row: `getSiguienteIDAdjunto` calcula `MAX+1` DESPUÉS de la siembra del dummy, así que nunca choca consigo mismo. (b) Validar rule 3 (256+ chars) requiere crear archivo físico con nombre 256+ chars; Windows MAX_PATH (260) bloquea esto cuando `%TEMP%\source\` ya mide ~85 chars. **Resolución**: BR-006 sigue `Verified-static`. Para desbloquear, el servicio necesita un seam testeable para forzar falla de persistencia (e.g., DAO.Database inyectable que pueda configurarse para fallar en INSERT, o flag interno `forceFail` para tests). | Verified-static / deuda arquitectónica |
| BR-007 | `EliminarAdjunto` borra el archivo físico y el registro en `tbAdjuntos` dentro de la misma transacción. | Código | Sí: `AdjuntosServicio.EliminarAdjunto` líneas 485-538. | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_EliminarAdjunto_BorraFilaYArchivo`, verde con Dysflow. | Verified-runtime parcial |
| BR-008 | `EliminarAdjuntosPorEtapas` admite `Variant` con `etapas` y compara contra `etapaWF` aceptando tanto nombre como ID numérico. | Código | Sí: `AdjuntosServicio.EliminarAdjuntosPorEtapas` líneas 414-454. | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_EliminarAdjuntosPorEtapas_BorraSoloEtapasObjetivo`, verde con Dysflow para nombres de etapa. | Verified-runtime parcial |
| BR-009 | Al subir un PDF de cierre, la callback atómica fija `decisionFinal = "APROBADO"` para el tipo correspondiente y transiciona a `estadoAprobada`. | Código | Sí: `AdjuntosServicio.SubirYCerrar` (líneas 545-588) y `AdjuntosServicio.EjecutarCallbackAprobacion` (líneas 591-700) — seam extraído del formulario en Slice A5 (2026-06-15). El formulario `frmGestionAdjuntos.EjecutarFlujoSubida` ahora delega en el seam y conserva la UI feedback (MsgBox, NotificarCambioExterno). | `tests/tests.adjuntos.json` → `Test_Adjuntos_Strict_SubirYCerrar_NoInvocaCallbackSiEtapaNoEsCierre` (verde 3211 ms): verifica que el seam NO invoca la callback para etapas distintas a "Documento Final Firmado". El happy path (callback se invoca y completa la transición a `estadoAprobada`) queda como WIP A5b: requiere que `WorkflowServicio.EjecutarTransicion` valide precondiciones de rol/estado, que no son fácilmente reproducibles en el sandbox del test. | Verified-runtime parcial |
| BR-010 | Al eliminar un `Documento Final Firmado` en `estadoAprobada`, se borra la decisión final del tipo y se revierte a `estadoFormalizacion`. | Código | Sí: `frmGestionAdjuntos.cmdEliminarAdjunto_Click` líneas 279-373. | Pendiente. | Verified-static / deuda UI |
| BR-011 | El visor web de adjuntos (`ABRIR_ADJUNTO:id`) usa `idAdjunto` para reconstruir la ruta. | Código | Sí: `frmWebVisor` ↔ `m_cacheTimeline` ↔ `WorkflowServicio.GenerarHTML_VisualizadorDeEstado`. | Pendiente. | Verified-static |
| BR-012 | `ActualizarFicheroAdjunto` reemplaza el archivo físico de un adjunto existente preservando intactos todos los campos de la fila en `tbAdjuntos` (idAdjunto, idSolicitud, etapaWF, nombreArchivo, fechaSubida, usuarioSubida, descripcion, TipoAccion). | Código | Sí: `AdjuntosServicio.ActualizarFicheroAdjunto` líneas 133-196. | `tests/tests.adjuntos.json` → 3 átomos strict TDD v2.4.2: `Test_Adjuntos_Strict_ActualizarFicheroAdjunto_ReemplazaArchivoSinTocarFila`, `..._RechazaIdAdjuntoInvalido`, `..._RechazaIdAdjuntoInexistente`, todos verdes con Dysflow. | Verified-runtime parcial |

### Validaciones observadas

- `idAdjunto <= 0` en `EliminarAdjunto` y `ActualizarFicheroAdjunto` produce error.
- `rutaArchivoOrigen` vacía o inexistente produce error.
- Para `Documento Final Firmado`, la extensión debe ser `pdf` (case-insensitive).
- `Validar` aplica las 5 reglas BR-004 antes de persistir.

### Transiciones de estado y navegación

- `frmGestionAdjuntos` no transiciona por sí solo, pero su callback atómica puede transicionar a `estadoAprobada` (BR-009).
- `frmElegirEtapaAdjunto` recoge la etapa seleccionada y la devuelve vía `Tag`, `EtapaSeleccionada`, `EtapaNombreSeleccionado`, `DescripcionAdjunto`.

### Casos límite y hallazgos

- La callback atómica que aprueba al subir el PDF de cierre está acoplada al formulario. Cualquier intento de probarla requiere extraer la decisión a un servicio. Es deuda gemela a la de CAP-001/CAP-004.
- `Eliminado de archivo si falla persistencia` está implementado, pero la copia es síncrona y bloqueante. Si el archivo es muy grande, el `lblBloqueo`/`DoCmd.Hourglass` solo se ve bien en operaciones rápidas.
- `EliminarAdjuntosPorEtapas` es polimórfico y compara con `IsNumeric`, pero la búsqueda `direccionUTE` y la comparación contra el nombre de etapa (`m_ObjEntorno.estados(etapaKey).nombreEstado`) pueden divergir si la cache de estados no está cargada.

### Señales de aceptación / presencia

- Existen `Adjunto.cls`, `AdjuntosServicio.cls`, `AdjuntoViewModel.cls`, `AdjuntoRepositorio.bas`.
- Existen `Form_frmGestionAdjuntos.cls` y `Form_frmElegirEtapaAdjunto.cls`; su code-behind está sincronizado, pero sus `.form.txt` tienen deriva accionable `bothChanged`.
- `tbAdjuntos` se persiste por `idAdjunto` (PK) y `idSolicitud` (FK a `tbSolicitudes`). `etapaWF` admite tanto ID numérico como nombre; la traducción a nombre legible se hace en `getAdjuntosViewModelPorSolicitud`.
- `m_ObjEntorno.URLDirectorioDocumentacion` es la ruta de archivo físico.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frmGestionAdjuntos.Form_Load` (con auto-carga por `Application.TempVars` para evitar el diálogo de selección de etapa).
  - `cmdAnadirAdjunto_Click` → `EjecutarFlujoSubida(etapa, nombreEtapa, descripcion)`.
  - `cmdEliminarAdjunto_Click` con la lógica de seguridad para `Documento Final Firmado` en `estadoAprobada`.
  - `cmdAbrirAdjunto_Click` y `lstAdjuntos_DblClick` → `AbrirEnLocal`.
  - `Form_frmElegirEtapaAdjunto` como subformulario modal de selección de etapa.
- **Puntos de entrada de código**:
  - `AdjuntosServicio.GuardarAdjuntoDesdeArchivo`, `ActualizarFicheroAdjunto`, `EliminarAdjunto`, `EliminarAdjuntosPorEtapas`.
  - `AdjuntosServicio.ExisteAdjuntoEtapa`, `GetIdAdjuntoPorEtapa`, `ExisteAdjuntoGeneradoCierre`, `GeneradoPDFParaFirma`.
  - `AdjuntosServicio.getAdjuntosPorSolicitud`, `getAdjuntosPorSolicitudTx`, `getAdjuntoPorID`, `getAdjuntosViewModelPorSolicitud` (con enriquecimiento de etapa).
  - `AdjuntosServicio.Validar` (privado).
  - `AdjuntoRepositorio.Guardar`, `getPorID`, `getPorIdSolicitud`, `getPorEtapa`, `Eliminar`, `ExisteAdjuntoEtapa`.
- **Datos afectados**:
  - `tbAdjuntos`: `idAdjunto`, `idSolicitud`, `etapaWF` (numérico o texto), `nombreArchivo`, `fechaSubida`, `usuarioSubida`, `descripcion`, `TipoAccion`.
  - `m_ObjEntorno.URLDirectorioDocumentacion`: directorio físico.
- **Dependencias**:
  - `WorkflowServicio` para `RevertirAFaseAnterior`, `EjecutarTransicion`, `GenerarYAdjuntarDocumentoBorrador`.
  - `DatosXServicio` (PC/PCSUB/CDCA/CDCASUB) para `GuardarDecisionFinal` y `EliminarDecisionFinal` en la callback atómica.
  - `m_ObjEntorno`, `m_ObjUsuarioActivo`, `rolUsuario`.
- **Sincronización fuente↔binario**: si se modifican `AdjuntosServicio` o `AdjuntoRepositorio`, basta `dysflow.import_code`. Si se modifican los formularios, `import-form` y verificar.
- **Valoración de diseño (tal-como-está vs ideal)**: la transacción de adjunto y la validación son correctas. La callback atómica del PDF de cierre está mezclada con UI y debería extraerse. La deduplicación de nombre es razonable. La deuda principal está en (a) la callback atómica en `frmGestionAdjuntos`, (b) la apertura con `AbrirEnLocal` que depende de la asociación del sistema operativo, (c) la falta de seam testeable para `GuardarAdjuntoDesdeArchivo`, (d) la limpieza de Word (CAP-002) que no aplica aquí pero comparte la responsabilidad de "Word cerrado al terminar". La pieza está bien hecha para el día a día; no se recomienda `Verified-runtime` sin seam.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `Adjunto.cls`, `AdjuntosServicio.cls`, `AdjuntoViewModel.cls` y `AdjuntoRepositorio.bas`.
2. Confirmar `tbAdjuntos` con PK `idAdjunto` y FK `idSolicitud`.
3. Confirmar `m_ObjEntorno.URLDirectorioDocumentacion` con la ruta correcta en el entorno de despliegue.
4. Restaurar `Form_frmGestionAdjuntos.cls` y `Form_frmElegirEtapaAdjunto.cls` con la auto-carga por `TempVars`.
5. Importar con `dysflow.import_modules` y compilar con `dysflow.compile_vba`. Verificar binario con `dysflow.verify_binary`; actualmente los dos formularios de adjuntos tienen `.form.txt` `bothChanged`.
6. Extender el manifest atómico `tests/tests.adjuntos.json` más allá de la slice strict actual: rollback de persistencia, callback atómica y eliminación de cierre siguen pendientes. Mientras tanto, la capacidad completa queda en `Verified-runtime` parcial.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/AdjuntosServicio.cls` (543 líneas).
  - `src/classes/Adjunto.cls`, `AdjuntoViewModel.cls`.
  - `src/modules/AdjuntoRepositorio.bas`.
  - `src/forms/Form_frmGestionAdjuntos.cls`, `Form_frmElegirEtapaAdjunto.cls`.
  - `src/classes/Correo.cls`, `src/modules/CorreoRepositorio.bas` (vinculación a notificación).
  - `src/modules/modSimuladorNotificaciones.bas` (simulador para tests).
  - `docs/ERD/condor_datos.md` → `tbAdjuntos`.
- **Evidencia Dysflow incorporada**:
  - `dysflow.verify_binary`: `AdjuntosServicio`, `Adjunto`, `AdjuntoViewModel`, `AdjuntoRepositorio`, `Correo`, `modSimuladorNotificaciones` y code-behind de formularios matched.
  - `CorreoRepositorio` solo diferencia no accionable `caseOnly`.
  - `Form_frmGestionAdjuntos.form.txt` y `Form_frmElegirEtapaAdjunto.form.txt` están `bothChanged` y requieren reconciliación UI.
- **Tests existentes**: `tests/tests.adjuntos.json` contiene 11 átomos strict TDD v2.4.2 para la slice de servicio/repositorio de adjuntos.
- **Evidencia runtime incorporada (2026-06-15)**:
  - `dysflow.import_modules` importó `Test_Adjuntos_Strict` en `CONDOR.accdb` con `allowWrites=true`.
  - `dysflow.compile_vba` devolvió `ok=true`.
  - `dysflow.test_vba` con `tests/tests.adjuntos.json` ejecutó 4/4 tests verdes: guardar/copia (3709 ms), eliminar fila+archivo (4386 ms), eliminar por etapas preservando la no objetivo (2590 ms), rechazar ruta inexistente sin side effects (2657 ms).
- **Evidencia runtime incorporada (2026-06-15, BR-002/BR-003)**:
  - `dysflow.import_modules` importó `Test_Adjuntos_Strict` en `CONDOR.accdb`.
  - `dysflow.compile_vba` devolvió `ok=true`.
  - `dysflow.test_vba` con `tests/tests.adjuntos.json` ejecutó 6/6 tests verdes: guardar/copia (2989 ms), eliminar fila+archivo (2799 ms), eliminar por etapas preservando la no objetivo (2449 ms), rechazar ruta inexistente sin side effects (2469 ms), rechazar no-PDF en `Documento Final Firmado` sin side effects (2641 ms), deduplicar destino existente (2719 ms).
- **Evidencia runtime incorporada (2026-06-15, BR-012 `ActualizarFicheroAdjunto`)**:
  - `dysflow.import_modules` importó `Test_Adjuntos_Strict` extendido con 3 átomos nuevos.
  - `dysflow.compile_vba` devolvió `ok=true`.
  - Cada átomo se ejecutó verde individualmente con Dysflow: reemplazar archivo preservando fila (3151 ms), rechazar `idAdjunto=0` con `Err 513` (4357 ms), rechazar id válido pero fila ausente con `Err 513 "Attachment not found"` (3901 ms). Los 6 átomos previos también siguen verdes en su corrida individual.
  - **Nota operativa**: el batch Dysflow de los 9 átomos en una sola llamada cae en `VBA_MANAGER_TIMEOUT` (~29 s) porque la suma de los `durationMs` (~28 s) está justo en el límite del runner. La evidencia sigue siendo válida — todos los átomos pasan individualmente — pero `dysflow.test_vba` con `tests/tests.adjuntos.json` debe invocarse en lotes de ≤ 6 átomos por llamada o subir el timeout del runner.
- **Evidencia runtime incorporada (2026-06-15, BR-004 sad paths de `GuardarAdjuntoDesdeArchivo`)**:
  - `dysflow.import_modules` importó `Test_Adjuntos_Strict` extendido con 2 átomos sad path.
  - `dysflow.compile_vba` devolvió `ok=true`.
  - Cada átomo se ejecutó verde individualmente con Dysflow: rechazar `idSolicitud=0` con `Err 513 "El ID de la Solicitud no es válido"` (3460 ms), rechazar `rutaArchivoOrigen=""` con `Err 513 "La ruta del archivo de origen no puede estar vacía"` (3967 ms). Sin side effects en `tbAdjuntos` (cardinalidad preservada) ni en el archivo origen.
  - Deuda explícita: regla 3 (≤255 chars) y reglas 4-5 (etapaWF/usuarioSubida/fechaSubida) de `Validar` siguen `Verified-static`; `Validar` es `Private Sub` y solo es accesible a través de las APIs públicas.
- **Evidencia no reclamada**: no cubre aún rollback por fallo de persistencia ni callbacks acopladas a `frmGestionAdjuntos`.
- **SDD/intención consultada**: SDD `document-template-mapping-e2e` (Cierre ↔ Word) y código actual.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| Guardar adjunto falla con error 513 sobre extensión | Etapa `Documento Final Firmado` y archivo no PDF. | prueba focal con stub de `fso.GetExtensionName`. | §2 BR-002 |
| Subir PDF de cierre no aprueba la solicitud | Callback atómica no se ejecuta o `WorkflowServicio.EjecutarTransicion` falla. | prueba con stub de `wfServ.EjecutarTransicion` y verificación de transición. | §2 BR-009 |
| Eliminar `Documento Final Firmado` no revierte a `estadoFormalizacion` | `WorkflowServicio.RevertirAFaseAnterior` no se invoca o falla. | prueba focal. | §2 BR-010 |
| La validación de adjunto permite nombre > 255 caracteres | `Validar` no se ejecuta por path. | prueba con `Len(nombreArchivo) = 256`. | §2 BR-004 |
| `getAdjuntosViewModelPorSolicitud` muestra IDs en lugar de nombres | `m_ObjEntorno.estados` no está cargado. | verificar carga de la cache. | §3 / §5 |
| Rollback de archivo no ocurre si falla CommitTrans | `ws.CommitTrans` falla antes de la compensación. | prueba con error simulado en `db.OpenRecordset`. | §2 BR-006 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Gestión de adjuntos y archivo | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Código sincronizado; pendiente manifest atómico y reconciliación de layout `.form.txt` de formularios de adjuntos. |

## §6 Notas de migración web

- **Conservar**: validación de adjunto por servicio, transacción atómica de persistencia, deduplicación de nombre, callback atómica de aprobación al subir PDF, sustitución con confirmación.
- **Transformar**: `m_ObjEntorno.URLDirectorioDocumentacion` a un servicio de almacenamiento de objetos (S3, Azure Blob, filesystem en servidor); `AbrirEnLocal` a un visor web con URL firmada; `fso` COM a `System.IO` server-side.
- **NO copiar**: transacción con `MS Access;PWD=…` en connection string, callback atómica acoplada a UI, dependencia de `Application.TempVars` para auto-carga, dependencia de `m_ObjUsuarioActivo` global.
- **Preguntas abiertas**: ¿La callback atómica de aprobación al subir PDF debe seguir en el cliente o pasar a un endpoint con contrato explícito? (responsable de producto). ¿La deduplicación por nombre con sufijo `_YYYYMMDD_HHMMSS` debe mantenerse o se prefiere `versionId` UUID? (equipo técnico).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| `AdjuntosServicio.GuardarAdjuntoDesdeArchivo` valida, copia con deduplicación, persiste `tbAdjuntos` y compensa archivo en disco si la persistencia falla. | Verified-static | `AdjuntosServicio.cls` líneas 18-130. | 2026-06-15 |
| `Validar` exige los cinco campos mínimos. | Verified-static | `AdjuntosServicio.cls` líneas 278-310. | 2026-06-15 |
| Para `Documento Final Firmado` la extensión debe ser `pdf`. | Verified-static | `AdjuntosServicio.cls` líneas 54-60. | 2026-06-15 |
| `EliminarAdjunto` borra archivo físico y registro en `tbAdjuntos` dentro de la misma transacción. | Verified-static | `AdjuntosServicio.cls` líneas 485-538. | 2026-06-15 |
| Al subir PDF de cierre se ejecuta callback atómica que aprueba la solicitud. | Verified-static / deuda UI | `frmGestionAdjuntos.cls` líneas 175-225. | 2026-06-15 |
| Al eliminar `Documento Final Firmado` en `estadoAprobada` se borra la decisión final y se revierte a `estadoFormalizacion`. | Verified-static / deuda UI | `frmGestionAdjuntos.cls` líneas 279-373. | 2026-06-15 |
| `getAdjuntosViewModelPorSolicitud` traduce `etapaWF` numérico a nombre legible. | Verified-static | `AdjuntosServicio.cls` líneas 356-413. | 2026-06-15 |
| `m_ObjEntorno.URLDirectorioDocumentacion` es la ruta física canónica. | Verified-static | `Variables Globales.bas` + `Entorno.cls`. | 2026-06-15 |
| `EliminarAdjuntosPorEtapas` admite `Variant` y compara con nombre o ID. | Verified-static | `AdjuntosServicio.cls` líneas 414-454. | 2026-06-15 |
| Existe un manifest atómico de pruebas de adjuntos que cumpla `access-vba-tdd` v2.4.2. | Verified-runtime parcial | `tests/tests.adjuntos.json` + `src/modules/Test_Adjuntos_Strict.bas`; `dysflow.test_vba` 11/11 verde por átomo individual el 2026-06-15 (batch único ≥ 9 cae en `VBA_MANAGER_TIMEOUT` por suma de `durationMs` ≈ 28 s vs límite 29 s del runner). | 2026-06-15 |
| Formularios de adjuntos están reconciliados fuente↔binario. | Divergent / blocker UI | `verify_binary`: `Form_frmGestionAdjuntos.form.txt` y `Form_frmElegirEtapaAdjunto.form.txt` con `bothChanged`; code-behind `.cls` matched. | 2026-06-15 |

**Divergencias pendientes de revisión humana**:

- BR-009: la callback atómica está en `frmGestionAdjuntos`. Extraer a un servicio `AdjuntosServicio.SubirYCerrar(idSolicitud, rutaPDF)` para que la prueba no dependa de UI.
- BR-010: la lógica de eliminación de `Documento Final Firmado` mezcla formulario y servicios. Extraer a un helper.
- BR-001: la validación de la etapa `Documento Final Firmado` y la compensación de archivo en disco son críticas; cualquier prueba focal debe ejercitar el rollback.
