# Capacidad: Ciclo de vida, búsqueda y navegación de solicitudes

## §0 Identidad

- **ID de capacidad**: CAP-006
- **Tier**: critical
- **Estado**: active con deuda de pruebas, seams UI y reconciliación layout
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary` sobre servicios/repositorios y formularios de alta/búsqueda/navegación. Code-behind `.cls` de los formularios y servicios principales está `matched`; `AplicacionRepositorio`/`SolicitudRepositorio` solo tienen diferencias no accionables (`caseOnly`/`whitespaceOnly`); 9 `.form.txt` de formularios tienen diferencias accionables `bothChanged`.
- **Confianza global**: mayoritariamente `Verified-static`. La búsqueda y alta están razonablemente implementadas y el code-behind está sincronizado, pero no hay manifest strict y el layout de formularios de navegación requiere reconciliación antes de declarar runtime UI.
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). La búsqueda/alta se apoyan en código que aún conserva patrones legacy (`SuiteSetup`, `SuiteTeardown`, `Nothing` como `DAO.Database`); la promoción a `Verified-runtime` exige la migración completa.

**Contrato TDD vigente**: las pruebas de búsqueda/alta deben migrarse a `access-vba-tdd` v2.4.2 — `Public Function` con retorno JSON canónico, fixture propio, schema-first, `DAO.Database` inyectado, cardinalidad y manifests atómicos.

**Justificación del nivel**: crítico, porque es el front-end de usuario para crear, buscar, filtrar y navegar solicitudes. Una regresión aquí bloquea cualquier caso de uso del producto.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: soportar el ciclo de vida de las solicitudes (alta, gestión, búsqueda, filtros avanzados, navegación a detalle, visor web) y servir de punto de entrada para los roles Técnico, Calidad y Administrador.
- **Usuarios / perfiles**: técnicos (tareas pendientes en `estadoDesarrolloTecnico`), calidad y administradores (filtros completos, "solo mis solicitudes", "solo pendientes"), auditores (visores de solo lectura).
- **Problema que resuelve**: estandariza la creación de solicitudes, la búsqueda por palabra clave, los filtros por responsables y tipo, y la navegación al detalle con visor web en HTML.
- **Valor de negocio**: reduce el tiempo de búsqueda de solicitudes, separa las tareas pendientes por rol, integra el visor web como herramienta de solo lectura y mantiene el contrato transaccional con la base de datos externa de No Conformidades.
- **No-objetivos**: este documento no cubre el detalle de los datos técnicos de cada tipo (PC, PCSUB, CD/CA, CDCASUB), que se documentan en CAP-003/CAP-004/CAP-005/CAP-001. Tampoco cubre el workflow puro, que vive en CAP-007.
- **Origen de la intención**: código actual + SDD `pcsub-nueva-solicitud` (gemelo en patrón), `modEnumeradores.bas` para enums de tipos, estados y bloques.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios principales

- **DADO** un usuario con rol `Calidad` o `Administrador` **CUANDO** abre `frmAltaSolicitud` y busca un expediente por Nemotécnico/CodExp **ENTONCES** el formulario propone código `DC-<claveExp>-<ordinal>`, ajusta el `RowSource` del tipo a `PC`/`CD_CA` (si `ContratistaPrincipal = "Sí"`) o `PC_SUB`/`CD_CA_SUB` (si no) y propone la NC vinculada por código CONDOR si existe. **Estado**: `Verified-static`; depende de `m_ObjEntorno.DirUTE` y de la cache de NCs.
- **DADO** una solicitud recién creada con NC externa **CUANDO** el usuario pulsa `cmdAltaSolicitud_Click` **ENTONCES** `SolicitudServicio.CrearNuevaSolicitud` + `Validar` + `Guardar` + `GuardarConVinculacionNC` se ejecutan en una única transacción local (`DBEngine.Workspaces(0)` + `ws.OpenDatabase(..., ";PWD=dpddpd")`). La transacción cubre `tbSolicitudes` y la base de datos externa de NCs. **Estado**: `Verified-static`.
- **DADO** una solicitud en edición con cambio de expediente y tipo **CUANDO** el usuario guarda **ENTONCES** se actualiza el registro de datos técnicos del tipo correspondiente (`DatosPCServicio.ActualizarCamposDependientesDeExpediente`, etc.) dentro de la misma transacción. **Estado**: `Verified-static`.
- **DADO** un usuario de Calidad o Técnico **CUANDO** abre `frmBuscarSolicitudes` con la marca "solo pendientes" **ENTONCES** la query `SolicitudServicio.getSolicitudesViewModel` filtra por estado: para Técnico deja solo `estadoDesarrolloTecnico`; para Calidad/Admin deja todo lo que no sea `estadoDesarrolloTecnico` ni `estadoAprobada`. **Estado**: `Verified-static`.
- **DADO** un usuario de Calidad **CUANDO** marca "solo mis solicitudes" **ENTONCES** la query filtra por `T_Exp.IDResponsableCalidad = m_ObjUsuarioActivo.id`; para Técnico, filtra por `EXISTS (SELECT 1 FROM TbExpedientesResponsables RT WHERE …)` con `m_ObjUsuarioActivo.id`. **Estado**: `Verified-static`.
- **DADO** un usuario con filtros avanzados activos (responsable técnico, responsable de calidad, suministrador, tipo de solicitud, estado) **CUANDO** pulsa "Buscar" **ENTONCES** `frmBuscarSolicitudes.RealizarBusqueda` invoca `getSolicitudesViewModel(palabraClave, m_FiltrosAvanzados, soloMis, soloPend)`. **Estado**: `Verified-static`.
- **DADO** un usuario **CUANDO** hace doble clic en una solicitud de la lista **ENTONCES** se abre `frmGestionSolicitud` con el `idSolicitud` como `OpenArgs`. **Estado**: `Verified-static`.
- **DADO** un usuario con visor web abierto **CUANDO** se refresca el timeline de la solicitud **ENTONCES** `WebVisorCacheServicio` sirve el HTML cacheado por 5 minutos; al transicionar un estado, `WorkflowServicio.InvalidarCacheTimeline` borra la entrada. **Estado**: `Verified-static`; requiere seam para verificar TTL en pruebas.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | El tipo de solicitud propuesto se filtra por `Expediente.ContratistaPrincipal`: principales → `PC`/`CD_CA`; subcontratistas → `PC_SUB`/`CD_CA_SUB`. | Código | Sí: `frmAltaSolicitud.CargarDatosExpediente` líneas 167-178. | Pendiente. | Verified-static |
| BR-002 | El código de solicitud propuesto se construye como `DC-<claveExp>-<siguienteOrdinal>`; el ordinal manual solo se aplica si es numérico y no colisiona con un código existente. | Código | Sí: `SolicitudServicio.ProponerCodigoSolicitud`; `frmAltaSolicitud.txtOrdinalManual_AfterUpdate`. | Pendiente. | Verified-static |
| BR-003 | El alta de solicitud persiste en `tbSolicitudes` y vincula la NC en la BD externa dentro de la misma transacción local con `";PWD=dpddpd"`. | Código + AGENTS | Sí: `frmAltaSolicitud.cmdAltaSolicitud_Click` (líneas 470-500) y `SolicitudServicio.GuardarConVinculacionNC`. | Pendiente; revisar NC cleanup post-commit. | Verified-static |
| BR-004 | El cambio de expediente en edición dispara `ActualizarCamposDependientesDeExpediente` para el tipo correspondiente, dentro de la misma transacción. | Código | Sí: `frmAltaSolicitud.cmdAltaSolicitud_Click` (líneas 534-549). | Pendiente. | Verified-static |
| BR-005 | La búsqueda de solicitudes excluye cualquier asignación de estado final cerrado que el rol no deba ver y combina filtros con `AND`. | Código | Sí: `SolicitudServicio.getSolicitudesViewModel`. | Pendiente. | Verified-static |
| BR-006 | El visor web sirve HTML cacheado por 5 minutos (`m_cacheTTL_Minutos = 5`) por `idSolicitud`; al transicionar estado, la entrada se invalida. | Código (Spec-118) | Sí: `WorkflowServicio.m_cacheTimeline`, `InvalidarCacheTimeline`, `WebVisorCacheServicio`. | Pendiente. | Verified-static |
| BR-007 | La búsqueda permite "solo pendientes" y "solo mis solicitudes" simultáneamente, combinados con filtros avanzados por `AND`. | Código | Sí: `SolicitudServicio.getSolicitudesViewModel` (líneas 109-158). | Pendiente. | Verified-static |
| BR-008 | La búsqueda devuelve una lista de `SolicitudBusquedaViewModel` con tipo, estado, código, nemotécnico, CodExp, responsable técnico y responsable de calidad. | Código | Sí: SQL en `getSolicitudesViewModel` líneas 84-107. | Pendiente. | Verified-static |
| BR-009 | La creación/edición dispara `SolicitudHaCambiado` y refresca el visualizador cuando aplica. | Código | Sí: `frmAltaSolicitud.RaiseEvent SolicitudCreadaConExito` y `SolicitudModificada`. | Pendiente. | Verified-static |
| BR-010 | La eliminación de solicitud (`SolicitudServicio.EliminarSolicitudCompleta`) barre TODAS las tablas de datos técnicos (`tbDatosPC`, `tbDatosCDCA`, `tbDatosCDCASUB`, etc.) dentro de la misma transacción. | Código | Sí: `SolicitudServicio.EliminarSolicitudCompleta` líneas 540-551. | Pendiente; revisar NC cleanup. | Verified-static |
| BR-011 | El cambio de NC asociado a una solicitud ya vinculada a la BD externa de NCs muestra un `MsgBox` de confirmación antes de sobrescribir. | Código | Sí: `frmAltaSolicitud.ValidarCambioNC` líneas 418-437. | Pendiente. | Verified-static |

### Validaciones observadas

- `idExpediente <= 0`, `tipoSolicitud` vacío o `codigoSolicitud` vacío en `SolicitudServicio.Validar` producen error 513.
- `revisionCalidadEstado` debe ser `PENDIENTE`, `APROBADO`, `RECHAZADO` o vacío.
- Duplicado de `codigoSolicitud` en alta o edición se rechaza con error 513.
- En edición, el código puede coincidir con el original de la propia solicitud (auto-match).

### Transiciones de estado y navegación

- `frmAltaSolicitud` no ejecuta transiciones por sí mismo, pero al guardar Datos Generales del tipo correspondiente invoca el servicio que puede transicionar Preregistro → Registro (ver CAP-004/CAP-005).
- `frmBuscarSolicitudes` no ejecuta transiciones: solo abre `frmGestionSolicitud` con el `idSolicitud` como `OpenArgs`.
- `frmGestionSolicitud` orquesta la gestión; la transición de estado se hace siempre por `WorkflowServicio` (CAP-007).

### Casos límite y hallazgos

- La query `getSolicitudesViewModel` con "solo pendientes" deja a `Tecnico` con un subconjunto limitado y a `Calidad/Admin` con casi todo. El comportamiento es el esperado según PRD, pero un futuro "pendientes" para Admin debería ser explícito.
- La eliminación de solicitud con NC externa hace un cleanup post-commit no bloqueante (ESC-131) que se loguea como `[WARN]` si falla. Si esa limpieza falla de forma consistente, la NC queda con `CodConcesionAsociada` apuntando a un código inexistente; no se rompe la integridad referencial, pero deja residuo.
- `SolicitudServicio.CalcularHashFaseModificacion` se apoya en `SnapshotServicio` para detectar cambios tras rechazo; este es el contrato base del rechazo/sub-sanación de Calidad (ver CAP-007).

### Señales de aceptación / presencia

- Existen `Solicitud.cls`, `SolicitudServicio.cls`, `SolicitudViewModel.cls`, `SolicitudBusquedaViewModel.cls`.
- Existen `Expediente.cls`, `ExpedienteServicio.cls`, `ExpedienteViewModel.cls`, `ExpedienteSuministrador.cls`.
- Existen los formularios `frm0Ppal`, `frm0PpalTecnico`, `frmAltaSolicitud`, `frmGestionSolicitud`, `frmBuscarSolicitudes`, `frmBuscarExpediente`, `frmDetalleExpediente`, `frmFiltrosAvanzadosSolicitudes`, `frmWebVisor`. Su code-behind está sincronizado, pero sus `.form.txt` tienen deriva accionable `bothChanged`.
- `tbSolicitudes` y `TbExpedientes` son las tablas núcleo. `tbSolicitudes.tipoSolicitud ∈ {PC, CD_CA, CD_CA_SUB, PC_SUB}` (también acepta `PCSUB` defensivo en `getNombreAmigableTipoSolicitud`).

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de UI**:
  - `Form_frm0Ppal.Form_Open` y `Form_frm0PpalTecnico.Form_Open` (shells de Calidad/Técnico y Técnico).
  - `Form_frmAltaSolicitud.Form_Load` (alta/edición) y `cmdAltaSolicitud_Click` (persistencia).
  - `Form_frmAltaSolicitud.cmdBuscarExpediente_Click`, `cmdIrABusquedaExpedientes_Click`, `cmdVerDetalleExpediente_Click`.
  - `Form_frmBuscarSolicitudes.Form_Load` y `RealizarBusqueda`; `cmdFiltrosAvanzados_Click` y `cmdQuitarFiltros_Click`; `lstResultados_DblClick` y `cmdVerDetalleSolicitud_Click`.
  - `Form_frmFiltrosAvanzadosSolicitudes` recoge responsable técnico, responsable de calidad, suministrador, tipo y estado; se entrega por `Tag = "OK"` y `filtros` como propiedad pública.
  - `Form_frmGestionSolicitud` orquesta la solicitud seleccionada, su workflow y sus pestañas técnicas.
  - `Form_frmWebVisor` consume el HTML cacheado de `WebVisorCacheServicio` y expone `colaComandos` para los botones.
- **Puntos de entrada de código**:
  - `SolicitudServicio.GuardarNuevaSolicitud`, `Validar`, `CrearNuevaSolicitud`, `Guardar`, `ProponerCodigoSolicitud`, `getSolicitudPorID`, `getSolicitudViewModelPorID`, `getSolicitudesViewModel`, `getHistorialDeEstados`, `EliminarSolicitudCompleta`, `CalcularHashFaseModificacion`, `GuardarConVinculacionNC`.
  - `ExpedienteServicio.getExpedientePorID`, `getExpedientes(palabraClave)`, `getNombresSuministradores`, `getSuministradoresPorExpediente`, `getClaveParaCodigoSolicitud`.
  - `SolicitudRepositorio.GuardarSolicitud`, `getSolicitudPorID`, `getSolicitudPorCodigo`, `ExisteCodigo`, `getSiguienteOrdinalParaClave`, `Eliminar`.
  - `ExpedienteRepositorio.getExpedientePorID` (sobre `getdbExpedientes()`).
  - `FiltrosSolicitud.HayFiltrosActivos` para señalizar visualmente.
  - `WebVisorCacheServicio` y `WorkflowServicio.InvalidarCacheTimeline`.
- **Datos afectados**:
  - `tbSolicitudes`: alta, edición, eliminación; `idEstadoInterno`, `tipoSolicitud`, `codigoSolicitud`, `idNCAsociada`, `revisionCalidad*`, `fecha*`, `usuario*`.
  - `TbExpedientes`: lectura para asignación de clave y nombres de suministradores; tablas vinculadas a `TbExpedientesResponsables` y `TbExpedientesSuministradores` para joins.
  - `TbUsuariosAplicaciones`/`TbUsuariosAplicacionesPermisos` (Lanzadera): para responsables técnico y de calidad, vía `getdbLanzadera()` y `getdbExpedientes()`.
- **Dependencias**:
  - `WorkflowServicio` para el visualizador HTML y para invalidar cache.
  - `DocumentoServicio` y `AdjuntosServicio` cuando la solicitud navega a una fase con adjunto.
  - `NoConformidadServicio` para la lista de NCs por expediente y la detección por código CONDOR.
  - `m_ObjEntorno`, `m_ObjUsuarioActivo`, `m_ObjUsuarioReal`, `rolUsuario`.
- **Sincronización fuente↔binario**: cambios de UI en `frmAltaSolicitud`, `frmBuscarSolicitudes`, `frmGestionSolicitud`, `frmFiltrosAvanzadosSolicitudes` y `frmWebVisor` requieren `import-form` para los `.form.txt` y `import-code` para los `.cls`. La cache de visor no se persiste; basta recompilar.
- **Valoración de diseño (tal-como-está vs ideal)**: la búsqueda de solicitudes es razonablemente legible y combina `wheres` con `AND` de forma segura; la asignación de tipo por `ContratistaPrincipal` está bien hecha y la transición Preregistro → Registro queda como contrato gemelo. La deuda principal está en (a) la query `getSolicitudesViewModel` con `getdb()` por defecto (debería aceptar `db` explícito para tests), (b) la cache del visor sin seam, (c) el `msgBox` de confirmación de NC en UI y (d) la duplicación del subquery de responsable técnico entre varios formularios. La pieza está bien hecha para la operación normal; no se recomienda promoción `Verified-runtime` sin migrar la suite.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `Solicitud.cls`, `SolicitudServicio.cls`, `SolicitudViewModel.cls`, `SolicitudBusquedaViewModel.cls` y `SolicitudRepositorio.bas`.
2. Restaurar `Expediente.cls`, `ExpedienteServicio.cls`, `ExpedienteViewModel.cls`, `ExpedienteSuministrador.cls` y `ExpedienteRepositorio.bas`.
3. Confirmar el esquema `tbSolicitudes` (10 columnas observables según ERD) y `TbExpedientes` (≈60 columnas).
4. Restaurar `FiltrosSolicitud.cls` y `AplicacionRepositorio.bas` (sobre `getdbLanzadera()`).
5. Restaurar `WebVisorCacheServicio.cls` y el cache de timeline en `WorkflowServicio.cls`.
6. Restaurar los formularios `frm0Ppal`, `frm0PpalTecnico`, `frmAltaSolicitud`, `frmBuscarSolicitudes`, `frmBuscarExpediente`, `frmDetalleExpediente`, `frmFiltrosAvanzadosSolicitudes`, `frmWebVisor` y `frmGestionSolicitud`.
7. Importar con `dysflow.import_modules` (UI primero, código después; nunca en paralelo) y compilar con `dysflow.compile_vba`.
8. Verificar binario con `dysflow.verify_binary` y resolver cualquier `bothChanged`; actualmente los 9 `.form.txt` de navegación/alta/búsqueda están en `bothChanged`.
9. Demostrar los escenarios de §2 con un manifest atómico `tests/tests.lifecycle.json`; mientras no exista, esta capacidad queda en `Verified-static`.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/SolicitudServicio.cls` (746 líneas).
  - `src/classes/Solicitud.cls`, `SolicitudViewModel.cls`, `SolicitudBusquedaViewModel.cls`.
  - `src/classes/ExpedienteServicio.cls`, `Expediente.cls`, `ExpedienteViewModel.cls`, `ExpedienteSuministrador.cls`.
  - `src/classes/FiltrosSolicitud.cls`.
  - `src/classes/WebVisorCacheServicio.cls`.
  - `src/modules/SolicitudRepositorio.bas`, `ExpedienteRepositorio.bas`, `AplicacionRepositorio.bas`.
  - `src/forms/Form_frm0Ppal.cls`, `Form_frm0PpalTecnico.cls`, `Form_frmAltaSolicitud.cls`, `Form_frmBuscarSolicitudes.cls`, `Form_frmBuscarExpediente.cls`, `Form_frmDetalleExpediente.cls`, `Form_frmFiltrosAvanzadosSolicitudes.cls`, `Form_frmWebVisor.cls`, `Form_frmGestionSolicitud.cls`.
  - `docs/ERD/condor_datos.md` → `tbSolicitudes`, `TbExpedientes`.
- **Evidencia Dysflow incorporada**:
  - `dysflow.verify_binary` de servicios/repositorios/formularios CAP-006: code-behind `.cls` matched; `AplicacionRepositorio`/`SolicitudRepositorio` solo no accionable; 9 formularios `.form.txt` accionables `bothChanged` (`frm0Ppal`, `frm0PpalTecnico`, `frmAltaSolicitud`, `frmBuscarSolicitudes`, `frmBuscarExpediente`, `frmDetalleExpediente`, `frmFiltrosAvanzadosSolicitudes`, `frmGestionSolicitud`, `frmWebVisor`).
- **Tests existentes**: no hay manifest atómico específico para búsqueda/alta; la cobertura histórica vive en `tests/tests.cdca.json` y `tests/tests.pcsub.json` (gemelos).
- **Evidencia no reclamada**: no se ejecutó `dysflow.test_vba` con un manifest de ciclo de vida; no se declara que ningún escenario de §2 esté en verde en este documento.
- **SDD/intención consultada**: SDD `pcsub-nueva-solicitud` (patrón), `modEnumeradores.bas` para enums.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación (Dysflow) | Ancla |
|---|---|---|---|
| La lista de búsqueda no devuelve resultados tras filtrar por tipo | El subquery `SubRT` (responsable técnico) hace `GROUP BY` con `FIRST`; verificar que la consulta SQL no se rompe en Access 2016+ con el `FIRST`. | prueba focal que invoque `getSolicitudesViewModel("DC-XXX", Nothing, False, False)`. | §2 BR-008 |
| El alta de solicitud con NC externa falla por error de credenciales | La cadena `";PWD=dpddpd"` está embebida en código (`frmAltaSolicitud.cls` línea 475) en lugar de leerse de `GetPasswordDB()`. | migrar a `GetPasswordDB()` y ejecutar prueba focal. | §2 BR-003 / §3 |
| El visor web queda obsoleto tras una transición | `WorkflowServicio.InvalidarCacheTimeline` no se invoca por algún path, o el cache no se limpia. | prueba focal que ejecute `EjecutarTransicion` y verifique que la cache cambia. | §2 BR-006 |
| La confirmación de NC vinculada no se muestra en edición | `frmAltaSolicitud.DetectarNCVinculadaEnBD` no se invoca tras seleccionar código. | auditoría del flujo de UI. | §2 BR-011 |
| La búsqueda con "solo pendientes" devuelve cero para Técnico | `estadoDesarrolloTecnico` no coincide con el ID canónico. | verificar `m_ObjEntorno.estados`. | §2 BR-007 / §3 |

### Trazabilidad de release

| Elemento | Ref. tracker | Versión de staging (UAT) | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Ciclo de vida, búsqueda y navegación | Pendiente | Pendiente de confirmación | pending | Pendiente | Pendiente | Code-behind sincronizado; pendiente manifest atómico y reconciliación de layout `.form.txt` de 9 formularios. |

## §6 Notas de migración web

- **Conservar**: combinación de filtros con `AND`, asignación de tipo por `ContratistaPrincipal`, transición Preregistro → Registro, cache de visor con invalidación por transición, integración con base de datos externa de NCs, `m_ObjUsuarioActivo` como usuario de sesión.
- **Transformar**: `NavigationSubform`/`BrowseTo`/`SendKeys` a rutas y async UI; `MsgBox` modal a diálogos de confirmación server-side; `DAO.Workspace` + password en `ConnectionString` a transacciones en servidor con ORM; la cache de visor web a `Redis`/`cache` con TTL explícito y versionado.
- **NO copiar**: dependencia de `m_ObjEntorno`/`m_ObjUsuarioActivo` globales, `getdb()`/`getdbLanzadera()`/`getdbExpedientes()`/`getdbNoConformidades()` como singletons, contraseña `dpddpd` embebida en `frmAltaSolicitud`, `First()` (Access-only) en la query de responsable técnico, doble cache de visor con reglas ad-hoc.
- **Preguntas abiertas**: ¿La transición Preregistro → Registro debe seguir siendo automática al guardar Datos Generales o debe ser manual? (responsable de producto). ¿La cache del visor web debe sobrevivir reinicio del servidor o basta con 5 minutos? (equipo técnico). ¿La búsqueda debe permitir "solo pendientes" como filtro explícito o se mantiene la convención por rol? (responsable de producto).

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| El alta de solicitud persiste en `tbSolicitudes` y vincula la NC externa en la misma transacción. | Verified-static | `frmAltaSolicitud.cls` líneas 470-500; `SolicitudServicio.GuardarConVinculacionNC`. | 2026-06-15 |
| La búsqueda de solicitudes combina filtros con `AND` y admite "solo pendientes" y "solo mis solicitudes". | Verified-static | `SolicitudServicio.getSolicitudesViewModel` líneas 67-172. | 2026-06-15 |
| El visualizador web sirve HTML cacheado por 5 minutos y se invalida al transicionar. | Verified-static | `WorkflowServicio.m_cacheTimeline`, `InvalidarCacheTimeline`; `WebVisorCacheServicio`. | 2026-06-15 |
| La contraseña `dpddpd` está embebida en `frmAltaSolicitud` en lugar de leerse de `GetPasswordDB()`. | Verified-static / deuda | `frmAltaSolicitud.cls` línea 475. | 2026-06-15 |
| La query `getSolicitudesViewModel` con "solo pendientes" deja al Técnico solo en `estadoDesarrolloTecnico`. | Verified-static | `SolicitudServicio.getSolicitudesViewModel` líneas 109-122. | 2026-06-15 |
| La eliminación de solicitud barre todas las tablas de datos técnicos en una transacción. | Verified-static | `SolicitudServicio.EliminarSolicitudCompleta` líneas 540-551. | 2026-06-15 |
| El cambio de expediente en edición dispara `ActualizarCamposDependientesDeExpediente` para el tipo. | Verified-static | `frmAltaSolicitud.cls` líneas 534-549. | 2026-06-15 |
| El subquery del responsable técnico usa `FIRST` (Access-only) y puede no ser portable. | Verified-static / riesgo de portabilidad | `SolicitudServicio.getSolicitudesViewModel` línea 86-89. | 2026-06-15 |
| El NC cleanup post-commit a la BD externa es no bloqueante y deja `[WARN]` si falla. | Verified-static / aceptable | `SolicitudServicio.EliminarSolicitudCompleta` líneas 564-587; fix ESC-131. | 2026-06-15 |
| Existe un manifest atómico de pruebas de ciclo de vida que cumpla `access-vba-tdd` v2.4.2. | Divergent / pendiente | `tests/` sin `tests.lifecycle.json`. | 2026-06-15 |
| Los formularios de alta/búsqueda/navegación están reconciliados fuente↔binario. | Divergent / blocker UI | `verify_binary`: 9 `.form.txt` con `bothChanged`; code-behind `.cls` matched. | 2026-06-15 |

**Divergencias pendientes de revisión humana**:

- BR-003: la contraseña `dpddpd` está embebida en `frmAltaSolicitud.cls` línea 475. Migrar a `GetPasswordDB()` y a un patrón coherente con el resto del proyecto.
- BR-008: el `FIRST` de Access en la query de responsable técnico limita la portabilidad. Considerar `MIN`/`MAX` con `TOP 1` ordenado, o calcular el responsable en una capa de servicio.
- BR-005/006: la cache del visor y la lógica de "solo pendientes" están razonablemente implementadas, pero cualquier cambio de contrato con Calidad/Admin debe revisarse con producto.
