# 05 · Capacidades — Índice

## Propósito

Índice maestro de las capacidades de negocio del ecosistema legacy. Sigue `access-vba-capability-docs`: cada capacidad se describe como contrato de comportamiento, con datos, dependencias, integraciones, excepciones, confianza y trazabilidad.

Este índice resume y enlaza los estudios de capacidades ya entregados, uno por aplicación, en `03-aplicaciones/<app>/prd/capacidades.md`. La ficha, la evidencia con cita `ruta:línea`, el commit fijado y los hallazgos viven en el PRD de cada aplicación; aquí no se reproducen. El índice fija la regla de tier, el estado de cada estudio y la fila por capacidad atómica.

## Convenciones

- **Tier** de cada capacidad: `critical` | `standard` | `minimal`. La regla que asigna cada valor está en [Regla de tier](#regla-de-tier).
- **Estado**: el valor que usa el PRD de cada aplicación frente al snapshot o al estudio previo. Las aplicaciones no comparten vocabulario; cada tabla declara el suyo.
- **Objetos de lógica**: los objetos que el PRD declara para esa capacidad, transcritos de su columna de objetos. El fichero PRD, el commit fijado y el rango de identificadores se declaran una sola vez por aplicación, en la cabecera de su tabla, no en cada fila. Si el PRD no declara objetos, la celda va con `—`.
- **Dependencia principal**: la dependencia que el PRD sostiene. Si el PRD no declara ninguna, la celda va con `—`.
- **Excepción**: la desviación, la ausencia o la disposición que el PRD registra. Si no hay ninguna, la celda va con `—`.
- **Confianza** por hecho: `Verified-runtime` / `Verified-static` / `Intended` / `Likely` / `Divergent`, con fecha y evidencia.
- **Trazabilidad**: enlace a GH issue/PR cuando exista; nunca como fuente de verdad.
- **Plantilla**: `access-vba-capability-docs/assets/capability-doc-template.md`.
- **Listado de capacidades** (formato del índice): `id · capacidad · tier · estado · objetos de lógica · dependencia principal · excepción`.

## Regla de tier

El tier se asigna leyendo la descripción de la capacidad en el PRD de su aplicación. No se asigna por palabra clave.

- **`critical`**: sostiene la identidad y la autorización (login, roles, permisos, auditoría de accesos) o el ciclo de vida del expediente y su trazabilidad (alta, edición, baja, estado, jerarquía, responsables, cambio de tipo, E2E y trazabilidad de exportación).
- **`standard`**: dominio de negocio propio con flujo y datos (solicitudes, riesgos, no conformidades, eventos, facturación, calibración, planificación, contratos de integración con otras aplicaciones).
- **`minimal`**: consulta, informe, exportación, contenido formativo y operaciones de apoyo sin estado propio.

Dos precisiones de aplicación:

- Una capacidad que cumple más de una regla toma el tier de la más específica: identidad, autorización y ciclo de vida fijan `critical` aunque la capacidad también sirva para consultar.
- Las filas que el PRD clasifica como «mecanismo legacy a retirar» no son capacidades de negocio: no llevan tier y su celda va con `—`.

## Estado por aplicación

| Aplicación | Estudio | Commit fijado | Fase 5 | Enlace |
|---|---|---|---|---|
| Expedientes | #1073 | `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` | cerrada | — |
| Lanzadera | #1074 | `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` | cerrada | — |
| Condor | #1075 | `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` | cerrada | — |
| Brass | #1078 | `3f77d586814096c04f24bd2b6c229831cfbfcd34` | cerrada | — |
| Gestión de Riesgos | #1079 | `36fcf60161fa5ec84c3956b1308537ddf2579256` | cerrada | — |
| HPS | #1076 | `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` | en curso | [#1076](https://github.com/DysTelefonica/access2web-blueprint/issues/1076) |
| HPS Solicitudes | #1077 | `cfa1937ab05a9e40150ea9236243898de96f84da` | cerrada | [#1077](https://github.com/DysTelefonica/access2web-blueprint/issues/1077) |
| No Conformidades | #1080 | `1d7a2f826a1e3c582874e4e29886a6183efb9067` | cerrada | [#1080](https://github.com/DysTelefonica/access2web-blueprint/issues/1080) |

Las siete filas cerradas tienen su fase 5 entregada en `main`: el fichero `prd/capacidades.md` de cada aplicación. HPS (#1076) sigue en curso y no tiene todavía fichero de capacidades, por lo que este índice no le asigna filas.

El README del PRD de Gestión de Riesgos conserva el estado de su fase 1; su fase 5 es el fichero `prd/capacidades.md`, presente en `main`.

## Reglas del índice

- Una fila por capacidad de negocio, no por formulario o módulo.
- La evidencia de cada capacidad vive en el PRD de su aplicación; el índice resume y enlaza, no reescribe.
- Las filas se enlazan desde `03-aplicaciones/<app>/` y desde la matriz de dependencias.
- APAP y APAP_WEB no aparecen en el índice ni en las fichas.

## Índice maestro

### Expedientes

Fuente: [`prd/capacidades.md`](../03-aplicaciones/expedientes/prd/capacidades.md) · commit fijado `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · identificadores `EXP-CAP-001..063` · vocabulario de estado: `sin cambios`, `modificada`, `añadida`, `retirada`, `solo formulario`, `sin equivalente comprobado`, `mecanismo legacy a retirar`.

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| EXP-CAP-001 | Alta transaccional del expediente, incluida el alta iniciada por sistemas autorizados. | critical | modificada | `Expediente.cls`, `ExpedienteOperaciones.cls`; retirado `Helper_ExpedienteAlta.bas` | expediente | — |
| EXP-CAP-002 | Edición del agregado con invariantes y concurrencia explícitas. | critical | modificada | `Expediente.cls`, `ExpedienteOperaciones.cls`; retirado `Helper_ExpedienteEdicion.bas` | expediente | — |
| EXP-CAP-003 | Baja condicionada, auditable y sin pérdida accidental de relacionados. | critical | modificada | `ExpedienteOperaciones.cls`; retirado `Helper_ExpedienteEliminacion.bas` | expediente | — |
| EXP-CAP-004 | Cambio de tipo con validación de elegibilidad y efectos derivados. | critical | modificada | `ExpedienteOperaciones.cls`, `ExpedienteCompleto.cls` | expediente | — |
| EXP-CAP-005 | Estado calculado y transiciones con fechas, garantía y motivos. | critical | modificada | `Expediente.cls`, `FUNCIONES UTILES.bas`; retirado `Helper_ExpedienteFechasLogic.bas` | expediente | — |
| EXP-CAP-006 | Jerarquía de acuerdos marco, lotes y basados, incluida integridad padre-hijo. | critical | modificada | `Expediente.cls` | expediente | — |
| EXP-CAP-007 | Derivación ordinal funcional dentro de la jerarquía. | critical | modificada | `Expediente.cls` | expediente | — |
| EXP-CAP-008 | Hitos y sus reglas temporales. | critical | modificada | `ExpedienteOperaciones.cls` | expediente | — |
| EXP-CAP-009 | Modificados, cambios de campo e historial reciente. | critical | sin cambios | — | expediente | — |
| EXP-CAP-010 | Anexos y referencias documentales con eliminación y retención controladas. | standard | modificada | `ExpedienteOperaciones.cls`; retirado `Helper_ExpedienteAnexos.bas` | expediente | — |
| EXP-CAP-011 | Anualidades vinculadas al expediente. | standard | modificada | `ExpedienteOperaciones.cls` | expediente | — |
| EXP-CAP-012 | Responsables por rol y jefaturas. | critical | modificada | `ExpedienteResponsable.cls`; añadidos `modResponsablesEdit.bas`, `clsResponsablesEditState.cls`; retirado `modResponsablePorRolService.bas` | expediente | — |
| EXP-CAP-013 | Relaciones con entidades, jurídicas y cadena de entidades. | standard | modificada | `ExpedienteEntidad.cls`, `ExpedienteEntidadOperaciones.cls`; retirados `Helper_EntidadCRUD.bas`, `Helper_ExpedienteEntidadesTabla.cls` | expediente | — |
| EXP-CAP-014 | Suministradores, contratistas, subcontratistas, UTE y jerarquía «el árbol manda». | standard | modificada | `ExpedienteSuministradorServicio.cls`, `ExpedienteSuministradorRepositorio.bas`; retirado `Helper_ExpedienteSuministradoresArbol.cls` | expediente | — |
| EXP-CAP-015 | Comerciales. | standard | modificada | `ComercialOperaciones.cls` | — | — |
| EXP-CAP-016 | CPV. | standard | sin cambios | — | — | — |
| EXP-CAP-017 | Ejércitos. | standard | modificada | `EjercitoOperaciones.cls` | — | — |
| EXP-CAP-018 | Suministradores. | standard | modificada | `SuministradorOperaciones.cls` | — | — |
| EXP-CAP-019 | Lugares de ejecución. | standard | modificada | `LugarEjecucionOperaciones.cls` | — | — |
| EXP-CAP-020 | PECAL. | standard | modificada | `PECALOperaciones.cls` | — | — |
| EXP-CAP-021 | RAC. | standard | modificada | `RACOperaciones.cls` | — | — |
| EXP-CAP-022 | Grados de clasificación. | standard | modificada | `GradoClasificacionOperaciones.cls` | — | — |
| EXP-CAP-023 | Órganos de contratación. | standard | modificada | `OrganoContratacionOperaciones.cls` | — | — |
| EXP-CAP-024 | Oficinas de programa. | standard | modificada | `OficinaProgramaOperaciones.cls` | — | — |
| EXP-CAP-025 | Bandeja paginada con filtros y detalle de solo lectura. | minimal | modificada | `ExpedienteCompleto.cls`, `Entorno.cls`; retirado `Helper_ExpedienteConsultas.bas` | expediente | — |
| EXP-CAP-026 | Búsqueda avanzada con semántica de filtros preservada. | minimal | modificada | `ExpedienteBusqueda.cls`; añadido `modExpedienteSearch.bas` | expediente | — |
| EXP-CAP-027 | Búsqueda técnica sujeta a permiso efectivo. | minimal | sin cambios | — | expediente | — |
| EXP-CAP-028 | Exportación Excel reproducible y autorizada. | minimal | modificada | añadido `modInformeExpedientesHelper.bas`; retirado `Helper_ExpedientesExportExcel.bas` | expediente | — |
| EXP-CAP-029 | Tareas calculadas por reglas del expediente, sin inventar una tabla de tareas. | minimal | modificada | `Entorno.cls`, `FUNCIONES UTILES.bas` | expediente | — |
| EXP-CAP-030 | Autosave de datos generales y fechas con transacción explícita. | standard | retirada | retirado `GuardadoAutomaticoHelper.bas` | expediente | — |
| EXP-CAP-031 | Autosave de hitos, modificados y restantes secciones editables. | standard | retirada | retirado `GuardadoAutomaticoHelper.bas` | expediente | — |
| EXP-CAP-032 | Idempotencia, prevención de doble envío y feedback accesible de progreso/error. | standard | sin equivalente comprobado | — | — | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-033 | DTO de dominio estable para intercambio. | standard | sin cambios | — | — | — |
| EXP-CAP-034 | Exportación JSON completa y determinista. | critical | modificada | `ExpedienteJsonExporter.bas`, `JsonConverter.bas`; retirado `ExpedienteE2EUtils.bas` | expediente | — |
| EXP-CAP-035 | Gestión y ejecución batch E2E. | critical | modificada | añadido `ManualE2EExport.bas`; retirados `E2EExportService.bas`, `E2EManagementService.bas` | expediente | — |
| EXP-CAP-036 | Hash de contenido con algoritmo y canonicalización versionados. | critical | modificada | `E2EHashService.bas` | expediente | — |
| EXP-CAP-037 | Generación del paquete de exportación E2E. | critical | modificada | añadido `ManualE2EExport.bas`; retirado `E2EExportService.bas` | expediente | — |
| EXP-CAP-038 | Trazabilidad entre expediente, batch, artefacto y resultado. | critical | añadida | añadidos `E2EColaHtmlSnapshot.bas`, `modE2EColaPanelHelper.bas`, `modNotifyChangeE2E.bas` | expediente | — |
| EXP-CAP-039 | Selección manual de expedientes para batch. | critical | añadida | añadidos `FormE2EGestionEnlacesLogic.bas`, `E2EEnlacesService.bas` | expediente | — |
| EXP-CAP-040 | Destino de exportación por usuario, sin rutas locales como contrato. | critical | sin equivalente comprobado | — | expediente | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-041 | Sesión de exportación, reanudación y cierre. | critical | sin equivalente comprobado | — | expediente | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-042 | Ordinal E2E único y consistente en dominio y trazabilidad. | critical | modificada | `ExpedienteEntidadOperaciones.cls`; añadido `M_E2EMigration.bas` | expediente | — |
| EXP-CAP-043 | Autorización deny-by-default por capacidad. | critical | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-044 | Login y principal actual consumidos desde Lanzadera. | critical | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-045 | Estado de usuario y sesión sin globals mutables. | critical | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-046 | Auditoría de accesos, cambios, exportaciones e integraciones. | critical | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-047 | Arranque y readiness con fallos diagnosticables. | minimal | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-048 | Configuración tipada por entorno e inyección de dependencias. | minimal | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-049 | Caché invalidable y read-models reconstruibles. | minimal | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-050 | Binding de backend por despliegue, sin selector interactivo ni rutas embebidas. | minimal | sin equivalente comprobado | — | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-051 | Alta/consulta desde HPS mediante contrato autenticado e idempotente. | critical | solo formulario | — | HPS | — |
| EXP-CAP-052 | Intercambio con AGEDYS con ownership y errores definidos. | standard | sin cambios | — | AGEDYS | — |
| EXP-CAP-053 | Enlace con Gestión de Riesgos por identificador estable. | standard | sin equivalente comprobado | — | Gestión de Riesgos | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-054 | Enlace con No Conformidades y su estado asociado. | standard | sin equivalente comprobado | — | No Conformidades | concepto de la web; sin objeto equivalente identificado en RC1 |
| EXP-CAP-055 | Notificaciones por correo mediante puerto y trazabilidad. | standard | modificada | `Correo.cls` | — | — |
| EXP-CAP-056 | Documentos en SharePoint u object storage mediante referencia opaca, retención y autorización. | standard | sin cambios | — | SharePoint u object storage | — |
| EXP-CAP-057 | Resolución de red Win32. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-058 | Gestión de procesos Win32. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-059 | Drag & drop OLE y controles ActiveX. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-060 | Globals y singletons mutables de sesión/entorno. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-061 | Cambio interactivo de backend y rutas embebidas. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-062 | Popup modal anti-spam/busy. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |
| EXP-CAP-063 | Menú JSON-hub y navegación acoplada a formularios. | — | mecanismo legacy a retirar | — | — | no es capacidad de negocio |

Disposición del ledger normativo ([`capability-ledger.md`](../03-aplicaciones/expedientes/capability-ledger.md)):

| Disposición | Capacidades |
|---|---|
| PRESERVAR | EXP-CAP-001..009, 011..013, 015..024, 026..028, 033..034, 036..039, 041..042, 052..054 |
| MODERNIZAR | EXP-CAP-010, 014, 025, 029..032, 035, 040, 043..051, 055..056 |
| RETIRAR | EXP-CAP-057..063 |

### Lanzadera

Fuente: [`prd/capacidades.md`](../03-aplicaciones/lanzadera/prd/capacidades.md) · commit fijado `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · identificadores `C-LAN-1..15` · vocabulario de estado: `sin cambios`, `modificada`, `añadida`, `retirada`.

Las 15 capacidades salen «sin cambios en el código»: los 59 ficheros de código coinciden con el snapshot. La disposición aprobada es una decisión de producto y va en la columna de excepción.

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| C-LAN-1 | Autenticación y sesión | critical | sin cambios | `Funciones Genreales.bas` (`Login`), `Variables Globales.bas` (`LoginCorrecto`), `Constructor.bas`, `Usuario.cls` | identidad de Lanzadera | — 
| C-LAN-2 | Alta, baja y perfil de usuario | critical | sin cambios | `Usuario.cls` | identidad de Lanzadera | — 
| C-LAN-3 | Restablecimiento y cambio de contraseña | critical | sin cambios | `Usuario.cls` (`ResetearPass`, histórico) | identidad de Lanzadera | — 
| C-LAN-4 | Catálogo de aplicaciones | standard | sin cambios | `Aplicacion.cls`, `Constructor.bas` | — | — 
| C-LAN-5 | Visibilidad por ubicación y estado | minimal | sin cambios | `Aplicacion.cls`, `Funciones Genreales.bas` | — | disposición aprobada: retirar la segmentación por ubicación 
| C-LAN-6 | Lanzamiento y paso de identidad | critical | sin cambios | `Funciones Genreales.bas` (`Lanzar`), `Aplicacion.cls` | identidad de Lanzadera | disposición aprobada: retirar el lanzador de escritorio 
| C-LAN-7 | Asignación usuario-aplicación | critical | sin cambios | `Usuario.cls`, `UsuarioAplicacionPermisos.cls` | permisos de Lanzadera | — 
| C-LAN-8 | Roles y permisos granulares | critical | sin cambios | `UsuarioAplicacionPermisos.cls` | permisos de Lanzadera | — 
| C-LAN-9 | Administrador global | critical | sin cambios | `Usuario.cls` (`EsAdministradorCalculado`) | permisos de Lanzadera | — 
| C-LAN-10 | Menús por perfil | minimal | sin cambios | `Menu.cls` | permisos de Lanzadera | disposición aprobada: sustituir por menú global con permisos 
| C-LAN-11 | Contenido formativo | minimal | sin cambios | `Video.cls`, `Categoria.cls`, `Visionado.cls`, `Funciones Genreales.bas` | — | disposición aprobada: retirar; el histórico queda archivado y no se migra 
| C-LAN-12 | Conexiones y auditoría | critical | sin cambios | `Conexion.cls`, `AplicacionApertura.cls` | identidad de Lanzadera | disposición aprobada: retirar la telemetría de SSID y ubicación 
| C-LAN-13 | Cola de correo | standard | sin cambios | `Correo.cls` | — | disposición aprobada: reemplazar el mecanismo por un servicio de notificaciones 
| C-LAN-14 | Configuración de backend y entorno | minimal | sin cambios | `Variables Globales.bas` (`getdb`, `LeeConfiguracionLocal`), `Config_BackendHelper.bas` | — | disposición aprobada: retirar como capacidad de administración 
| C-LAN-15 | Tareas, batch y exportaciones | minimal | sin cambios | campos `ParaTareasProgramadas`, `DiaEnvioTareas` en `Usuario.cls` | — | disposición aprobada: fusionar con el scheduler compartido 

### Condor

Fuente: [`prd/capacidades.md`](../03-aplicaciones/condor/prd/capacidades.md) · commit fijado `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · identificadores `C-CON-1..22` · vocabulario de estado: `sin cambios`, `modificada`, `retirada` (parcial, en `C-CON-7`).

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| C-CON-1 | Arranque, identidad y validación crítica | critical | sin cambios | `Variables Globales.bas` (`EVE`, `EVE.ValidacionCritica`) | identidad de Lanzadera | — 
| C-CON-2 | Roles y jerarquía | critical | modificada | `modEnumeradores.bas` (`Enum rol`), `UsuarioServicio.DeterminarRol` | permisos de Lanzadera | `JerarquiaRolesHelper.bas` no existe en el commit 
| C-CON-3 | Configuración de backend | minimal | retirada | `Variables Globales.bas` (`TbConfiguracionBackends`) | — | la configuración de conexión vive en el código (`FU:233-252`) 
| C-CON-4 | Modo de pruebas y sandbox seguro | minimal | modificada | `FUNCIONES UTILES.bas` (`m_TestingMode`, `m_BackendSandboxURL`) | — | el aislamiento pasa a ser copia de ficheros, no enrutado de `getdb()` 
| C-CON-5 | Login por comando de Lanzadera | critical | sin cambios | `Variables Globales.bas` (`m_Command`) | identidad de Lanzadera | — 
| C-CON-6 | Gestión de solicitudes | standard | sin cambios | `SolicitudServicio.cls`, `Form_frmGestionSolicitud.cls` | — | — 
| C-CON-7 | Cuatro tipos de solicitud | standard | retirada (una de las cuatro) | `DatosPC`, `DatosCDCA`, `DatosCDCASUB`, `DatosPCSUB` y sus servicios | — | el tipo `PC_SUB` no existe en el commit; quedan tres tipos 
| C-CON-8 | Catálogo de estados y transiciones | standard | sin cambios | `modEnumeradores.bas` (`enumEstados`), `WorkflowRepositorio.bas`, `tbTransiciones` | — | — 
| C-CON-9 | Validación de calidad | standard | sin cambios | `Solicitud.cls` (`revisionCalidadEstado`, `revisionCalidadComentarios`), `ValidacionRevisionServicio.cls` | — | — 
| C-CON-10 | Historial de rechazos | standard | modificada | `tbHistorialRechazos`, `Rechazo.cls`, `RechazoServicio.cls` | — | `tbHistorialRechazos` no existe en el volcado vigente 
| C-CON-11 | Adjuntos por etapa | standard | sin cambios | `AdjuntosServicio.cls`, `AdjuntoRepositorio.bas` | — | — 
| C-CON-12 | Visor de estado embebido | minimal | sin cambios | `Form_frmGestionSolicitud.cls` (`webInfo`), `WebVisorCacheServicio.cls`, `WorkflowServicio.GenerarHTML_VisualizadorDeEstado` | — | — 
| C-CON-13 | Motor de workflow | standard | sin cambios | `WorkflowServicio.cls` | — | — 
| C-CON-14 | Vinculación con expediente | standard | sin cambios | `Expediente.cls`, `ExpedienteServicio.cls` | Expedientes | — 
| C-CON-15 | Vinculación con no conformidades | standard | sin cambios | `NoConformidad.cls`, `NoConformidadServicio.cls` | No Conformidades | — 
| C-CON-16 | Suministradores | standard | sin cambios | `Suministrador.cls`, `SuministradorServicio.cls` | — | — 
| C-CON-17 | Documentos y PDF | standard | sin cambios | `DocumentoServicio.cls` | — | — 
| C-CON-18 | Notificaciones | standard | modificada | `NotificacionServicio.cls`, `MockNotifServ.cls` | — | `MockNotifServ` no existe; la simulación envía notificaciones reales 
| C-CON-19 | Auditoría en tres tablas | critical | sin cambios | `LogServicio.cls`, `CondorError.cls` | — | — 
| C-CON-20 | Mapeo de campos | standard | sin cambios | `MapeoServicio.cls`, `MapeoRepositorio.bas` | — | — 
| C-CON-21 | Hashing y snapshot | standard | sin cambios | `HashHelper.bas`, `SnapshotHelper.cls` | — | — 
| C-CON-22 | Batería de pruebas | minimal | sin cambios | `modBattery_Canonical.bas` | — | — 

### Brass

Fuente: [`prd/capacidades.md`](../03-aplicaciones/brass/prd/capacidades.md) · commit fijado `3f77d586814096c04f24bd2b6c229831cfbfcd34` · identificadores `CAP-1..12` (provisionales en el PRD) · vocabulario de estado: `sin cambios`, `modificada`.

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| CAP-1 | Arranque, identidad y permisos | critical | modificada | `Variables Globales.bas`, `Constructor.bas`, `Usuario.cls`, `UsuarioAplicacionPermisos.cls`, `Entorno.cls` | identidad de Lanzadera | sin formulario de login propio 
| CAP-2 | Opciones y configuración de la BD | minimal | sin cambios | `Funciones Generales.bas`, `Entorno.cls` | — | — 
| CAP-3 | Eventos (alta, edición, copia, gestión, SLA, franqueo, alta masiva) | standard | sin cambios | `Evento.cls`, `EventoBusqueda.cls`, `Constructor.bas`, módulos SLA | — | — 
| CAP-4 | Actividades y tipos de actividad | standard | sin cambios | `Actividad.cls` | — | — 
| CAP-5 | Equipos, equipos de medida y calibraciones | standard | sin cambios | `Equipo.cls`, `EquipoMedida.cls`, `EquipoMedidaCalibracion.cls`, `EventoEquipoMedida.cls`, `EventoEquipoMedidaAntes.cls` | — | — 
| CAP-6 | Planificación y partes | standard | sin cambios | `Parte.cls`, `Constructor.bas` (no existe `Planificacion.cls`) | — | — 
| CAP-7 | Materiales, reparaciones y repuestos | standard | sin cambios | `Material.cls`, `MaterialSeguimiento.cls` | — | — 
| CAP-8 | Facturación, gastos y subcontratación | standard | sin cambios | `Factura.cls`, `FacturaPerfiles.cls`, `Gasto.cls`, `Subcontratacion.cls`, `LIbranza.cls` | — | — 
| CAP-9 | Técnicos, ausencias, fiestas y precios | standard | sin cambios | `Tecnico.cls`, `TecnicoHoras.cls`, `TipoTecnico.cls`, `TipoTecnicoPrecio.cls`, `TipoDia.cls`, `Fiesta.cls` | — | — 
| CAP-10 | Informes (SLA, RAC, desde lista, por factura) | minimal | sin cambios | `Informe.cls`, `GestorInforme.cls`, `SLAReportService.bas`, `SLAHTMLService.bas`, `SLAValidator.bas` | — | — 
| CAP-11 | Anexos | standard | sin cambios | `Anexo.cls` | — | — 
| CAP-12 | Estructura y catálogos (BUI, nodos, subsistemas, ubicación, originador) | standard | sin cambios | `Constructor.bas`, `Funciones Generales.bas` | — | — 

### Gestión de Riesgos

Fuente: [`prd/capacidades.md`](../03-aplicaciones/gestion-riesgos/prd/capacidades.md) · commit fijado `36fcf60161fa5ec84c3956b1308537ddf2579256` · el estudio no numera las capacidades, por lo que la columna `id` va con `—` · vocabulario de estado: `sin cambios`, `modificada`, `añadida`, `pendiente de ficha`.

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| — | Arranque e identidad | critical | sin cambios | `EVE` (`Variables Globales.bas:204`), `Usuario.cls`, `Entorno.cls` | identidad de Lanzadera | — 
| — | Roles (administrador, calidad, técnico) | critical | modificada | `Usuario.cls`, `UsuarioAplicacionPermisos.cls` | identidad de Lanzadera | la identidad se resuelve leyendo el backend de Lanzadera, no un servicio 
| — | Configuración temporal | minimal | modificada | `Variables Globales.bas` | — | el modelo conmutable se fuerza a «nuevo» en el arranque 
| — | Proyecto | standard | sin cambios | `Proyecto.cls`, `ProyectoSuministrador.cls` | — | — 
| — | Edición | standard | sin cambios | `Edicion.cls` | — | — 
| — | Riesgo (alta, edición, borrado) | standard | sin cambios | `Riesgo.cls` | — | — 
| — | Estados del riesgo (14 estados) | standard | pendiente de ficha | `Riesgo.cls`, `RiesgosGestionRiesgo` | — | — 
| — | Causa raíz y biblioteca | standard | sin cambios | `RiesgoBiblioteca.cls` | — | — 
| — | Riesgo externo | standard | sin cambios | `RiesgoExterno.cls` | — | — 
| — | Materialización | standard | modificada | `RiesgoMaterializacion.cls` | — | el formulario de alta no está en `main` 
| — | Planes de mitigación (PM) | standard | sin cambios | `PM.cls`, `PMAccion.cls`, `PMAccionReversa.cls` | — | — 
| — | Planes de contingencia (PC) | standard | sin cambios | `PC.cls`, `PCAccion.cls`, `PCAccionReversa.cls` | — | — 
| — | Evidencias y anexos | standard | sin cambios | `Anexo.cls`, `AnexoAntiguo.cls`, `EdicionSuministrador.cls` | — | — 
| — | Suministradores por edición | standard | sin cambios | `EdicionSuministrador.cls`, `Suministrador.cls` | — | — 
| — | Control de cambios | standard | sin cambios | `CCVersion.cls`, `CCCambio.cls`, `CCDocumentoCambio.cls` | — | — 
| — | Publicabilidad | standard | sin cambios | `PublicabilidadEdicion.bas`, `PublicabilidadRiesgo.bas`, `PublicacionLog.cls` | — | — 
| — | Calidad — tareas | standard | sin cambios | `TareasCalidad.cls`, `ArbolTareasCalidad.cls` | — | — 
| — | Técnico — tareas | standard | sin cambios | `TareasTecnico.cls`, `ArbolTareasTecnico.cls` | — | — 
| — | Indicadores | minimal | modificada | `mIndicador.bas`, `modIndicadorDashboard.bas`, `IndicadorRiesgosRepositorio.bas` | — | se apoyan en consultas efímeras, sin consulta guardada que replicar 
| — | Informes | minimal | modificada | `InformeRiesgoHTML.bas`, `InformeRiesgoPDFServicio.cls`, `ExcelInforme.bas` | — | la salida es un fichero en disco, no una descarga autenticada 
| — | Histórico de estados del riesgo | standard | pendiente de ficha | `Riesgo.cls`, `RiesgosGestionRiesgo` | — | — 
| — | Catálogos | standard | sin cambios | `AreaImpacto`, `CarenciasExplicacion` (vía `CambioExplicacion`), `Juridica`, `MitigacionValor`, `NC`, `OrganoContratacion`, `Pedido`, `RAC` | — | — 
| — | Exportación e impresión | minimal | modificada | `ExcelInforme.bas`, `InformeRiesgoHTML.bas`, `InformeRiesgoPDFServicio.cls` | — | la salida es un fichero local 
| — | Integraciones | standard | modificada | `Constructor.bas`, `Expediente.cls`, `ExpedienteResponsable.cls`, `NC.cls` | seis backends externos | se abre un backend por integración: 15 tablas vinculadas de seis backends 
| — | Acoplamiento DAO directo | minimal | modificada | `getdb` (`Variables Globales.bas:485`) | — | 362 llamadas a `getdb` en 42 ficheros de producción 
| — | Concurrencia sin atomicidad | standard | añadida | `PublicabilidadEdicion.bas`, `FuncionesGENERALES.bas` | — | un único `BeginTrans` y ningún `RollbackTrans`; identificadores por `Max + 1` 

El PRD resume 28 dominios (14 sin cambios, 9 modificados, 1 añadido y 4 pendientes de ficha), pero su tabla enumera 26 entradas. Este índice reproduce las 26 entradas enumeradas; la diferencia queda sin resolver porque el PRD no identifica las dos restantes.

### HPS Solicitudes

Fuente: [`prd/capacidades.md`](../03-aplicaciones/hps-solicitudes/prd/capacidades.md) · commit fijado `cfa1937ab05a9e40150ea9236243898de96f84da` · identificadores `C-HPS-S-1..18` · vocabulario de estado: `sin cambios`.

Las 18 capacidades salen «sin cambios en el código»: los 74 ficheros de código coinciden con el snapshot. La disposición aprobada es una decisión de producto y va en la columna de excepción.

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| C-HPS-S-1 | Arranque e identidad. | critical | sin cambios | `Variables Globales.bas` (`EVE`), `constructor.bas` (`getUsuario`), `Form_frmSplash` | identidad compartida con Lanzadera (D-HPS-8) | disposición aprobada: preservar y modernizar; integrar con la identidad única de la web (D-HPS-8) |
| C-HPS-S-2 | Configuración de backend. | minimal | sin cambios | `constructor.bas` (`getConfiguracion`), `Configuracion.cls`, `ConfiguracionOperaciones.cls`, `Form_FormConfiguracion` | — | disposición aprobada: preservar `TbConfiguracion` en backend (singleton, decisión del propietario) y sustituir las TempVars por `GET /configuracion` |
| C-HPS-S-3 | Flags de operación. | minimal | sin cambios | `Variables Globales.bas` (`ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado`) | — | disposición aprobada: mantener como flags de aplicación durante la coexistencia; candidato a tipar como `BOOLEAN` en la web |
| C-HPS-S-4 | Roles (dos niveles: administrador y técnico). | critical | sin cambios | `Variables Globales.bas` (`EsUsuarioTecnicoCalculado`), `Form_Form0BDOpciones` | identidad compartida con Lanzadera (D-HPS-8) | disposición aprobada: preservar y modernizar; el modelo de dos niveles se traduce a `role: 'admin' \| 'tech'` y se valida en cada endpoint |
| C-HPS-S-5 | Solicitudes: alta, edición, renovación, cambio de tipo y baja. | standard | sin cambios | `Solicitud.cls`, `SolicitudOperaciones.cls`, `SolicitudFechas.cls`, `SolicitudFechasOperaciones.cls` | — | disposición aprobada: preservar y modernizar; la API web cubrirá las cinco operaciones |
| C-HPS-S-6 | Traspasos a ONS. | standard | sin cambios | `Form_FormAdjuntaTraspasoONS`, `Form_FormSolicitudAltaTraspaso`, `Form_FormSolicitudesAltaTraspasoDatos`, `Form_FormSolicitudesTraspasoFechas`, `SolicitudOperaciones.RegistroEnvioTraspasoONS` | — | disposición aprobada: preservar y modernizar; el traspaso se cubre con `POST /solicitudes/{id}/traspaso-ons` y transacción atómica obligatoria (D-HPS-11) |
| C-HPS-S-7 | Adjuntos Excel. | standard | sin cambios | `Form_FormAdjuntarExcelSolicitante`, `SolicitudOperaciones.RegistroEnvioExcel` | — | disposición aprobada: preservar y modernizar; el adjunto se cubre con `POST /solicitudes/{id}/adjuntos` y validación de tipo MIME y tamaño |
| C-HPS-S-8 | Plantillas HTML. | standard | sin cambios | `Form_FormPlantillasHTML`, `Form_FormWeb`, `Form_FormCorreo` (`ComandoVerHTML_Click`), `Form_FormEnvioCorreo` (`EstablecerDatos`), `Form_FormCorreoPruebas` (`cmdVerHTML_Click`) | — | disposición aprobada: preservar y modernizar; las plantillas se sirven firmadas con `ContentVersion` y el render local se sustituye por el navegador del usuario |
| C-HPS-S-9 | Renovación. | standard | sin cambios | `Form_FormSolicitudRenovacion`, `Form_FormSolicitudRenovacionDatos`, `Form_FormSolicitudRenovacionFechas`, `Form_FormSolicitudesAltaRenovacionFechas` | — | disposición aprobada: preservar y modernizar; la API web cubre la renovación como `POST /solicitudes/{id}/renovacion` |
| C-HPS-S-10 | Cambio de tipo. | standard | sin cambios | `Form_FormCambioTipo` | — | disposición aprobada: preservar y modernizar; la API web cubre el cambio como `PATCH /solicitudes/{id}/tipo` con validación de rol |
| C-HPS-S-11 | Justificaciones. | standard | sin cambios | `Justificacion.cls`, `JustificacionServicio.cls`, `Form_frmJustificacion`, `Form_frmBuscarJustificacion` | — | disposición aprobada: preservar y modernizar; `GET /justificaciones?activa=...&q=...` con índice `GIN` sobre `titulo` |
| C-HPS-S-12 | Responsables. | standard | sin cambios | `Responsable.cls`, `ResponsableOperaciones.cls`, `Form_FormResponsable`, `Form_FormResponsablesGestion` | usuarios de la web (FK a `responsables.id`, decisión P-EXP-3) | disposición aprobada: preservar y modernizar; la FK por `email` se sustituye por FK numérica a `responsables.id` (decisión P-EXP-3) |
| C-HPS-S-13 | Motivos HPS. | standard | sin cambios | `MotivoHPS.cls`, `Form_FormMotivoHPS`, `Form_FormMotivosHPS` | — | disposición aprobada: preservar y modernizar; `GET /motivos-hps?tipo=...` con índice en `(TipoHPS, Grado)` |
| C-HPS-S-14 | Entidades y expedientes. | minimal | sin cambios | `Form_FormEntidades`, `Form_FormExpedientesBusqueda`, `Form_FormExpedienteDetalle` | Expedientes | disposición aprobada: preservar y modernizar; la búsqueda de expediente se traduce a `GET /expedientes?palabra=...&page=...` con `ILIKE '%…%'` y paginación en servidor |
| C-HPS-S-15 | Tareas pendientes del tramitador (11 estados). | minimal | sin cambios | `Form_FormTareasTramitadorPendientes` | — | disposición aprobada: preservar y modernizar; `GET /tareas?estado=...&tramitador=...&page=...` en lugar de las once consultas |
| C-HPS-S-16 | Log general (`TbLogsGeneral`). | critical | sin cambios | `LogGeneral.cls`, `LogGeneralOperaciones.cls` | — | disposición aprobada: preservar y modernizar; partición por `Fecha` (mensual o semestral) y seudonimización de `DNI` y `Usuario` (D-HPS-3) |
| C-HPS-S-17 | Correos. | standard | sin cambios | `Correo.cls`, `CorreoOperaciones.cls`, `CorreoServicio.cls`, `Form_FormCorreo`, `Form_FormCorreosGestion`, `Form_FormEnvioCorreo`, `Form_FormCorreoPruebas` | — | disposición aprobada: preservar y modernizar; el envío se cubre como `POST /correos` y `POST /correos/preview` con SMTP autenticado y TLS (D-HPS-12) |
| C-HPS-S-18 | Registro en HPS. | standard | sin cambios | `TbSolicitudes.IDUsuarioHPS`, flag `RegistroEnHPS` | HPS | disposición aprobada: preservar y modernizar; la integración se hace vía `POST /hps/registro` con autenticación federada (decisión abierta P-EXP-3) |

### No Conformidades

Fuente: [`prd/capacidades.md`](../03-aplicaciones/no-conformidades/prd/capacidades.md) · commit fijado `1d7a2f826a1e3c582874e4e29886a6183efb9067` · identificadores `CAP-NC-1..45` · vocabulario de estado: `sin cambios`, `modificada`, `mixta`, `añadida`, `retirada`.

El PRD clasifica 45 capacidades frente al snapshot: 8 sin cambios, 30 modificadas, 1 mixta, 1 añadida y 5 retiradas. La excepción `Nemotecnico` —los identificadores funcionales `CodigoNC`/`Nemotecnico` necesitan decisión de unicidad y de vacíos— viaja con la fila de alta de NC de proyecto (CAP-NC-7).

| id | capacidad | tier | estado | objetos de lógica | dependencia principal | excepción |
|---|---|---|---|---|---|---|
| CAP-NC-1 | Arranque: la portada carga el entorno con `EVE`, calcula el rol y decide entre menú reducido de técnico o menú completo. | critical | modificada | `Form_Form0BDTecnicos.cls`, `Form_Form0BDOpciones.cls`, `Variables Globales.bas`, `Entorno.cls` | identidad de Lanzadera | modificada: los formularios son iguales; `Variables Globales.bas` y `Entorno.cls` difieren del snapshot |
| CAP-NC-2 | Identidad y roles: administrador, calidad y técnico, con exclusión mutua; usuario actual resuelto por `getUsuario`. | critical | modificada | `Variables Globales.bas`, `constructor.bas`, `UsuarioAplicacionPermisos.cls` | identidad de Lanzadera | modificada: `main` calcula los roles en el objeto de permisos; el usuario y sus permisos no se mantienen aquí |
| CAP-NC-3 | Configuración de backend e `IDAplicacion` leídos de la configuración local, con distinción de entorno de pruebas. | minimal | modificada | `Variables Globales.bas` | — | modificada: `main` lee `IDAplicacion` de la configuración además del valor fijo por entorno; el destino de la configuración está por decidir con el ERD global |
| CAP-NC-4 | Menús: raíz, submenú de proyectos y de auditorías, con el indicador «Seguimiento usuario / total». | minimal | mixta | `Form_Form0BDOpciones.cls`, `Form_Form0BDOpcionesParteProyectos.cls`, `Form_Form0BDOpcionesAuditorias.cls`, `Funciones Generales.bas` | — | mixta: el menú raíz no cambia y los dos submenús sí; el indicador se repinta al abrir y al volver |
| CAP-NC-5 | Selección de auditoría y alta o edición de una NC de auditoría (datos únicos, requiere acción correctiva y control de eficacia). | standard | modificada | `Form_FormAuditoriaSeleccion.cls`, `Form_FormNCAuditoria.cls`, `Form_FormNCAuditoriaGeneral.cls`, `NCaUDITORIAOperaciones.cls` | — | modificada: `Form_FormNCAuditoriaGeneral` y `NCaUDITORIAOperaciones.cls` difieren del snapshot |
| CAP-NC-6 | Gestión de NC de auditoría: filtros, selección, alta, borrado y habilitación, historial, correo e informe. | standard | modificada | `Form_FormNCAuditoriaGestion.cls` | — | modificada |
| CAP-NC-7 | Alta y edición de una NC de proyecto: expediente, descripción, causa raíz, tipología, responsables, control de eficacia y riesgos. | standard | modificada | `Form_FormNCProyecto.cls`, `Form_FormNCProyectoGeneral.cls`, `NCProyectoOperaciones.cls` | Expedientes | modificada: `Form_FormNCProyectoGeneral` y `NCProyectoOperaciones.cls` difieren del snapshot; el identificador funcional `CodigoNC`/`Nemotecnico` necesita decisión (unicidad y vacíos) |
| CAP-NC-8 | Búsqueda y listado de NC de proyecto con filtros combinables y filtrado en memoria; exportación de la lista a Excel y a Word. | minimal | modificada | `Form_FormNCProyectoGestion.cls` | — | modificada: el formulario tiene código distinto del snapshot |
| CAP-NC-9 | Borrado lógico con motivo y habilitación de NC de proyecto y de auditoría, con recálculo de datos calculados. | standard | modificada | `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls`, `Form_FormNCProyectoGestion.cls`, `Form_FormNCAuditoriaGestion.cls` | — | modificada: ambas clases de operaciones difieren del snapshot |
| CAP-NC-10 | Alta de una NC de proyecto vinculada a otra existente y borrado en cascada de las vinculadas. | standard | modificada | `Form_FormNCProyectoGeneralConVinculoNC.cls`, `NCProyectoOperaciones.cls` | — | modificada: el formulario es igual y `NCProyectoOperaciones.cls` difiere |
| CAP-NC-11 | Acciones correctivas (AC) de proyecto y de auditoría: alta, edición, borrado, responsable, número de acción. | standard | modificada | `Form_FormNCProyectoAcciones.cls`, `Form_FormNCAuditoriaAcciones.cls`, `ACProyecto.cls`, `ACProyectoOperaciones.cls`, `ACAuditoriaOperaciones.cls` | — | modificada: `ACProyecto.cls`, `ACProyectoOperaciones.cls`, `ACAuditoriaOperaciones.cls` y ambos formularios difieren |
| CAP-NC-12 | Tareas o acciones realizadas (AR): alta, detalle, edición, borrado, responsable y fechas de inicio, fin previsto y fin real. | standard | modificada | `Form_FormNCProyectoAcciones.cls`, `Form_FormNCAuditoriaAcciones.cls`, `ARProyecto.cls`, `ARProyectoOperaciones.cls`, `ARAuditoriaOperaciones.cls` | — | modificada: `ARProyecto.cls`, `ARProyectoOperaciones.cls`, `ARAuditoriaOperaciones.cls` y ambos formularios difieren |
| CAP-NC-13 | Historial de AC y de AR, y recálculo de estados de acciones. | standard | modificada | `Form_FormNCProyectoAcciones.cls`, `Form_FormNCAuditoriaAcciones.cls`, `ACProyectoOperaciones.cls`, `ARProyectoOperaciones.cls` | — | modificada |
| CAP-NC-14 | Plan de control de eficacia (fecha prevista y descripción) y su alta modal. | standard | modificada | `Form_FormNCProyectoControlEficaciaAlta.cls`, `Form_FormNCAuditoriaControlEficaciaAlta.cls`, `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls` | — | modificada: el formulario de auditoría difiere, el de proyecto es igual y las operaciones difieren |
| CAP-NC-15 | Resultado del control de eficacia (fecha, resultado y conformidad) y su eliminación. | standard | modificada | `Form_FormNCProyectoControlEficacia.cls`, `Form_FormNCAuditoriaControlEficacia.cls`, `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls` | — | modificada |
| CAP-NC-16 | Motivo por el que la NC no requiere control de eficacia. | standard | sin cambios | `Form_FormMotivosNoRequiereControlEficacia.cls` | Brass y Gestión de Riesgos (catálogo compartido, D86) | sin cambios en el formulario; los formularios que lo abren están modificados; el texto se devuelve al formulario de origen y no escribe tablas por sí mismo |
| CAP-NC-17 | Replanificación de acciones y tareas con motivo, registrada en la operación de AC y AR. | standard | modificada | `ACProyectoOperaciones.cls`, `ARProyectoOperaciones.cls`, `ACAuditoriaOperaciones.cls`, `ARAuditoriaOperaciones.cls`, `ReplanificacionesProyectoOperaciones.cls` | — | modificada; el alta por formulario no existe (a verificar) |
| CAP-NC-18 | Consulta de solo lectura de replanificaciones por NC (acción, tarea y motivo). | minimal | modificada | `Form_FormNCProyectoReplanificaciones.cls`, `Form_FormNCAuditoriaReplanificaciones.cls`, `ReplanificacionesAuditoriaOperaciones.cls` | — | modificada: los formularios no cambian y la capacidad se clasifica como modificada por depender de CAP-NC-17; la pantalla no escribe datos |
| CAP-NC-19 | Documentos de NC: listar, anexar (Examinar y Aceptar), cambiar título, abrir y eliminar. | standard | modificada | `Form_FormNCProyectoDocumentos.cls`, `Form_FormNCAuditoriaDocumentos.cls`, `DocumentoProyectoOperaciones.cls` | — | modificada: ambos formularios y `DocumentoProyectoOperaciones.cls` difieren |
| CAP-NC-20 | Documentos de evidencia de AR (proyecto y auditoría). | standard | modificada | `Form_FormARProyectoDocumentos.cls`, `Form_FormARAuditoriaDocumentos.cls`, `DocumentoProyectoOperaciones.cls`, `DocumentoAuditoriaOperaciones.cls` | — | modificada: ambos formularios difieren |
| CAP-NC-21 | Documentos de auditoría, con modo memoria (antes del alta) y modo directo. | standard | modificada | `Form_FormAuditoriaDocumentos.cls`, `AuditoriaOperaciones.cls`, `Form_FormAuditoriasGestion.cls` | — | modificada: el formulario tiene código distinto del snapshot |
| CAP-NC-22 | Nota libre de la NC: guardar y borrar; en una NC vinculada se muestra la de la asociada en solo lectura. | standard | modificada | `Form_FormNCProyectoNota.cls`, `Form_FormNCAuditoriaNota.cls`, `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls` | — | modificada por las operaciones; los formularios son iguales |
| CAP-NC-23 | Motivo de borrado de la NC: ver y corregir. | standard | modificada | `Form_FormNCProyectoMotivoEliminado.cls`, `Form_FormNCAuditoriaMotivoEliminado.cls`, `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls` | — | modificada por las operaciones; los formularios son iguales |
| CAP-NC-24 | Seguimiento de tareas (pestaña «Tareas»): filtros por responsable, estado y expediente o auditoría, apertura de la NC y exportación a Excel. | minimal | modificada | `Form_FormNCProyectoSeguimientoTareas.cls`, `Form_FormNCAuditoriaSeguimientoTareas.cls`, `Form_FormNCProyectoSeguimiento.cls`, `Form_FormNCAuditoriaSeguimiento.cls`, `SegTareasProyecto.cls` | — | modificada: los cuatro formularios y `SegTareasProyecto.cls` difieren |
| CAP-NC-25 | Revisión de estados de NC (pestaña «Rev. Estados») con los mismos filtros. | minimal | modificada | `Form_FormNCProyectoSeguimientoNC.cls`, `Form_FormNCAuditoriaSeguimientoNC.cls` | — | modificada: ambos formularios difieren |
| CAP-NC-26 | Log de actividad por NC (proyecto y auditoría) y consulta del historial en HTML. | critical | modificada | `Form_FormNCProyectoGestion.cls`, `Form_FormNCAuditoriaGestion.cls`, `LogNCProyecto.cls`, `LogNCAuditoria.cls`, `Funciones Generales.bas` | — | modificada: `LogNCProyecto.cls` y `LogNCAuditoria.cls` no cambian, pero `Funciones Generales.bas`, que escribe el log, difiere |
| CAP-NC-27 | Alta y edición de la cabecera de una auditoría: tipo (interna o externa), fecha de inicio y de fin. | standard | sin cambios | `Form_FormAuditoria.cls`, `AuditoriaOperaciones.cls` | — | sin cambios en el formulario; la clase `AuditoriaOperaciones.cls` es igual |
| CAP-NC-28 | Catálogo de auditorías: filtrar por año y tipo, cambiar fechas, anexos, eliminar y ver sus NC. | standard | modificada | `Form_FormAuditoriasGestion.cls`, `AuditoriaOperaciones.cls` | — | modificada: el formulario tiene código distinto del snapshot |
| CAP-NC-29 | Catálogo de tipologías de NC de proyecto: alta, edición y eliminación. | standard | sin cambios | `Form_FormNCProyectoTipologiaGestion.cls`, `Form_FormTipologiaNCProyecto.cls`, `TipologiaNCProyectos.cls` | — | sin cambios |
| CAP-NC-30 | Asociar riesgos del expediente a una NC de proyecto (multiselección, en memoria durante el alta). | standard | modificada | `Form_formRiesgosSeleccion.cls`, `RiesgoServicio1.bas`, `RiesgoRepositorio.bas`, `NCProyectoOperaciones.cls` | Gestión de Riesgos | modificada: `RiesgoServicio1.bas` es nuevo respecto al snapshot y está vacío en `main`; el riesgo no se edita aquí |
| CAP-NC-31 | Búsqueda de expedientes y devolución del elegido por evento. | minimal | sin cambios | `Form_FormExpedientesBusqueda.cls`, `Expediente.cls`, `Form_FormNCProyectoGestion.cls` | Expedientes | sin cambios: el formulario es igual y `Expediente.cls` no tiene diferencias; los expedientes no se editan aquí |
| CAP-NC-32 | Panel de indicadores semestrales de NC de proyecto (abiertas, replanificadas, stock activo y con riesgo) con exportación del detalle a Excel. | minimal | sin cambios | `Form_FormIndicadores.cls`, `IndicadorServicio.cls`, `IndicadorRepositorio.bas` | — | sin cambios |
| CAP-NC-33 | Correo de la NC: componer el informe HTML y registrarlo como correo pendiente (el envío real no es de este formulario). | standard | sin cambios | `Form_FormCorreo.cls`, `Correo.cls`, `HTML.bas` | — | sin cambios |
| CAP-NC-34 | Aviso por correo al administrador ante errores. | standard | modificada | `Funciones Generales.bas`, `Correo.cls` | — | modificada: `Funciones Generales.bas` difiere y `Correo.cls` no cambia |
| CAP-NC-35 | Informe Word de una NC o de la lista filtrada, y exportación de la lista a Excel. | minimal | modificada | `Informe.cls`, `Funciones Generales.bas`, `Form_FormNCProyectoGestion.cls`, `Form_FormNCAuditoriaGestion.cls` | — | modificada: `Funciones Generales.bas` difiere |
| CAP-NC-36 | Informe HTML de NC y de historial. | minimal | sin cambios | `HTML.bas` | — | sin cambios |
| CAP-NC-37 | Utilidades de normalización de datos del `Instalador` (estados, tipologías, responsables). | minimal | sin cambios | `Instalador.bas` | — | sin cambios; su alcance en la web está por decidir |
| CAP-NC-38 | Clase de informe de NC de auditorías (`InformeNCAuditorias`), vacía de procedimientos. | minimal | añadida | `InformeNCAuditorias.cls` | — | añadida: no está en el snapshot; solo declara el error, sin comportamiento que aceptar |
| CAP-NC-39 | Caché de indicadores (`ModuloCacheIndicadores`). | minimal | retirada | — | — | retirada: el módulo no está en `main`; los indicadores se calculan sobre las tablas, sin caché |
| CAP-NC-40 | Caché de listados de NC (`NCAuditoriaListadoCache`, `CacheNCCacheRepositorio`). | minimal | retirada | — | — | retirada: los módulos y la clase no están; el listado filtra en memoria |
| CAP-NC-41 | Kill switch y bancos de prueba de caché (`Test_KillSwitch`, `Test_Cache*`). | minimal | retirada | — | — | retirada: no existe interruptor de caché en `main` |
| CAP-NC-42 | Ayudantes `modFrm*Helper`, `modFormInteractionHelper` y servicios `CierreNCService`, `DocumentoService`, `ReplanificacionesService`. | minimal | retirada | — | — | retirada: la lógica vive en los formularios y en las clases `*Operaciones` |
| CAP-NC-43 | Viewmodels `NCProyectoListItemVM` y `NCProyectoDetailVM`. | minimal | retirada | — | — | retirada: la web define sus propios DTO de lista y detalle |
| CAP-NC-44 | Caché por NC (tabla de caché, JSON de NC, AC, AR, replanificaciones y riesgos), con transacción de caché. | minimal | modificada | `CacheNCProyecto.bas` | — | modificada; la decisión de portarla se toma con el ERD global |
| CAP-NC-45 | Mantenimiento de la caché: estadísticas, integridad, regeneración, población masiva y limpieza de logs. | minimal | modificada | `InicializadorCache.bas` | — | modificada |

## Checklist

- [ ] Cada capacidad del índice tiene su evidencia en el PRD de su aplicación.
- [ ] Cada regla de negocio sin test se etiqueta `Verified-static` con su test pendiente.
- [ ] Las excepciones se referencian desde la matriz global y desde la ficha.

## Pendientes abiertos

- **ERD de origen.** Sigue pendiente en las ocho aplicaciones: la extracción de solo lectura con `dysflow-erd` depende del runner Windows del propietario. Ninguna de las capacidades de este índice se apoya en esa extracción.
- **Fases 5 en curso.** Solo HPS (#1076) sigue con la fase 5 abierta. HPS Solicitudes (#1077) y No Conformidades (#1080) han entregado su clasificación de capacidades en `main` y sus filas ya están en este índice, con la excepción `Nemotecnico` de No Conformidades incluida en la fila de alta de NC de proyecto.

## Siguiente paso

Cruzar este índice con la matriz legacy→web ([`docs/07-migracion/matriz-legacy-a-web.md`](../07-migracion/matriz-legacy-a-web.md)), que ya reproduce una fila por capacidad de las siete aplicaciones con fase 5 cerrada, con componente web candidato, riesgo principal, dependencia, excepción y confianza, y una sección para las capacidades cuya lógica vive en un evento de formulario.
