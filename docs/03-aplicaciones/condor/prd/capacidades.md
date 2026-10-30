# Capacidades de Condor frente al estudio previo

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Issue:** #1075 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Clasifica las capacidades del [estudio previo](../capabilities.md) como **añadida**, **modificada**, **retirada** o **sin cambios** respecto al commit fijado, con evidencia, y las enlaza con las hojas de este PRD. El estudio previo no numera las capacidades; aquí se numeran C-CON-1 a C-CON-22 solo para referirse a ellas.

Abreviaturas de cita: `FU`, `modules/FUNCIONES UTILES.bas`; `RECH`, `modules/RechazoRepositorio.bas`; `SB`, `modules/SandboxBuilder.bas`; `SIM`, `modules/modSimuladorNotificaciones.bas`; `US`, `classes/UsuarioServicio.cls`; `WF`, `classes/WorkflowServicio.cls`.

## Método y límites

- La clasificación compara cada capacidad con los objetos del **commit fijado**, no con el snapshot: el estudio previo se hizo sobre `data/staging` del 2026-08-06, que ya no está en `main` del legacy. El delta de objetos está en el [README](../README.md).
- «Sin cambios» significa que la capacidad sigue existiendo con el mismo comportamiento observado; no prueba que funcione ni que el estudio previo la describiera bien.
- Este estudio **no revisa la épica ni los walkthroughs** y **no crea issues de implementación**: las propuestas están en `issues-propuestas` cuando llegue.
- Las citas son `ruta:línea` del commit fijado; se comprobaron mecánicamente.

## Clasificación

| Capacidad | Objetos de lógica | Estado frente al estudio previo | Fichas del PRD | Observación |
|---|---|---|---|---|
| C-CON-1 Arranque, identidad y validación crítica | `Variables Globales.bas` (`EVE`, `EVE.ValidacionCritica`) | sin cambios | [SnapshotYHash](modulos/Modulo_SnapshotYHash.md), [Usuarios](clases/Clase_UsuariosYNotificaciones.md) | `EVE` sigue siendo el arranque y valida conexión, usuario y entorno con error 513 (`Variables Globales.bas:56`, `:96`, `:112`, `:143`) |
| C-CON-2 Roles y jerarquía | `modEnumeradores.bas` (`Enum rol`), `UsuarioServicio.DeterminarRol` | **modificada** | [Usuarios](clases/Clase_UsuariosYNotificaciones.md) | El estudio citaba `JerarquiaRolesHelper.bas`, que **no existe** en el commit. La jerarquía vive ahora en `DeterminarRol` (`US:78-111`) y en `PermisoSuficiente` (`WF:1410-1418`) |
| C-CON-3 Configuración de backend | `Variables Globales.bas` (`TbConfiguracionBackends`) | **retirada** | [FuncionesUtiles](modulos/Modulo_FuncionesUtiles.md) | No hay ninguna cita de `TbConfiguracionBackends` en el commit. La configuración de conexión vive en el propio módulo, con rutas y contraseña (`FU:233-252`) |
| C-CON-4 Modo de pruebas y sandbox seguro | `FUNCIONES UTILES.bas` (`m_TestingMode`, `m_BackendSandboxURL`) | **modificada** | [EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) | Ni `m_TestingMode` ni `m_BackendSandboxURL` existen. El aislamiento de pruebas lo dan ahora `SandboxBuilder`, `SandboxProvisioner`, `SnapshotHelper` y `TestSandbox`, por copia de ficheros y no por enrutado de `getdb()` |
| C-CON-5 Login por comando de Lanzadera | `Variables Globales.bas` (`m_Command`) | sin cambios | [Usuarios](clases/Clase_UsuariosYNotificaciones.md) | `m_Command` se lee del argumento de línea de órdenes (`Variables Globales.bas:103`) |
| C-CON-6 Gestión de solicitudes | `SolicitudServicio.cls`, `Form_frmGestionSolicitud.cls` | sin cambios | [Solicitud](clases/Clase_Solicitud.md) | Alta, edición, baja, cambio de estado, rechazo, validación y visor |
| C-CON-7 Cuatro tipos de solicitud | `DatosPC`, `DatosCDCA`, `DatosCDCASUB`, `DatosPCSUB` y sus servicios | **retirada** (una de las cuatro) | [DatosPC](clases/Clase_DatosPC.md), [DatosCDCA](clases/Clase_DatosCDCA.md), [DatosCDCASUB](clases/Clase_DatosCDCASUB.md) | El tipo `PC_SUB` **no existe** en el commit: faltan siete formularios y cinco clases (D-CON-1). Quedan tres tipos |
| C-CON-8 Catálogo de estados y transiciones | `modEnumeradores.bas` (`enumEstados`), `WorkflowRepositorio.bas`, `tbTransiciones` | sin cambios | [Catalogos](modulos/Modulo_RepositoriosCatalogos.md), [Workflow](clases/Clase_WorkflowServicio.md) | Nueve estados, numerados del 1 al 9, y transiciones en tabla |
| C-CON-9 Validación de calidad | `Solicitud.cls` (`revisionCalidadEstado`, `revisionCalidadComentarios`), `ValidacionRevisionServicio.cls` | sin cambios | [RevisionYRechazo](clases/Clase_RevisionYRechazo.md) | Los dos campos siguen en la entidad y el ciclo de revisión por ordinal sigue vigente |
| C-CON-10 Historial de rechazos | `tbHistorialRechazos`, `Rechazo.cls`, `RechazoServicio.cls` | **modificada** | [RevisionYRechazo](clases/Clase_RevisionYRechazo.md), [RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) | `tbHistorialRechazos` **no existe** en el volcado vigente: el rechazo vive en `tbRechazos` con marca de activo y desactivación del anterior (`RECH:146`, `RECH:179`) |
| C-CON-11 Adjuntos por etapa | `AdjuntosServicio.cls`, `AdjuntoRepositorio.bas` | sin cambios | [Adjuntos](clases/Clase_Adjuntos.md) | Alta, baja y asociación por etapa del flujo, con copia del fichero |
| C-CON-12 Visor de estado embebido | `Form_frmGestionSolicitud.cls` (`webInfo`), `WebVisorCacheServicio.cls`, `WorkflowServicio.GenerarHTML_VisualizadorDeEstado` | sin cambios | [EntornoYVisor](clases/Clase_EntornoYVisor.md), [Workflow](clases/Clase_WorkflowServicio.md) | El control de navegador sigue presente y el servicio sigue generando el HTML con caché en disco |
| C-CON-13 Motor de workflow | `WorkflowServicio.cls` | sin cambios | [Workflow](clases/Clase_WorkflowServicio.md) | Transiciones, precondiciones, permisos por bloque y vuelta atrás |
| C-CON-14 Vinculación con expediente | `Expediente.cls`, `ExpedienteServicio.cls` | sin cambios | [Expediente](clases/Clase_Expediente.md) | Lectura desde la base de Lanzadera y composición del código de solicitud |
| C-CON-15 Vinculación con no conformidades | `NoConformidad.cls`, `NoConformidadServicio.cls` | sin cambios | [RevisionYRechazo](clases/Clase_RevisionYRechazo.md) | Lectura de la base externa y escritura del código de concesión al formalizar |
| C-CON-16 Suministradores | `Suministrador.cls`, `SuministradorServicio.cls` | sin cambios | [Usuarios](clases/Clase_UsuariosYNotificaciones.md) | Catálogo y valores por defecto del bloque |
| C-CON-17 Documentos y PDF | `DocumentoServicio.cls` | sin cambios | [Expediente](clases/Clase_Expediente.md) | Generación con plantilla de Word, exportación a PDF e importación de RAC |
| C-CON-18 Notificaciones | `NotificacionServicio.cls`, `MockNotifServ.cls` | **modificada** | [Usuarios](clases/Clase_UsuariosYNotificaciones.md) | `MockNotifServ` no existe; la simulación la cubre `modSimuladorNotificaciones`, que **envía notificaciones reales** (`SIM:6-107`) |
| C-CON-19 Auditoría en tres tablas | `LogServicio.cls`, `CondorError.cls` | sin cambios | [Trazabilidad](clases/Clase_Trazabilidad.md) | Cambios, errores y estados; el error tipado se autoregistra al lanzarse |
| C-CON-20 Mapeo de campos | `MapeoServicio.cls`, `MapeoRepositorio.bas` | sin cambios | [EntornoYVisor](clases/Clase_EntornoYVisor.md) | Mapeo por tipo de solicitud y siembra de marcadores en la plantilla |
| C-CON-21 Hashing y snapshot | `HashHelper.bas`, `SnapshotHelper.cls` | sin cambios | [SnapshotYHash](modulos/Modulo_SnapshotYHash.md) | SHA-256 con la API de Windows y snapshot JSON por tipo de solicitud |
| C-CON-22 Batería de pruebas | `modBattery_Canonical.bas` | sin cambios | [EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) | 166 procedimientos, 161 de ellos pruebas unitarias del flujo y los datos |

**Resumen:** 0 capacidades añadidas, 5 modificadas (C-CON-2, C-CON-4, C-CON-7 parcialmente, C-CON-10, C-CON-18), 2 retiradas (C-CON-3 y el tipo `PC_SUB` de C-CON-7) y 15 sin cambios. Ninguna capacidad nueva ha aparecido en el commit respecto al snapshot.

## Capacidades retiradas o sustituidas y su consecuencia

| Capacidad | Qué pasa con ella | Consecuencia para la web |
|---|---|---|
| C-CON-3 Configuración de backend | La configuración de conexión sale del modelo de datos y pasa al código (`FU:233-252`) | El destino externaliza la configuración; no hay capacidad equivalente que migrar |
| C-CON-4 Sandbox por enrutado | El aislamiento pasa de ser un conmutador de `getdb()` a clonar ficheros (`SB:51-618`) | El entorno de pruebas del destino es una base propia, no una copia de ficheros |
| C-CON-7 tipo `PC_SUB` | Desaparece el cuarto tipo con sus siete formularios y cinco clases | Si el producto sigue necesitando ese tipo, es una capacidad **añadida** que este PRD debe recoger |
| C-CON-10 `tbHistorialRechazos` | El historial pasa a ser una tabla con marca de activo | El destino necesita la unicidad de «un rechazo activo» como restricción, no como código |
| C-CON-18 `MockNotifServ` | La simulación de notificaciones pasa a dispararlas de verdad | El entorno de pruebas necesita un doble explícito antes de poder ejercitar el flujo |

## Discrepancias que quedan abiertas

- El estudio previo afirma «`getdb()` con **1 caller**». La medición del commit fijado da **155 llamadas en 13 ficheros** de producción: el patrón es el contrario al que el estudio describe.
- El estudio previo cita `SandboxConfig.cls`, `SandboxGestor.cls` y `SandboxValidator.cls`; en el commit no existen como clases y solo aparecen nombrados en comentarios de `SandboxBuilder` y `modBattery_Canonical`.
- El estudio enumera nueve estados; el enumerador del commit mantiene los nueve, pero el flujo observado en las hojas usa siete de ellos en las precondiciones.