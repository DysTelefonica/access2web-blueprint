# NoConformidades — capacidades observadas

## Resultado

La aplicación cubre un agregado de gestión de No Conformidades (NC) sobre dos ejes: **Auditorías** (NC originadas en auditorías internas/externas con AC, AR, Control de Eficacia) y **Proyectos** (NC originadas en la operativa de proyectos con AC, AR, Control de Eficacia, Replanificaciones). Tiene además una **capa de caché selectivo** muy madura con kill switch y diagnósticos, un sistema de **vinculación con Riesgos** (Gestion_Riesgos) y un módulo de **indicadores**. La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE` con `LeeConfiguracionLocal`, `Constructor.getUsuario`, carga de `Entorno.ColItems`, IDAplicacion `8` (producción) / `81` (pruebas) | `src/modules/Variables Globales.bas:497` (EVE), `:145` (getdb); `src/modules/constructor.bas:119` (getUsuario) |
| Configuración de backend | `TbConfiguracionBackends` con `BackendActivo` (PROD/LOCAL/SANDBOX), `BackendProduccion`, `BackendSandbox`, `IDAplicacion`, `EnPruebas`/`EnDesarrollo`, `PasswordBackend` | `src/modules/Variables Globales.bas:202-357` (`LeeConfiguracionLocal`) |
| Roles | Tres: Administrador, Calidad, Técnico. Calculados en `EVE` por lookup en `m_ObjEntorno.ColUsuariosAdministradores` | `src/modules/Variables Globales.bas:497+`, `src/forms/Form_Form0BDOpciones.cls` |
| NC de Auditoría | alta, edición, baja, AC (Acciones Correctivas), AR (Acciones Preventivas), Control de Eficacia, Replanificaciones, documentos, notas, motivos de eliminación, seguimientos | `NCAuditoria.cls`, `NCaUDITORIAOperaciones.cls`, `Form_FormNCAuditoria*.cls` (~15 forms derivados) |
| AC/AR de Auditoría | generación, edición, cierre, responsables, fechas | `ACAuditoria.cls`, `ACAuditoriaOperaciones.cls`, `ARAuditoria.cls`, `ARAuditoriaOperaciones.cls`, `DocumentoAuditoria.cls`, `DocumentoAuditoriaOperaciones.cls`, `Form_FormNCAuditoriaAC.cls`, `Form_FormNCAuditoriaAR.cls` |
| Control de Eficacia (Auditoría) | alta, edición, motivo de no requerir, gestión | `Form_FormNCAuditoriaControlEficacia.cls`, `Form_FormNCAuditoriaControlEficaciaAlta.cls` |
| Replanificaciones (Auditoría) | generación, gestión | `ReplanificacionesAuditoria.cls`, `ReplanificacionesAuditoriaOperaciones.cls`, `Form_FormNCAuditoriaReplanificaciones.cls` |
| NC de Proyecto | alta, edición, baja, AC, AR, Control de Eficacia, Replanificaciones, documentos, notas, motivos de eliminación, seguimientos, vínculo con Riesgo | `NCProyecto.cls`, `NCProyectoOperaciones.cls`, `Form_FormNCProyecto*.cls` (~15 forms derivados) |
| AC/AR de Proyecto | análogo a Auditoría | `ACProyecto.cls`, `ACProyectoOperaciones.cls`, `ARProyecto.cls`, `ARProyectoOperaciones.cls`, `DocumentoProyecto.cls`, `DocumentoProyectoOperaciones.cls`, `Form_FormNCProyectoAC.cls`, `Form_FormNCProyectoAR.cls` |
| Control de Eficacia (Proyecto) | análogo a Auditoría | `Form_FormNCProyectoControlEficacia.cls`, `Form_FormNCProyectoControlEficaciaAlta.cls` |
| Replanificaciones (Proyecto) | análogo a Auditoría | `ReplanificacionesProyecto.cls`, `ReplanificacionesProyectoOperaciones.cls`, `Form_FormNCProyectoReplanificaciones.cls` |
| Gestión de Auditorías | listado, selección, alta | `Auditoria.cls`, `AuditoriaOperaciones.cls`, `Form_FormAuditoria.cls`, `Form_FormAuditoriaSeleccion.cls`, `Form_FormAuditoriasGestion.cls`, `Form_FormAuditoriaDocumentos.cls` |
| Gestión de Proyectos | entrada, opciones | `Form_Form0BDOpciones.cls`, `Form_Form0BDOpcionesParteProyectos.cls`, `Form_Form0BDOpcionesAuditorias.cls` |
| Gestión de Tipologías | alta, edición, baja de tipologías NC | `TipologiaNCProyectos.cls`, `Form_FormTipologiaNCProyecto.cls`, `Form_FormNCProyectoTipologiaGestion.cls` |
| Vinculación con Riesgos | selección de riesgos para NC de proyecto | `Riesgo.cls`, `RiesgoServicio.cls`, `Form_formRiesgosSeleccion.cls` |
| Vinculación con Expedientes | búsqueda de expedientes para NC | `Expediente.cls`, `ExpedienteResponsable.cls`, `Form_FormExpedientesBusqueda.cls` |
| Logs y seguimientos | log de auditoría, log de proyecto, seguimientos NC, seguimientos tareas | `LogNCAuditoria.cls`, `LogNCProyecto.cls`, `SegNCAuditoria.cls`, `SegNCProyecto.cls`, `SegTareasAuditoria.cls`, `SegTareasProyecto.cls`, `Form_FormNCAuditoriaSeguimiento*.cls`, `Form_FormNCProyectoSeguimiento*.cls` |
| Seguridad | clases dedicadas a seguridad por NC | `SegNCAuditoria.cls`, `SegNCProyecto.cls` |
| **Caché selectivo maduro** | tabla `TbCacheNCProyecto` con `CacheValida`, `FechaCache`, `FechaUltimoUso`, `HitsConsultas`, `TamanioBytes`, `Version`, `DatosNC`/`DatosACs`/`DatosARs` (JSON), `TbLogCache` (logs), kill switch, comandos de mantenimiento (invalidar, eliminar, regenerar, poblar masivamente, limpiar logs), diagnóstico de integridad, estadísticas de uso, distribución de tamaño | `src/modules/InicializadorCache.bas` (líneas 630-1049 — `AyudaCache`, `InvalidarCachesObsoletos`, `EliminarCachesInvalidos`, `MostrarEstadisticasUso`, `DiagnosticarIntegridad`, `RegenerarCachesInvalidos`, `PoblarCacheMasivo`, `LimpiarLogsAntiguos`, `MostrarRendimiento`); `src/classes/CacheNCCacheRepositorio.cls` |
| Indicadores | servicio de indicadores, repositorio, integración con caché | `IndicadorServicio.cls`, `IndicadorRepositorio.cls`, `ModuloCacheIndicadores.bas`, `Form_FormIndicadores.cls` |
| Informes | generación de informes y de NC de Auditorías | `Informe.cls`, `InformeNCAuditorias.cls` |
| Tareas (worklists) | clases de seguimiento de tareas | `SegTareasAuditoria.cls`, `SegTareasProyecto.cls` |
| Correos | envío de correos al administrador en errores | `Correo.cls`, `Form_FormCorreo.cls` |
| Catálogos | jurídicas | `Juridica.cls` |
| ViewModel de listado | `NCProyectoListItemVM` con propiedades flat (`Estado`, `FechaApertura`, `FECHACIERRE`, `Proyecto`, `VEHICULO`, `ResponsableTelefonica`, `RESPONSABLECALIDAD`, `Cerrada`, `RequiereACR`, `ACR`, `RequiereControlEficacia`, `Nemotecnico`, `Expediente`) | `src/classes/NCProyectoListItemVM.cls:63-128` |
| Identidad / permisos | `Usuario`, `UsuarioAplicacionPermisos`, `Entorno` | `src/classes/Usuario.cls`, `src/classes/UsuarioAplicacionPermisos.cls`, `src/classes/Entorno.cls` |
| HTML / JSON helpers | `HTML`, `JsonConverter`, `JSONHelper`, `HTML.bas` | sus respectivos `.cls`/`.bas` |
| Integraciones | Lanzadera (identidad), Expedientes, Gestion_Riesgos (vinculación), HPS, AGEDYS | `constructor.bas` (vía `getdbLanzadera` o vía tablas compartidas) |
| Acoplamiento DAO directo | `getdb()` con **344 callers** (el más alto de las 8 aplicaciones). Múltiples tablas, transacciones y cacheo pesado. | `src/modules/Variables Globales.bas:145` (`getdb`) |

## Reglas de conservación

- La aplicación distingue explícitamente NC de **Auditoría** vs NC de **Proyecto**: dos jerarquías de clases paralelas, dos juegos de forms paralelos, dos `*Operaciones` separados. La nueva plataforma debe **preservar esta distinción como dos dominios dentro del mismo módulo**, no como dos módulos separados (comparten identidad, catálogos, caché, infraestructura).
- El **caché selectivo** es uno de los patrones más maduros del ecosistema: kill switch, métricas, logs, comandos de mantenimiento, diagnóstico de integridad. Se preserva **íntegramente** como referencia para el puerto de caché de la nueva plataforma (D91).
- La **vinculación NC ↔ Riesgo** se hace por código de riesgo, no por FK física; se respeta la regla de "no asumir propiedad de integraciones externas".
- El **`ViewModel`** (`NCProyectoListItemVM`) es el patrón canónico de la nueva plataforma: cada agregado expone una vista de lista optimizada para UI, separada de la entidad de dominio. Se traduce a un DTO de respuesta en FastAPI + Pydantic.
- La capa de **logs** (`TbLogCache`, `SegTareas*`) es **evidencia de comportamiento** que se conserva como referencia para los logs estructurados canónicos del blueprint (D27).
- APAP y APAP_WEB no aparecen ni se mencionan (regla transversal del blueprint).

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre `00_main` se consultó primero. La inspección Dysflow sobre el binario y el backend autoritativo queda **bloqueada por el fallo de `list_objects`** (ver [Seguridad § D89](security-rules.md#d89-diagnóstico-del-fallo-de-list_objects-de-dysflow)); requiere resolución previa.