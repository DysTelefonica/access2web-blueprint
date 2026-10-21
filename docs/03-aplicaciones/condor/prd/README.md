# PRD — Condor

> **Estado:** fase 1 (fijación e inventario). Los formularios, módulos y clases quedan **pendientes**; las fichas, la matriz, la coexistencia, las capacidades y las issues propuestas llegan en las PR siguientes de la cadena.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/condor |
| Commit fijado | `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` (`main`, 2026-07-18, «chore(agents): delegate dysflow config to global») |
| Índice CodeGraph-VBA | 203 ficheros indexados, 9 039 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/condor/src` (2026-08-06) |
| Issue | #1075 (método y plantilla de #1069) |

## Alcance actual

Fija el commit y concilia el inventario. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis. El repositorio legacy contiene además `ERD/Estructura_Datos.md` (documentación manual del modelo de datos, 493 líneas) y `CONDOR.accdb`; ninguno se ha leído todavía como evidencia del ERD de origen.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `c88347e` | Snapshot 2026-08-06 | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/Form_*.cls`, con su `.form.txt`) | 38 | 45 | −7 (los siete de `PC_SUB`) |
| Módulos estándar (`src/modules/*.bas`) | 42 | 74 | −34 sin equivalente en el commit, +2 nuevos |
| Clases (`src/classes/*.cls`) | 51 | 57 | −7 sin equivalente, +1 nueva |

Los 93 ficheros de módulos y clases del commit coinciden con el recuento de la issue (42 + 51). La comparación normaliza las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas.

### Delta de código frente al snapshot

- **Formularios retirados (7):** `Form_frmDatosPCSUB`, `Form_subfrmDatosPCSUB_AprobacionSuministrador`, `Form_subfrmDatosPCSUB_DecisionFinal`, `Form_subfrmDatosPCSUB_DictamenRAC`, `Form_subfrmDatosPCSUB_Generales`, `Form_subfrmDatosPCSUB_Impacto`, `Form_subfrmDatosPCSUB_Propuesta`. Con ellos desaparecen `DatosPCSUB`, `DatosPCSUBServicio`, `DatosPCSUBViewModel`, `DatosPCSUBGuardarHelper` y `DatosPCSUBRepositorio`. El estudio previo daba cuatro tipos de solicitud, entre ellos `PC_SUB` ([epic.md](../epic.md), F2): queda por confirmar en la ficha de `frmAltaSolicitud` si el tipo sigue existiendo (D-CON-1).
- **Módulos y clases del snapshot ausentes en el commit (11 módulos no de prueba, 7 clases, 23 módulos `Test_*`):** helpers (`DecisionFinalHelper`, `DictamenRACDefaultsHelper`, `JerarquiaRolesHelper`, `NotificacionHelper`, `PrecondicionesWorkflowHelper`, `DatosTipoSolicitudHelper`), utilidades de pruebas y sandbox (`SandboxCloneHelper`, `SandboxConfig`, `SandboxGestor`, `SandboxValidator`, `SchemaInspector`, `TestHelper`, `MockNotifServ`) y 23 módulos `Test_*`. Su lógica puede haberse integrado en otros objetos: no se presume (D-CON-2).
- **Nuevos en el commit:** `ResultadoImportacionRAC` (módulo y clase) y `Módulo1` (4 líneas).
- **Variante de formulario:** `src/forms/Form_frmGestionSolicitud.form-admCasa.txt` convive con el `.form.txt` ordinario; no se cuenta como formulario (D-CON-3).
- **Distintos del snapshot:** 15 formularios, 27 módulos y 19 clases; `WorkflowServicio` (4 160 líneas, 82 procedimientos) es el mayor.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub`, `Function` y `Property` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide tras la normalización.

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Form_frm0OtrosAdmin | src/forms/Form_frm0OtrosAdmin.cls | 65 | 30 | 5 | igual | analizado | [formularios/Form_frm0OtrosAdmin.md](formularios/Form_frm0OtrosAdmin.md) |
| Form_frm0Ppal | src/forms/Form_frm0Ppal.cls | 322 | 39 | 17 | igual | analizado | [formularios/Form_frm0Ppal.md](formularios/Form_frm0Ppal.md) |
| Form_frm0PpalTecnico | src/forms/Form_frm0PpalTecnico.cls | 196 | 31 | 10 | igual | analizado | [formularios/Form_frm0PpalTecnico.md](formularios/Form_frm0PpalTecnico.md) |
| Form_frmAltaSolicitud | src/forms/Form_frmAltaSolicitud.cls | 646 | 43 | 19 | distinto | analizado | [formularios/Form_frmAltaSolicitud.md](formularios/Form_frmAltaSolicitud.md) |
| Form_frmBuscarExpediente | src/forms/Form_frmBuscarExpediente.cls | 242 | 31 | 14 | igual | analizado | [formularios/Form_frmBuscarExpediente.md](formularios/Form_frmBuscarExpediente.md) |
| Form_frmBuscarSolicitudes | src/forms/Form_frmBuscarSolicitudes.cls | 359 | 32 | 18 | distinto | analizado | [formularios/Form_frmBuscarSolicitudes.md](formularios/Form_frmBuscarSolicitudes.md) |
| Form_frmDatosCDCA | src/forms/Form_frmDatosCDCA.cls | 735 | 33 | 20 | distinto | pendiente | — |
| Form_frmDatosCDCASUB | src/forms/Form_frmDatosCDCASUB.cls | 779 | 33 | 21 | distinto | pendiente | — |
| Form_frmDatosPC | src/forms/Form_frmDatosPC.cls | 811 | 33 | 20 | distinto | pendiente | — |
| Form_frmDetalleExpediente | src/forms/Form_frmDetalleExpediente.cls | 88 | 48 | 3 | igual | analizado | [formularios/Form_frmDetalleExpediente.md](formularios/Form_frmDetalleExpediente.md) |
| Form_frmDetalleLogCambio | src/forms/Form_frmDetalleLogCambio.cls | 59 | 31 | 2 | igual | pendiente | — |
| Form_frmDetalleLogError | src/forms/Form_frmDetalleLogError.cls | 59 | 29 | 2 | igual | pendiente | — |
| Form_frmElegirEtapaAdjunto | src/forms/Form_frmElegirEtapaAdjunto.cls | 67 | 22 | 6 | igual | pendiente | — |
| Form_frmFiltrosAvanzadosSolicitudes | src/forms/Form_frmFiltrosAvanzadosSolicitudes.cls | 134 | 33 | 11 | igual | analizado | [formularios/Form_frmFiltrosAvanzadosSolicitudes.md](formularios/Form_frmFiltrosAvanzadosSolicitudes.md) |
| Form_frmGestionAdjuntos | src/forms/Form_frmGestionAdjuntos.cls | 380 | 23 | 10 | distinto | pendiente | — |
| Form_frmGestionSolicitud | src/forms/Form_frmGestionSolicitud.cls | 2416 | 32 | 53 | distinto | pendiente | — |
| Form_frmSplash | src/forms/Form_frmSplash.cls | 78 | 15 | 2 | igual | analizado | [formularios/Form_frmSplash.md](formularios/Form_frmSplash.md) |
| Form_frmVerLogCambios | src/forms/Form_frmVerLogCambios.cls | 187 | 26 | 10 | igual | pendiente | — |
| Form_frmVerLogErrores | src/forms/Form_frmVerLogErrores.cls | 183 | 25 | 10 | igual | pendiente | — |
| Form_frmWebVisor | src/forms/Form_frmWebVisor.cls | 426 | 21 | 15 | igual | pendiente | — |
| Form_subfrmDatosCDCASUB_AprobacionSuministrador | src/forms/Form_subfrmDatosCDCASUB_AprobacionSuministrador.cls | 198 | 58 | 13 | distinto | pendiente | — |
| Form_subfrmDatosCDCASUB_DecisionFinal | src/forms/Form_subfrmDatosCDCASUB_DecisionFinal.cls | 78 | 27 | 3 | distinto | pendiente | — |
| Form_subfrmDatosCDCASUB_DictamenRAC | src/forms/Form_subfrmDatosCDCASUB_DictamenRAC.cls | 133 | 30 | 5 | igual | pendiente | — |
| Form_subfrmDatosCDCASUB_Generales | src/forms/Form_subfrmDatosCDCASUB_Generales.cls | 205 | 37 | 6 | igual | pendiente | — |
| Form_subfrmDatosCDCASUB_Impacto | src/forms/Form_subfrmDatosCDCASUB_Impacto.cls | 99 | 67 | 3 | igual | pendiente | — |
| Form_subfrmDatosCDCASUB_Propuesta | src/forms/Form_subfrmDatosCDCASUB_Propuesta.cls | 27 | 31 | 3 | igual | pendiente | — |
| Form_subfrmDatosCDCA_AprobacionSuministrador | src/forms/Form_subfrmDatosCDCA_AprobacionSuministrador.cls | 224 | 58 | 13 | distinto | pendiente | — |
| Form_subfrmDatosCDCA_DecisionFinal | src/forms/Form_subfrmDatosCDCA_DecisionFinal.cls | 99 | 27 | 3 | distinto | pendiente | — |
| Form_subfrmDatosCDCA_DictamenRAC | src/forms/Form_subfrmDatosCDCA_DictamenRAC.cls | 120 | 25 | 5 | distinto | pendiente | — |
| Form_subfrmDatosCDCA_Generales | src/forms/Form_subfrmDatosCDCA_Generales.cls | 209 | 35 | 6 | igual | pendiente | — |
| Form_subfrmDatosCDCA_Impacto | src/forms/Form_subfrmDatosCDCA_Impacto.cls | 129 | 67 | 3 | igual | pendiente | — |
| Form_subfrmDatosCDCA_Propuesta | src/forms/Form_subfrmDatosCDCA_Propuesta.cls | 56 | 31 | 3 | igual | pendiente | — |
| Form_subfrmDatosPC_AprobacionSuministrador | src/forms/Form_subfrmDatosPC_AprobacionSuministrador.cls | 150 | 25 | 7 | distinto | pendiente | — |
| Form_subfrmDatosPC_DecisionFinal | src/forms/Form_subfrmDatosPC_DecisionFinal.cls | 122 | 32 | 3 | distinto | pendiente | — |
| Form_subfrmDatosPC_DictamenRAC | src/forms/Form_subfrmDatosPC_DictamenRAC.cls | 122 | 25 | 5 | igual | pendiente | — |
| Form_subfrmDatosPC_Generales | src/forms/Form_subfrmDatosPC_Generales.cls | 206 | 30 | 6 | igual | pendiente | — |
| Form_subfrmDatosPC_Impacto | src/forms/Form_subfrmDatosPC_Impacto.cls | 130 | 84 | 3 | distinto | pendiente | — |
| Form_subfrmDatosPC_Propuesta | src/forms/Form_subfrmDatosPC_Propuesta.cls | 34 | 24 | 3 | igual | pendiente | — |

### Módulos estándar

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| AdjuntoRepositorio | src/modules/AdjuntoRepositorio.bas | 221 | — | 8 | igual | pendiente | — |
| AplicacionRepositorio | src/modules/AplicacionRepositorio.bas | 60 | — | 1 | distinto | pendiente | — |
| ChecklistHelper | src/modules/ChecklistHelper.bas | 70 | — | 5 | igual | pendiente | — |
| CorreoRepositorio | src/modules/CorreoRepositorio.bas | 90 | — | 2 | distinto | pendiente | — |
| DatosCDCARepositorio | src/modules/DatosCDCARepositorio.bas | 455 | — | 13 | distinto | pendiente | — |
| DatosCDCASUBRepositorio | src/modules/DatosCDCASUBRepositorio.bas | 410 | — | 13 | distinto | pendiente | — |
| DatosPCRepositorio | src/modules/DatosPCRepositorio.bas | 545 | — | 13 | distinto | pendiente | — |
| ErrorLogger | src/modules/ErrorLogger.bas | 53 | — | 2 | distinto | pendiente | — |
| ErrorPresenter | src/modules/ErrorPresenter.bas | 186 | — | 5 | igual | pendiente | — |
| EstadoRepositorio | src/modules/EstadoRepositorio.bas | 27 | — | 1 | igual | pendiente | — |
| ExpedienteRepositorio | src/modules/ExpedienteRepositorio.bas | 107 | — | 3 | distinto | pendiente | — |
| FUNCIONES UTILES | src/modules/FUNCIONES UTILES.bas | 1035 | — | 40 | distinto | pendiente | — |
| Factoria | src/modules/Factoria.bas | 65 | — | 1 | distinto | pendiente | — |
| FormulariosPadreAuxiliares | src/modules/FormulariosPadreAuxiliares.bas | 273 | — | 6 | distinto | pendiente | — |
| HashHelper | src/modules/HashHelper.bas | 82 | — | 1 | igual | pendiente | — |
| JsonConverter | src/modules/JsonConverter.bas | 1447 | — | 32 | distinto | pendiente | — |
| JsonHelper | src/modules/JsonHelper.bas | 119 | — | 4 | igual | pendiente | — |
| LogCambioRepositorio | src/modules/LogCambioRepositorio.bas | 123 | — | 5 | distinto | pendiente | — |
| LogErrorRepositorio | src/modules/LogErrorRepositorio.bas | 106 | — | 4 | distinto | pendiente | — |
| LogEstadoRepositorio | src/modules/LogEstadoRepositorio.bas | 134 | — | 4 | igual | pendiente | — |
| MapeoRepositorio | src/modules/MapeoRepositorio.bas | 26 | — | 1 | distinto | pendiente | — |
| Módulo1 | src/modules/Módulo1.bas | 4 | — | 0 | solo en commit | pendiente | — |
| NoConformidadRepositorio | src/modules/NoConformidadRepositorio.bas | 233 | — | 8 | distinto | pendiente | — |
| Pruebas | src/modules/Pruebas.bas | 55 | — | 2 | distinto | pendiente | — |
| RechazoRepositorio | src/modules/RechazoRepositorio.bas | 326 | — | 9 | igual | pendiente | — |
| RepositorioComun | src/modules/RepositorioComun.bas | 201 | — | 6 | igual | pendiente | — |
| ResultadoImportacionRAC | src/modules/ResultadoImportacionRAC.bas | 10 | — | 0 | solo en commit | pendiente | — |
| SandboxBuilder | src/modules/SandboxBuilder.bas | 618 | — | 16 | distinto | pendiente | — |
| SandboxProvisioner | src/modules/SandboxProvisioner.bas | 1003 | — | 17 | distinto | pendiente | — |
| SnapshotHelper | src/modules/SnapshotHelper.bas | 701 | — | 13 | distinto | pendiente | — |
| SolicitudRepositorio | src/modules/SolicitudRepositorio.bas | 217 | — | 7 | igual | pendiente | — |
| SuministradorRepositorio | src/modules/SuministradorRepositorio.bas | 28 | — | 1 | distinto | pendiente | — |
| TestSandbox | src/modules/TestSandbox.bas | 591 | — | 21 | distinto | pendiente | — |
| UsuarioRepositorio | src/modules/UsuarioRepositorio.bas | 66 | — | 2 | distinto | pendiente | — |
| ValidacionRevisionRepositorio | src/modules/ValidacionRevisionRepositorio.bas | 606 | — | 16 | distinto | pendiente | — |
| Variables Globales | src/modules/Variables Globales.bas | 177 | — | 3 | distinto | pendiente | — |
| WorkflowRepositorio | src/modules/WorkflowRepositorio.bas | 58 | — | 1 | igual | pendiente | — |
| modActualizaciones | src/modules/modActualizaciones.bas | 58 | — | 2 | igual | pendiente | — |
| modBattery_Canonical | src/modules/modBattery_Canonical.bas | 7813 | — | 166 | distinto | pendiente | — |
| modDevTools | src/modules/modDevTools.bas | 117 | — | 3 | distinto | pendiente | — |
| modEnumeradores | src/modules/modEnumeradores.bas | 55 | — | 0 | distinto | pendiente | — |
| modSimuladorNotificaciones | src/modules/modSimuladorNotificaciones.bas | 107 | — | 10 | igual | pendiente | — |

### Clases

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Adjunto | src/classes/Adjunto.cls | 52 | — | 3 | igual | pendiente | — |
| AdjuntoViewModel | src/classes/AdjuntoViewModel.cls | 39 | — | 2 | igual | pendiente | — |
| AdjuntosServicio | src/classes/AdjuntosServicio.cls | 541 | — | 14 | distinto | pendiente | — |
| CondorError | src/classes/CondorError.cls | 103 | — | 9 | distinto | pendiente | — |
| Correo | src/classes/Correo.cls | 60 | — | 3 | igual | pendiente | — |
| DatosCDCA | src/classes/DatosCDCA.cls | 130 | — | 3 | igual | pendiente | — |
| DatosCDCASUB | src/classes/DatosCDCASUB.cls | 149 | — | 3 | igual | pendiente | — |
| DatosCDCASUBServicio | src/classes/DatosCDCASUBServicio.cls | 730 | — | 24 | distinto | pendiente | — |
| DatosCDCASUBViewModel | src/classes/DatosCDCASUBViewModel.cls | 100 | — | 2 | distinto | pendiente | — |
| DatosCDCAServicio | src/classes/DatosCDCAServicio.cls | 909 | — | 25 | distinto | pendiente | — |
| DatosCDCAViewModel | src/classes/DatosCDCAViewModel.cls | 113 | — | 2 | distinto | pendiente | — |
| DatosPC | src/classes/DatosPC.cls | 129 | — | 3 | igual | pendiente | — |
| DatosPCServicio | src/classes/DatosPCServicio.cls | 887 | — | 24 | distinto | pendiente | — |
| DatosPCViewModel | src/classes/DatosPCViewModel.cls | 104 | — | 2 | distinto | pendiente | — |
| DocumentoServicio | src/classes/DocumentoServicio.cls | 1095 | — | 20 | distinto | pendiente | — |
| Entorno | src/classes/Entorno.cls | 645 | — | 33 | distinto | pendiente | — |
| Estado | src/classes/Estado.cls | 51 | — | 3 | igual | pendiente | — |
| EstadoServicio | src/classes/EstadoServicio.cls | 36 | — | 1 | igual | pendiente | — |
| Expediente | src/classes/Expediente.cls | 67 | — | 3 | igual | pendiente | — |
| ExpedienteServicio | src/classes/ExpedienteServicio.cls | 307 | — | 7 | distinto | pendiente | — |
| ExpedienteSuministrador | src/classes/ExpedienteSuministrador.cls | 51 | — | 3 | igual | pendiente | — |
| ExpedienteViewModel | src/classes/ExpedienteViewModel.cls | 36 | — | 2 | igual | pendiente | — |
| FiltrosSolicitud | src/classes/FiltrosSolicitud.cls | 34 | — | 1 | igual | pendiente | — |
| LogCambio | src/classes/LogCambio.cls | 58 | — | 3 | igual | pendiente | — |
| LogCambioServicio | src/classes/LogCambioServicio.cls | 33 | — | 1 | igual | pendiente | — |
| LogError | src/classes/LogError.cls | 55 | — | 3 | igual | pendiente | — |
| LogEstado | src/classes/LogEstado.cls | 52 | — | 3 | igual | pendiente | — |
| LogServicio | src/classes/LogServicio.cls | 225 | — | 10 | igual | pendiente | — |
| MapeoCampos | src/classes/MapeoCampos.cls | 49 | — | 3 | igual | pendiente | — |
| MapeoServicio | src/classes/MapeoServicio.cls | 88 | — | 4 | distinto | pendiente | — |
| NoConformidad | src/classes/NoConformidad.cls | 46 | — | 3 | igual | pendiente | — |
| NoConformidadServicio | src/classes/NoConformidadServicio.cls | 105 | — | 7 | igual | pendiente | — |
| NotificacionServicio | src/classes/NotificacionServicio.cls | 506 | — | 11 | distinto | pendiente | — |
| Rechazo | src/classes/Rechazo.cls | 20 | — | 0 | igual | pendiente | — |
| RechazoServicio | src/classes/RechazoServicio.cls | 72 | — | 3 | igual | pendiente | — |
| ResultadoImportacionRAC | src/classes/ResultadoImportacionRAC.cls | 18 | — | 0 | solo en commit | pendiente | — |
| RevisionServicio | src/classes/RevisionServicio.cls | 161 | — | 3 | igual | pendiente | — |
| SnapshotServicio | src/classes/SnapshotServicio.cls | 103 | — | 4 | distinto | pendiente | — |
| Solicitud | src/classes/Solicitud.cls | 64 | — | 3 | igual | pendiente | — |
| SolicitudBusquedaViewModel | src/classes/SolicitudBusquedaViewModel.cls | 65 | — | 3 | igual | pendiente | — |
| SolicitudServicio | src/classes/SolicitudServicio.cls | 741 | — | 16 | distinto | pendiente | — |
| SolicitudViewModel | src/classes/SolicitudViewModel.cls | 54 | — | 2 | igual | pendiente | — |
| Suministrador | src/classes/Suministrador.cls | 66 | — | 4 | igual | pendiente | — |
| SuministradorServicio | src/classes/SuministradorServicio.cls | 59 | — | 2 | distinto | pendiente | — |
| Usuario | src/classes/Usuario.cls | 92 | — | 3 | igual | pendiente | — |
| UsuarioAplicacionPermisos | src/classes/UsuarioAplicacionPermisos.cls | 55 | — | 3 | igual | pendiente | — |
| UsuarioServicio | src/classes/UsuarioServicio.cls | 149 | — | 5 | distinto | pendiente | — |
| ValidacionRevision | src/classes/ValidacionRevision.cls | 54 | — | 3 | igual | pendiente | — |
| ValidacionRevisionServicio | src/classes/ValidacionRevisionServicio.cls | 99 | — | 7 | igual | pendiente | — |
| WebVisorCacheServicio | src/classes/WebVisorCacheServicio.cls | 180 | — | 9 | distinto | pendiente | — |
| WorkflowServicio | src/classes/WorkflowServicio.cls | 4160 | — | 82 | distinto | pendiente | — |

Totales: 9 analizados · 0 no aplicables · 122 pendientes.

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-CON-1 | El estudio previo da cuatro tipos de solicitud (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`); en el commit faltan los siete formularios de `PC_SUB` | [epic.md](../epic.md) F2; inventario | pendiente |
| D-CON-2 | Faltan en el commit 11 módulos no de prueba, 7 clases y 23 módulos `Test_*` del snapshot | Inventario | pendiente |
| D-CON-3 | Variante `form-admCasa.txt` de `frmGestionSolicitud` | `src/forms/Form_frmGestionSolicitud.form-admCasa.txt` | pendiente |
| D-CON-4 | La issue #1075 fecha el commit el 2026-10-21; el commit pertenece al 2026-07-18 | `git log` del commit fijado | informativo |
| D-CON-5 | ERD de origen: solo existe la documentación manual del legacy; falta la extracción de solo lectura del `.accdb` | `ERD/Estructura_Datos.md` del legacy | pendiente del propietario |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/condor/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.
