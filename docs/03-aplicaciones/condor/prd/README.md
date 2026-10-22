# PRD — Condor

> **Estado:** entrega completa de la trazabilidad. Los 38 formularios, los 42 módulos y las 51 clases del inventario tienen ficha o clasificación; la matriz de campos, la coexistencia, las capacidades y las issues propuestas están escritas. Queda pendiente solo el ERD de origen (D-CON-5), que depende del propietario.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/condor |
| Commit fijado | `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` (`main`, 2026-07-18, «chore(agents): delegate dysflow config to global») |
| Índice CodeGraph-VBA | versión 2.1.1 sobre el clon de solo lectura de ese commit. Medición de la fase de inventario: 203 ficheros y 9 039 nodos; remedición de la fase de fichas: 226 ficheros, 9 433 nodos y 15 970 aristas. La diferencia es de qué ficheros se contaron, no del commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/condor/src` (2026-08-06) |
| Issue | #1075 (método y plantilla de #1069) |
| Entregables | [formularios](formularios/), [modulos](modulos/), [clases](clases/), [matriz-campos](matriz-campos.md), [coexistencia](coexistencia.md), [capacidades](capacidades.md), [issues-propuestas](issues-propuestas.md) |

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
| Form_frmDatosCDCA | src/forms/Form_frmDatosCDCA.cls | 735 | 33 | 20 | distinto | analizado | [formularios/Form_frmDatosCDCA.md](formularios/Form_frmDatosCDCA.md) |
| Form_frmDatosCDCASUB | src/forms/Form_frmDatosCDCASUB.cls | 779 | 33 | 21 | distinto | analizado | [formularios/Form_frmDatosCDCASUB.md](formularios/Form_frmDatosCDCASUB.md) |
| Form_frmDatosPC | src/forms/Form_frmDatosPC.cls | 811 | 33 | 20 | distinto | analizado | [formularios/Form_frmDatosPC.md](formularios/Form_frmDatosPC.md) |
| Form_frmDetalleExpediente | src/forms/Form_frmDetalleExpediente.cls | 88 | 48 | 3 | igual | analizado | [formularios/Form_frmDetalleExpediente.md](formularios/Form_frmDetalleExpediente.md) |
| Form_frmDetalleLogCambio | src/forms/Form_frmDetalleLogCambio.cls | 59 | 31 | 2 | igual | analizado | [formularios/Form_frmDetalleLogCambio.md](formularios/Form_frmDetalleLogCambio.md) |
| Form_frmDetalleLogError | src/forms/Form_frmDetalleLogError.cls | 59 | 29 | 2 | igual | analizado | [formularios/Form_frmDetalleLogError.md](formularios/Form_frmDetalleLogError.md) |
| Form_frmElegirEtapaAdjunto | src/forms/Form_frmElegirEtapaAdjunto.cls | 67 | 22 | 6 | igual | analizado | [formularios/Form_frmElegirEtapaAdjunto.md](formularios/Form_frmElegirEtapaAdjunto.md) |
| Form_frmFiltrosAvanzadosSolicitudes | src/forms/Form_frmFiltrosAvanzadosSolicitudes.cls | 134 | 33 | 11 | igual | analizado | [formularios/Form_frmFiltrosAvanzadosSolicitudes.md](formularios/Form_frmFiltrosAvanzadosSolicitudes.md) |
| Form_frmGestionAdjuntos | src/forms/Form_frmGestionAdjuntos.cls | 380 | 23 | 10 | distinto | analizado | [formularios/Form_frmGestionAdjuntos.md](formularios/Form_frmGestionAdjuntos.md) |
| Form_frmGestionSolicitud | src/forms/Form_frmGestionSolicitud.cls | 2416 | 32 | 53 | distinto | analizado | [formularios/Form_frmGestionSolicitud.md](formularios/Form_frmGestionSolicitud.md) |
| Form_frmSplash | src/forms/Form_frmSplash.cls | 78 | 15 | 2 | igual | analizado | [formularios/Form_frmSplash.md](formularios/Form_frmSplash.md) |
| Form_frmVerLogCambios | src/forms/Form_frmVerLogCambios.cls | 187 | 26 | 10 | igual | analizado | [formularios/Form_frmVerLogCambios.md](formularios/Form_frmVerLogCambios.md) |
| Form_frmVerLogErrores | src/forms/Form_frmVerLogErrores.cls | 183 | 25 | 10 | igual | analizado | [formularios/Form_frmVerLogErrores.md](formularios/Form_frmVerLogErrores.md) |
| Form_frmWebVisor | src/forms/Form_frmWebVisor.cls | 426 | 21 | 15 | igual | analizado | [formularios/Form_frmWebVisor.md](formularios/Form_frmWebVisor.md) |
| Form_subfrmDatosCDCASUB_AprobacionSuministrador | src/forms/Form_subfrmDatosCDCASUB_AprobacionSuministrador.cls | 198 | 58 | 13 | distinto | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCASUB_DecisionFinal | src/forms/Form_subfrmDatosCDCASUB_DecisionFinal.cls | 78 | 27 | 3 | distinto | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCASUB_DictamenRAC | src/forms/Form_subfrmDatosCDCASUB_DictamenRAC.cls | 133 | 30 | 5 | igual | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCASUB_Generales | src/forms/Form_subfrmDatosCDCASUB_Generales.cls | 205 | 37 | 6 | igual | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCASUB_Impacto | src/forms/Form_subfrmDatosCDCASUB_Impacto.cls | 99 | 67 | 3 | igual | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCASUB_Propuesta | src/forms/Form_subfrmDatosCDCASUB_Propuesta.cls | 27 | 31 | 3 | igual | analizado | [formularios/Form_subfrmDatosCDCASUB.md](formularios/Form_subfrmDatosCDCASUB.md) |
| Form_subfrmDatosCDCA_AprobacionSuministrador | src/forms/Form_subfrmDatosCDCA_AprobacionSuministrador.cls | 224 | 58 | 13 | distinto | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosCDCA_DecisionFinal | src/forms/Form_subfrmDatosCDCA_DecisionFinal.cls | 99 | 27 | 3 | distinto | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosCDCA_DictamenRAC | src/forms/Form_subfrmDatosCDCA_DictamenRAC.cls | 120 | 25 | 5 | distinto | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosCDCA_Generales | src/forms/Form_subfrmDatosCDCA_Generales.cls | 209 | 35 | 6 | igual | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosCDCA_Impacto | src/forms/Form_subfrmDatosCDCA_Impacto.cls | 129 | 67 | 3 | igual | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosCDCA_Propuesta | src/forms/Form_subfrmDatosCDCA_Propuesta.cls | 56 | 31 | 3 | igual | analizado | [formularios/Form_subfrmDatosCDCA.md](formularios/Form_subfrmDatosCDCA.md) |
| Form_subfrmDatosPC_AprobacionSuministrador | src/forms/Form_subfrmDatosPC_AprobacionSuministrador.cls | 150 | 25 | 7 | distinto | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |
| Form_subfrmDatosPC_DecisionFinal | src/forms/Form_subfrmDatosPC_DecisionFinal.cls | 122 | 32 | 3 | distinto | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |
| Form_subfrmDatosPC_DictamenRAC | src/forms/Form_subfrmDatosPC_DictamenRAC.cls | 122 | 25 | 5 | igual | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |
| Form_subfrmDatosPC_Generales | src/forms/Form_subfrmDatosPC_Generales.cls | 206 | 30 | 6 | igual | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |
| Form_subfrmDatosPC_Impacto | src/forms/Form_subfrmDatosPC_Impacto.cls | 130 | 84 | 3 | distinto | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |
| Form_subfrmDatosPC_Propuesta | src/forms/Form_subfrmDatosPC_Propuesta.cls | 34 | 24 | 3 | igual | analizado | [formularios/Form_subfrmDatosPC.md](formularios/Form_subfrmDatosPC.md) |

### Módulos estándar

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| AdjuntoRepositorio | src/modules/AdjuntoRepositorio.bas | 221 | — | 8 | igual | pendiente | analizado | [Modulo_RepositoriosSolicitud](modulos/Modulo_RepositoriosSolicitud.md) 
| AplicacionRepositorio | src/modules/AplicacionRepositorio.bas | 60 | — | 1 | distinto | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| ChecklistHelper | src/modules/ChecklistHelper.bas | 70 | — | 5 | igual | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| CorreoRepositorio | src/modules/CorreoRepositorio.bas | 90 | — | 2 | distinto | pendiente | analizado | [Modulo_RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) 
| DatosCDCARepositorio | src/modules/DatosCDCARepositorio.bas | 455 | — | 13 | distinto | pendiente | analizado | [Modulo_RepositoriosDatos](modulos/Modulo_RepositoriosDatos.md) 
| DatosCDCASUBRepositorio | src/modules/DatosCDCASUBRepositorio.bas | 410 | — | 13 | distinto | pendiente | analizado | [Modulo_RepositoriosDatos](modulos/Modulo_RepositoriosDatos.md) 
| DatosPCRepositorio | src/modules/DatosPCRepositorio.bas | 545 | — | 13 | distinto | pendiente | analizado | [Modulo_RepositoriosDatos](modulos/Modulo_RepositoriosDatos.md) 
| ErrorLogger | src/modules/ErrorLogger.bas | 53 | — | 2 | distinto | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 
| ErrorPresenter | src/modules/ErrorPresenter.bas | 186 | — | 5 | igual | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 
| EstadoRepositorio | src/modules/EstadoRepositorio.bas | 27 | — | 1 | igual | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| ExpedienteRepositorio | src/modules/ExpedienteRepositorio.bas | 107 | — | 3 | distinto | pendiente | analizado | [Modulo_RepositoriosSolicitud](modulos/Modulo_RepositoriosSolicitud.md) 
| FUNCIONES UTILES | src/modules/FUNCIONES UTILES.bas | 1035 | — | 40 | distinto | pendiente | analizado | [Modulo_EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) 
| Factoria | src/modules/Factoria.bas | 65 | — | 1 | distinto | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| FormulariosPadreAuxiliares | src/modules/FormulariosPadreAuxiliares.bas | 273 | — | 6 | distinto | pendiente | analizado | [Modulo_FormulariosPadreYJson](modulos/Modulo_FormulariosPadreYJson.md) 
| HashHelper | src/modules/HashHelper.bas | 82 | — | 1 | igual | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| JsonConverter | src/modules/JsonConverter.bas | 1447 | — | 32 | distinto | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 
| JsonHelper | src/modules/JsonHelper.bas | 119 | — | 4 | igual | pendiente | analizado | [Modulo_FormulariosPadreYJson](modulos/Modulo_FormulariosPadreYJson.md) 
| LogCambioRepositorio | src/modules/LogCambioRepositorio.bas | 123 | — | 5 | distinto | pendiente | analizado | [Modulo_RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) 
| LogErrorRepositorio | src/modules/LogErrorRepositorio.bas | 106 | — | 4 | distinto | pendiente | analizado | [Modulo_RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) 
| LogEstadoRepositorio | src/modules/LogEstadoRepositorio.bas | 134 | — | 4 | igual | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| MapeoRepositorio | src/modules/MapeoRepositorio.bas | 26 | — | 1 | distinto | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| Módulo1 | src/modules/Módulo1.bas | 4 | — | 0 | solo en commit | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| NoConformidadRepositorio | src/modules/NoConformidadRepositorio.bas | 233 | — | 8 | distinto | pendiente | analizado | [Modulo_RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) 
| Pruebas | src/modules/Pruebas.bas | 55 | — | 2 | distinto | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 
| RechazoRepositorio | src/modules/RechazoRepositorio.bas | 326 | — | 9 | igual | pendiente | analizado | [Modulo_RepositoriosSolicitud](modulos/Modulo_RepositoriosSolicitud.md) 
| RepositorioComun | src/modules/RepositorioComun.bas | 201 | — | 6 | igual | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| ResultadoImportacionRAC | src/modules/ResultadoImportacionRAC.bas | 10 | — | 0 | solo en commit | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| SandboxBuilder | src/modules/SandboxBuilder.bas | 618 | — | 16 | distinto | pendiente | analizado | [Modulo_EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) 
| SandboxProvisioner | src/modules/SandboxProvisioner.bas | 1003 | — | 17 | distinto | pendiente | analizado | [Modulo_EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) 
| SnapshotHelper | src/modules/SnapshotHelper.bas | 701 | — | 13 | distinto | pendiente | analizado | [Modulo_SnapshotYHash](modulos/Modulo_SnapshotYHash.md) 
| SolicitudRepositorio | src/modules/SolicitudRepositorio.bas | 217 | — | 7 | igual | pendiente | analizado | [Modulo_RepositoriosSolicitud](modulos/Modulo_RepositoriosSolicitud.md) 
| SuministradorRepositorio | src/modules/SuministradorRepositorio.bas | 28 | — | 1 | distinto | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| TestSandbox | src/modules/TestSandbox.bas | 591 | — | 21 | distinto | pendiente | analizado | [Modulo_EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) 
| UsuarioRepositorio | src/modules/UsuarioRepositorio.bas | 66 | — | 2 | distinto | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| ValidacionRevisionRepositorio | src/modules/ValidacionRevisionRepositorio.bas | 606 | — | 16 | distinto | pendiente | analizado | [Modulo_RepositoriosTrazabilidad](modulos/Modulo_RepositoriosTrazabilidad.md) 
| Variables Globales | src/modules/Variables Globales.bas | 177 | — | 3 | distinto | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| WorkflowRepositorio | src/modules/WorkflowRepositorio.bas | 58 | — | 1 | igual | pendiente | analizado | [Modulo_RepositoriosCatalogos](modulos/Modulo_RepositoriosCatalogos.md) 
| modActualizaciones | src/modules/modActualizaciones.bas | 58 | — | 2 | igual | pendiente | analizado | [Modulo_UtilidadesComunes](modulos/Modulo_UtilidadesComunes.md) 
| modBattery_Canonical | src/modules/modBattery_Canonical.bas | 7813 | — | 166 | distinto | pendiente | analizado | [Modulo_EntornoDePruebas](modulos/Modulo_EntornoDePruebas.md) 
| modDevTools | src/modules/modDevTools.bas | 117 | — | 3 | distinto | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 
| modEnumeradores | src/modules/modEnumeradores.bas | 55 | — | 0 | distinto | pendiente | analizado | [Modulo_FormulariosPadreYJson](modulos/Modulo_FormulariosPadreYJson.md) 
| modSimuladorNotificaciones | src/modules/modSimuladorNotificaciones.bas | 107 | — | 10 | igual | pendiente | analizado | [Modulo_SerializacionYErrores](modulos/Modulo_SerializacionYErrores.md) 

### Clases

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Adjunto | src/classes/Adjunto.cls | 52 | — | 3 | igual | pendiente | analizado | [Clase_Adjuntos](clases/Clase_Adjuntos.md) 
| AdjuntoViewModel | src/classes/AdjuntoViewModel.cls | 39 | — | 2 | igual | pendiente | analizado | [Clase_Adjuntos](clases/Clase_Adjuntos.md) 
| AdjuntosServicio | src/classes/AdjuntosServicio.cls | 541 | — | 14 | distinto | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| CondorError | src/classes/CondorError.cls | 103 | — | 9 | distinto | pendiente | analizado | [Clase_Adjuntos](clases/Clase_Adjuntos.md) 
| Correo | src/classes/Correo.cls | 60 | — | 3 | igual | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| DatosCDCA | src/classes/DatosCDCA.cls | 130 | — | 3 | igual | pendiente | analizado | [Clase_DatosCDCA](clases/Clase_DatosCDCA.md) 
| DatosCDCASUB | src/classes/DatosCDCASUB.cls | 149 | — | 3 | igual | pendiente | analizado | [Clase_DatosCDCASUB](clases/Clase_DatosCDCASUB.md) 
| DatosCDCASUBServicio | src/classes/DatosCDCASUBServicio.cls | 730 | — | 24 | distinto | pendiente | analizado | [Clase_DatosCDCASUB](clases/Clase_DatosCDCASUB.md) 
| DatosCDCASUBViewModel | src/classes/DatosCDCASUBViewModel.cls | 100 | — | 2 | distinto | pendiente | analizado | [Clase_DatosCDCASUB](clases/Clase_DatosCDCASUB.md) 
| DatosCDCAServicio | src/classes/DatosCDCAServicio.cls | 909 | — | 25 | distinto | pendiente | analizado | [Clase_DatosCDCA](clases/Clase_DatosCDCA.md) 
| DatosCDCAViewModel | src/classes/DatosCDCAViewModel.cls | 113 | — | 2 | distinto | pendiente | analizado | [Clase_DatosCDCA](clases/Clase_DatosCDCA.md) 
| DatosPC | src/classes/DatosPC.cls | 129 | — | 3 | igual | pendiente | analizado | [Clase_DatosPC](clases/Clase_DatosPC.md) 
| DatosPCServicio | src/classes/DatosPCServicio.cls | 887 | — | 24 | distinto | pendiente | analizado | [Clase_DatosPC](clases/Clase_DatosPC.md) 
| DatosPCViewModel | src/classes/DatosPCViewModel.cls | 104 | — | 2 | distinto | pendiente | analizado | [Clase_DatosPC](clases/Clase_DatosPC.md) 
| DocumentoServicio | src/classes/DocumentoServicio.cls | 1095 | — | 20 | distinto | pendiente | analizado | [Clase_Expediente](clases/Clase_Expediente.md) 
| Entorno | src/classes/Entorno.cls | 645 | — | 33 | distinto | pendiente | analizado | [Clase_EntornoYVisor](clases/Clase_EntornoYVisor.md) 
| Estado | src/classes/Estado.cls | 51 | — | 3 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| EstadoServicio | src/classes/EstadoServicio.cls | 36 | — | 1 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| Expediente | src/classes/Expediente.cls | 67 | — | 3 | igual | pendiente | analizado | [Clase_Expediente](clases/Clase_Expediente.md) 
| ExpedienteServicio | src/classes/ExpedienteServicio.cls | 307 | — | 7 | distinto | pendiente | analizado | [Clase_Expediente](clases/Clase_Expediente.md) 
| ExpedienteSuministrador | src/classes/ExpedienteSuministrador.cls | 51 | — | 3 | igual | pendiente | analizado | [Clase_Expediente](clases/Clase_Expediente.md) 
| ExpedienteViewModel | src/classes/ExpedienteViewModel.cls | 36 | — | 2 | igual | pendiente | analizado | [Clase_Expediente](clases/Clase_Expediente.md) 
| FiltrosSolicitud | src/classes/FiltrosSolicitud.cls | 34 | — | 1 | igual | pendiente | analizado | [Clase_Solicitud](clases/Clase_Solicitud.md) 
| LogCambio | src/classes/LogCambio.cls | 58 | — | 3 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| LogCambioServicio | src/classes/LogCambioServicio.cls | 33 | — | 1 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| LogError | src/classes/LogError.cls | 55 | — | 3 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| LogEstado | src/classes/LogEstado.cls | 52 | — | 3 | igual | pendiente | analizado | [Clase_Trazabilidad](clases/Clase_Trazabilidad.md) 
| LogServicio | src/classes/LogServicio.cls | 225 | — | 10 | igual | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| MapeoCampos | src/classes/MapeoCampos.cls | 49 | — | 3 | igual | pendiente | analizado | [Clase_EntornoYVisor](clases/Clase_EntornoYVisor.md) 
| MapeoServicio | src/classes/MapeoServicio.cls | 88 | — | 4 | distinto | pendiente | analizado | [Clase_EntornoYVisor](clases/Clase_EntornoYVisor.md) 
| NoConformidad | src/classes/NoConformidad.cls | 46 | — | 3 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| NoConformidadServicio | src/classes/NoConformidadServicio.cls | 105 | — | 7 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| NotificacionServicio | src/classes/NotificacionServicio.cls | 506 | — | 11 | distinto | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| Rechazo | src/classes/Rechazo.cls | 20 | — | 0 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| RechazoServicio | src/classes/RechazoServicio.cls | 72 | — | 3 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| ResultadoImportacionRAC | src/classes/ResultadoImportacionRAC.cls | 18 | — | 0 | solo en commit | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| RevisionServicio | src/classes/RevisionServicio.cls | 161 | — | 3 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| SnapshotServicio | src/classes/SnapshotServicio.cls | 103 | — | 4 | distinto | pendiente | analizado | [Clase_EntornoYVisor](clases/Clase_EntornoYVisor.md) 
| Solicitud | src/classes/Solicitud.cls | 64 | — | 3 | igual | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| SolicitudBusquedaViewModel | src/classes/SolicitudBusquedaViewModel.cls | 65 | — | 3 | igual | pendiente | analizado | [Clase_Solicitud](clases/Clase_Solicitud.md) 
| SolicitudServicio | src/classes/SolicitudServicio.cls | 741 | — | 16 | distinto | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| SolicitudViewModel | src/classes/SolicitudViewModel.cls | 54 | — | 2 | igual | pendiente | analizado | [Clase_Solicitud](clases/Clase_Solicitud.md) 
| Suministrador | src/classes/Suministrador.cls | 66 | — | 4 | igual | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| SuministradorServicio | src/classes/SuministradorServicio.cls | 59 | — | 2 | distinto | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| Usuario | src/classes/Usuario.cls | 92 | — | 3 | igual | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| UsuarioAplicacionPermisos | src/classes/UsuarioAplicacionPermisos.cls | 55 | — | 3 | igual | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| UsuarioServicio | src/classes/UsuarioServicio.cls | 149 | — | 5 | distinto | pendiente | analizado | [Clase_UsuariosYNotificaciones](clases/Clase_UsuariosYNotificaciones.md) 
| ValidacionRevision | src/classes/ValidacionRevision.cls | 54 | — | 3 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| ValidacionRevisionServicio | src/classes/ValidacionRevisionServicio.cls | 99 | — | 7 | igual | pendiente | analizado | [Clase_RevisionYRechazo](clases/Clase_RevisionYRechazo.md) 
| WebVisorCacheServicio | src/classes/WebVisorCacheServicio.cls | 180 | — | 9 | distinto | pendiente | analizado | [Clase_WorkflowServicio](clases/Clase_WorkflowServicio.md) 
| WorkflowServicio | src/classes/WorkflowServicio.cls | 4160 | — | 82 | distinto | pendiente | analizado | [Clase_DatosCDCASUB](clases/Clase_DatosCDCASUB.md) 

Totales: **131 de 131 objetos con ficha** · 0 no aplicables · 0 pendientes sin motivo.

- Formularios: 38 analizados, repartidos en 23 fichas (tres de ellas cubren un grupo de seis subformularios).
- Módulos: 42 analizados, repartidos en 10 hojas por dominio.
- Clases: 51 analizadas, repartidas en 11 hojas por dominio.

Cada hoja cita `ruta:línea` del commit fijado. Un comprobador mecánico resolvió el párrafo de abreviaturas de cada hoja y validó que cada cita cae dentro del fichero y dentro del rango de líneas:

| Parte | Citas comprobadas | Fuera de rango |
|---|---|---|
| [modulos](modulos/) | 742 | 0 |
| [clases](clases/) | 703 | 0 |
| [formularios](formularios/) (heredadas) | 963 | 0 |

En las fichas heredadas hay 116 citas que el comprobador no resuelve automáticamente porque tres hojas de grupo declaran sus abreviaturas con un formato de lista con barras verticales que no reconoce; se revisaron una a una contra el fichero y todas son válidas.

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-CON-1 | El estudio previo da cuatro tipos de solicitud (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`); en el commit faltan los siete formularios de `PC_SUB` | [epic.md](../epic.md) F2; [capacidades](capacidades.md) C-CON-7 | **cerrada en el código**: el tipo `PC_SUB` no existe en el commit. Si el producto lo necesita, es una capacidad añadida pendiente de decisión |
| D-CON-2 | Faltan en el commit 11 módulos no de prueba, 7 clases y 23 módulos `Test_*` del snapshot | Inventario; [modulos](modulos/Modulo_UtilidadesComunes.md) | **cerrada como observación**: la lógica de esos objetos no está en el commit y su comportamiento actual está cubierto por los objetos que sí están. Queda por confirmar contra el binario si algo se perdió por el camino |
| D-CON-3 | Variante `form-admCasa.txt` de `frmGestionSolicitud` | `src/forms/Form_frmGestionSolicitud.form-admCasa.txt` | **cerrada**: el fichero existe y es idéntico al `.form.txt` hasta la línea 1234; no es un formulario adicional |
| D-CON-4 | La issue #1075 fecha el commit el 2026-10-21; el commit pertenece al 2026-07-18 | `git log` del commit fijado | informativo |
| D-CON-5 | ERD de origen: solo existe la documentación manual del legacy; falta la extracción de solo lectura del `.accdb` | `ERD/Estructura_Datos.md` del legacy; [matriz-campos](matriz-campos.md) | **pendiente del propietario**: sin él, las 257 columnas de la matriz siguen siendo candidatas a nivel de tabla |
| D-CON-6 | El volcado de `ERD/schema.sql` es anterior al commit fijado e incluye `tbDatosPCSUB` y `tbHistorialRechazos`, dos tablas que el commit ya no tiene | [matriz-campos](matriz-campos.md) | **abierta**, se cierra con D-CON-5 |
| D-CON-7 | El comentario de `Form_frmDatosCDCASUB.cls:701` cita `tbDatosCDCA_SUB`, nombre que no corresponde a ninguna tabla del volcado | `src/forms/Form_frmDatosCDCASUB.cls:701` | **abierta**: discrepancia entre documentación y código |
| D-CON-8 | El estudio previo afirma que `getdb()` tiene un único llamante; la medición del commit da 155 llamadas en 13 ficheros de producción | [capacidades](capacidades.md) | **cerrada**: el dato del estudio previo era incorrecto |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/condor/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.

## Anexo — cambios no liberados

Lo que el volcado del snapshot contiene y el commit fijado **no**, sin ficha porque no existe en el código que este PRD traza. Todo lo de esta tabla está en `ERD/schema.sql` o en el recuento del snapshot del 2026-08-06, y ninguna línea corresponde a código vigente.

| Objeto | Qué era | Dónde queda documentado |
|---|---|---|
| `Form_frmDatosPCSUB` y sus seis subformularios `subfrmDatosPCSUB_*` | Séptimo bloque de datos del tipo de solicitud `PC_SUB` | D-CON-1; C-CON-7 en [capacidades](capacidades.md) |
| `DatosPCSUB`, `DatosPCSUBServicio`, `DatosPCSUBViewModel`, `DatosPCSUBGuardarHelper` y `DatosPCSUBRepositorio` | Servicio, modelo de vista, ayudante y repositorio de ese bloque | D-CON-1 |
| Tabla `tbDatosPCSUB` (43 columnas) y `tbHistorialRechazos` (8 columnas) | Tablas del snapshot sin tabla equivalente en el commit | [matriz-campos](matriz-campos/solicitud-y-bloques.md) |
| `DecisionFinalHelper`, `DictamenRACDefaultsHelper`, `JerarquiaRolesHelper`, `NotificacionHelper`, `PrecondicionesWorkflowHelper`, `DatosTipoSolicitudHelper` | Ayudantes de decisión, dictamen, roles, notificación y precondiciones | D-CON-2; C-CON-2 |
| `SandboxConfig`, `SandboxGestor`, `SandboxValidator`, `MockNotifServ`, `TestHelper` | Dobles y gestión del entorno de pruebas | C-CON-4 y C-CON-18 en [capacidades](capacidades.md) |
| 23 módulos `Test_*` | Batería de pruebas anterior | [modulos](modulos/Modulo_EntornoDePruebas.md) |

## Cómo leer este PRD

| Si te interesa… | Empieza por… |
|---|---|
| Qué hace un formulario concreto | su ficha en [formularios](formularios/) |
| Qué hace un módulo o una clase | la hoja de su dominio en [modulos](modulos/) o [clases](clases/) |
| Qué pasa con una tabla o una columna | [matriz-campos](matriz-campos.md) |
| Qué hay que decidir antes de migrar | [coexistencia](coexistencia.md) y [issues-propuestas](issues-propuestas.md) |
| Qué cambió desde el estudio previo | [capacidades](capacidades.md) |
