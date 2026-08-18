-- Schema dump: Gestion_Riesgos_Datos.accdb
-- Tables: 84

CREATE TABLE [ProyectoRespCalidad] (
  [Proyecto] TEXT(255),
  [UsuarioCalidad] TEXT(255)
);
-- Rows: 16

CREATE TABLE [TbAnexos] (
  [IDAnexo] LONG(4),
  [IDProyecto] LONG(4),
  [IDEdicion] LONG(4),
  [IDRiesgo] LONG(4),
  [FechaAnexo] SHORT_DATE_TIME(8),
  [Titulo] TEXT(255),
  [NombreArchivo] TEXT(255),
  [EvidenciaUTE] TEXT(2),
  [EvidenciaSuministrador] TEXT(2)
);
-- Rows: 484

CREATE TABLE [TbAnexosAntigua] (
  [IDAnexo] LONG(4),
  [IDProyecto] LONG(4),
  [IDEdicion] LONG(4),
  [FechaAnexo] SHORT_DATE_TIME(8),
  [CodigoUnico] TEXT(255),
  [Titulo] TEXT(255),
  [Descripcion] MEMO,
  [NombreArchivo] TEXT(255),
  [EvidenciaUTE] TEXT(2)
);
-- Rows: 2

-- LINKED/SKIPPED: [TbAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbArbolCacheMeta] (
  [IDEdicion] LONG(4),
  [CacheSetIdActivo] LONG(4),
  [BuiltAt] SHORT_DATE_TIME(8),
  [BuildMs] LONG(4),
  [Estado] TEXT(20)
);
-- Rows: 0

CREATE TABLE [TbAux] (
  [IDProyecto] LONG(4)
);
-- Rows: 13

CREATE TABLE [TbAuxProyectosRiesgos] (
  [IDProyecto] LONG(4),
  [Riesgo] TEXT(255),
  [FechaDetectado] SHORT_DATE_TIME(8),
  [NombreProyectoCompleto] TEXT(255),
  [FechaMaterializado] SHORT_DATE_TIME(8)
);
-- Rows: 2

CREATE TABLE [TbAuxSumExp] (
  [IdSumAux] LONG(4),
  [IDExp] LONG(4)
);
-- Rows: 6

CREATE TABLE [TbBibliotecaRiesgos] (
  [IDRiesgoTipo] LONG(4),
  [CODIGO] TEXT(255),
  [Tipo] TEXT(255),
  [Descripcion] MEMO,
  [Usuario] TEXT(255),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaModificacion] SHORT_DATE_TIME(8),
  [Familia] TEXT(255),
  [Activo] TEXT(2)
);
-- Rows: 65

CREATE TABLE [TbCacheArbolRiesgosMeta] (
  [IDEdicion] LONG(4),
  [ActiveBuildId] LONG(4),
  [UpdatedAt] SHORT_DATE_TIME(8)
);
-- Rows: 31

CREATE TABLE [TbCacheArbolRiesgosNodo] (
  [IDEdicion] LONG(4),
  [BuildId] LONG(4),
  [NodeKey] TEXT(255),
  [ParentKey] TEXT(255),
  [NodeType] TEXT(30),
  [IDRiesgo] LONG(4),
  [IDMitigacion] LONG(4),
  [IDContingencia] LONG(4),
  [IDAccion] LONG(4),
  [EsVisibleSinRetirados] BOOLEAN(1),
  [TextConDescripcion] MEMO,
  [TextSinDescripcion] MEMO,
  [IconName] TEXT(255),
  [ForeColor] LONG(4),
  [SortIndex] LONG(4),
  [Depth] LONG(4)
);
-- Rows: 2089

CREATE TABLE [TbCacheControlCambiosMeta] (
  [IDProyecto] LONG(4),
  [Edicion] LONG(4),
  [ActiveBuildId] LONG(4),
  [UpdatedAt] SHORT_DATE_TIME(8),
  [CacheVersion] TEXT(20)
);
-- Rows: 50

CREATE TABLE [TbCacheControlCambiosRow] (
  [IDProyecto] LONG(4),
  [Edicion] LONG(4),
  [BuildId] LONG(4),
  [CodigoRiesgo] TEXT(50),
  [EstadoHtml] MEMO,
  [MitigacionHtml] MEMO,
  [ContingenciaHtml] MEMO
);
-- Rows: 321

CREATE TABLE [TbCachePublicabilidadEdicion] (
  [IDEdicion] LONG(4),
  [Tipo] TEXT(10),
  [IDRiesgo] LONG(4),
  [Publicable] BOOLEAN(1),
  [Veredicto] LONG(4),
  [CodigoRiesgo] TEXT(50),
  [Descripcion] MEMO,
  [ChecksJson] MEMO,
  [AlgoritmoVersion] LONG(4),
  [UpdatedAt] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [tbCambios] (
  [IDCambio] LONG(4),
  [IDProyecto] LONG(4),
  [EdicionInicial] INT(2),
  [EdicionFinal] INT(2),
  [Riesgo] TEXT(255),
  [NombreCampo] TEXT(255),
  [Descripcion] MEMO,
  [FechaRegistro] SHORT_DATE_TIME(8)
);
-- Rows: 6725

CREATE TABLE [TbCambiosExplicacion] (
  [Apartado] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 5

CREATE TABLE [tbCambiosParaPublicacion] (
  [IDCambio] LONG(4),
  [IDProyecto] LONG(4),
  [EdicionInicial] INT(2),
  [EdicionFinal] INT(2),
  [Riesgo] TEXT(255),
  [NombreCampo] TEXT(255),
  [Descripcion] MEMO,
  [FechaRegistro] SHORT_DATE_TIME(8)
);
-- Rows: 68

CREATE TABLE [TbConfiguracionVisionRiesgos] (
  [IDConfiguracion] LONG(4),
  [Usuario] TEXT(255),
  [VerSoloNoRetirados] TEXT(2),
  [VerDescripcion] TEXT(2)
);
-- Rows: 35

-- LINKED/SKIPPED: [TbCorreosEnviados] (given file does not exist: C:\00repos\datos\Correos_datos.accdb)
-- LINKED/SKIPPED: [TbDocumentos] (given file does not exist: C:\00repos\datos\AGEDO20_Datos.accdb)
-- LINKED/SKIPPED: [TbDocumentosID] (given file does not exist: C:\00repos\datos\AGEDO20_Datos.accdb)
-- LINKED/SKIPPED: [TbDPD] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientes1] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesResponsables] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesSuministradores] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExplicacionCarencias] (
  [ID] LONG(4),
  [Apartado] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 6

CREATE TABLE [TbFelix] (
  [IDriesgo] LONG(4),
  [Prioridad] LONG(4)
);
-- Rows: 47

CREATE TABLE [TbHerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 26

CREATE TABLE [TbIDAccionPlanContingencia] (
  [IDAccionContingencia] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 251

CREATE TABLE [TbIDAccionPlanMitigacion] (
  [IDAccionMitigacion] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 1074

CREATE TABLE [TbIDAnexos] (
  [IDAnexo] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 107

CREATE TABLE [TbIDBibliotecaRiesgos] (
  [IDRiesgoTipo] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 51

CREATE TABLE [TbIDCorreosEnviados] (
  [IDCorreo] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 264

CREATE TABLE [TbIDDocumentos] (
  [IDDocumento] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 27

CREATE TABLE [TbIDEdiciones] (
  [IDEdicion] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 209

CREATE TABLE [TbIDLog] (
  [IDLog] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 1077

CREATE TABLE [TbIDPlanContingencia] (
  [IDContingencia] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 154

CREATE TABLE [TbIDPlanMitigacion] (
  [IDMitigacion] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 640

CREATE TABLE [TbIDProyectos] (
  [IDProyecto] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 71

CREATE TABLE [TbIDRiesgos] (
  [IDRiesgo] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 889

CREATE TABLE [TbIDRiesgosAIntegrar] (
  [IDRiesgoExt] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 58

CREATE TABLE [TbIDTareas] (
  [IDtarea] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 53

CREATE TABLE [TbLog] (
  [IDLog] LONG(4),
  [IDProyecto] LONG(4),
  [IDEdicion] LONG(4),
  [IDRiesgo] LONG(4),
  [IDMitigacion] LONG(4),
  [IDContingencia] LONG(4),
  [IDAccionMitigacion] LONG(4),
  [IDAccionContingencia] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Titulo] TEXT(255),
  [Linea] MEMO
);
-- Rows: 5817

CREATE TABLE [TbLogPublicaciones] (
  [ID] LONG(4),
  [IDEdicion] LONG(4),
  [FechaPreparadaParaPublicar] SHORT_DATE_TIME(8),
  [UsuarioFechaPreparadaParaPublicar] TEXT(255),
  [FechaPreparadaParaPublicarQuitar] SHORT_DATE_TIME(8),
  [UsuarioFechaPreparadaParaPublicarQuitar] TEXT(255),
  [FechaPublicacion] SHORT_DATE_TIME(8),
  [UsuarioFechaPublicacion] TEXT(255),
  [PropuestaRechazadaPorCalidadFecha] SHORT_DATE_TIME(8),
  [PropuestaRechazadaPorCalidadMotivo] MEMO,
  [FechaRegistro] SHORT_DATE_TIME(8)
);
-- Rows: 329

-- LINKED/SKIPPED: [TbNoConformidades] (given file does not exist: C:\00repos\datos\NoConformidades_Datos.accdb)
CREATE TABLE [TbOrigenesRiesgosDetalles] (
  [Origen] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 5

CREATE TABLE [TbProyectoEdicionesCorreoRevision] (
  [IDEnvioCorreoTecnico] LONG(4),
  [IDEdicion] LONG(4),
  [FechaCorreoRevision] SHORT_DATE_TIME(8),
  [UsuarioCalidad] TEXT(255),
  [IDCorreo] LONG(4)
);
-- Rows: 3

CREATE TABLE [TbProyectos] (
  [IDProyecto] LONG(4),
  [IDExpediente] LONG(4),
  [Proyecto] TEXT(255),
  [Juridica] TEXT(255),
  [NombreProyecto] TEXT(255),
  [Cliente] TEXT(255),
  [FechaPrevistaCierre] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [FechaRegistroInicial] SHORT_DATE_TIME(8),
  [Elaborado] TEXT(255),
  [Revisado] TEXT(255),
  [Aprobado] TEXT(255),
  [CodigoDocumento] TEXT(255),
  [ParaInformeAvisos] TEXT(2),
  [FechaFirmaContrato] TEXT(255),
  [RiesgosDeLaOferta] TEXT(2),
  [RiesgosDelSubContratista] TEXT(2),
  [NombreUsuarioCalidad] TEXT(255),
  [EnUTE] TEXT(2),
  [FechaMaxProximaPublicacion] SHORT_DATE_TIME(8),
  [RequiereRiesgoDeBiblioteca] TEXT(2),
  [CorreoRAC] TEXT(255),
  [Ordinal] LONG(4),
  [CadenaNombreAutorizados] MEMO,
  [NombreParaNodo] TEXT(255)
);
-- Rows: 81

-- LINKED/SKIPPED: [TbProyectos1] (given file does not exist: C:\00repos\datos\Gestion_Riesgos_Datos.accdb)
CREATE TABLE [TbProyectosEdiciones] (
  [IDEdicion] LONG(4),
  [IDProyecto] LONG(4),
  [FechaEdicion] SHORT_DATE_TIME(8),
  [Edicion] INT(2),
  [Elaborado] TEXT(255),
  [Revisado] TEXT(255),
  [Aprobado] TEXT(255),
  [FechaPublicacion] SHORT_DATE_TIME(8),
  [EntregadoAClienteORAC] TEXT(2),
  [Comentarios] MEMO,
  [PermitidoImprimirExcel] TEXT(2),
  [IDDocumentoAGEDO] LONG(4),
  [NombreArchivoInforme] TEXT(255),
  [FechaMaxProximaPublicacion] SHORT_DATE_TIME(8),
  [FechaPreparadaParaPublicar] SHORT_DATE_TIME(8),
  [UsuarioProponePublicar] TEXT(255),
  [PropuestaRechazadaPorCalidadFecha] SHORT_DATE_TIME(8),
  [PropuestaRechazadaPorCalidadMotivo] MEMO,
  [UsuarioCalidadRechazaPropuesta] TEXT(255),
  [NotasCalidadParaPublicar] MEMO,
  [FechaUltimoCambio] SHORT_DATE_TIME(8),
  [UsuarioUltimoCambio] TEXT(255)
);
-- Rows: 325

CREATE TABLE [TbProyectosEdicionesSuministradores] (
  [ID] LONG(4),
  [IDEdicion] LONG(4),
  [IDSuministrador] LONG(4),
  [IDAnexo] LONG(4)
);
-- Rows: 20

CREATE TABLE [TbProyectosResponsablesCalidad] (
  [IDProyecto] LONG(4),
  [UsuarioCalidad] TEXT(255)
);
-- Rows: 39

CREATE TABLE [TbProyectosSuministradores] (
  [ID] LONG(4),
  [IDProyecto] LONG(4),
  [IDSuministrador] LONG(4),
  [GestionCalidad] TEXT(2)
);
-- Rows: 0

CREATE TABLE [TbProyectosTipo] (
  [IDTipoProyecto] LONG(4),
  [TipoProyecto] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 0

-- LINKED/SKIPPED: [TbResponsablesExpedientes] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbRiesgos] (
  [IDRiesgo] LONG(4),
  [IDEdicion] LONG(4),
  [CodigoUnico] TEXT(255),
  [CodigoRiesgo] TEXT(5),
  [FechaDetectado] SHORT_DATE_TIME(8),
  [DetectadoPor] TEXT(255),
  [EntidadDetecta] TEXT(255),
  [Plazo] TEXT(255),
  [Calidad] TEXT(255),
  [Coste] TEXT(255),
  [ImpactoGlobal] TEXT(15),
  [Vulnerabilidad] TEXT(15),
  [Valoracion] TEXT(15),
  [Mitigacion] TEXT(15),
  [Contingencia] TEXT(15),
  [RequierePlanContingencia] TEXT(2),
  [Descripcion] MEMO,
  [CausaRaiz] MEMO,
  [Estado] TEXT(255),
  [FechaEstado] TEXT(255),
  [FechaMaterializado] SHORT_DATE_TIME(8),
  [FechaRetirado] SHORT_DATE_TIME(8),
  [FechaCerrado] SHORT_DATE_TIME(8),
  [FechaMitigacionAceptar] SHORT_DATE_TIME(8),
  [JustificacionAceptacionRiesgo] MEMO,
  [FechaJustificacionAceptacionRiesgo] SHORT_DATE_TIME(8),
  [FechaAprobacionAceptacionPorCalidad] SHORT_DATE_TIME(8),
  [FechaRechazoAceptacionPorCalidad] SHORT_DATE_TIME(8),
  [JustificacionRetiroRiesgo] MEMO,
  [FechaJustificacionRetiroRiesgo] SHORT_DATE_TIME(8),
  [FechaAprobacionRetiroPorCalidad] SHORT_DATE_TIME(8),
  [FechaRechazoRetiroPorCalidad] SHORT_DATE_TIME(8),
  [RequiereRiesgoDeBiblioteca] TEXT(2),
  [CodRiesgoBiblioteca] TEXT(255),
  [RiesgoPendienteRetipificacion] TEXT(2),
  [FechaRiesgoParaRetipificar] SHORT_DATE_TIME(8),
  [FechaRiesgoRetipificado] SHORT_DATE_TIME(8),
  [DiasSinRespuestaCalidadAceptacion] TEXT(255),
  [DiasSinRespuestaCalidadRetiro] TEXT(255),
  [DiasSinRespuestaCalidadRetipificacion] TEXT(255),
  [Origen] TEXT(255),
  [HayErrorEnRiesgo] TEXT(255),
  [NombreNodoDesc] TEXT(255),
  [NombreNodoEstado] TEXT(255),
  [NombreIcono] TEXT(255),
  [ColorIcono] TEXT(255),
  [Priorizacion] LONG(4)
);
-- Rows: 2924

CREATE TABLE [TbRiesgosAIntegrar] (
  [IDRiesgoExt] LONG(4),
  [CodRiesgo] TEXT(255),
  [IDRiesgo] LONG(4),
  [Origen] TEXT(255),
  [IDEdicion] LONG(4),
  [Descripcion] MEMO,
  [CausaRaiz] TEXT(255),
  [FechaDetectado] SHORT_DATE_TIME(8),
  [FechaAltaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [MotivoNoIntegrado] MEMO,
  [FechaMotivo] SHORT_DATE_TIME(8),
  [Trasladar] TEXT(2),
  [Suministrador] TEXT(255),
  [Pedido] TEXT(255),
  [ProveedorPedido] TEXT(255),
  [CausaRiesgoPedido] MEMO,
  [RequiereRiesgoDeBiblioteca] TEXT(2),
  [CodRiesgoBiblioteca] TEXT(255),
  [RiesgoPendienteRetipificacion] TEXT(2)
);
-- Rows: 175

CREATE TABLE [TbRiesgosAreasImpacto] (
  [Tipo] TEXT(255),
  [Indice] TEXT(255),
  [Descripcion] MEMO,
  [ordinal] BYTE(1)
);
-- Rows: 15

CREATE TABLE [TbRiesgosMaterializaciones] (
  [ID] LONG(4),
  [IDProyecto] LONG(4),
  [CodigoRiesgo] TEXT(5),
  [Fecha] SHORT_DATE_TIME(8),
  [IDEdicion] LONG(4),
  [EsMaterializacion] TEXT(2),
  [IDNC] LONG(4),
  [ParaNC] TEXT(2),
  [FechaDecison] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [IDPlanContingencia] LONG(4)
);
-- Rows: 32

CREATE TABLE [TbRiesgosNC] (
  [ID] LONG(4),
  [IDRiesgo] LONG(4),
  [IDNC] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [ParaNC] TEXT(2),
  [FechaDecison] SHORT_DATE_TIME(8)
);
-- Rows: 32

CREATE TABLE [TbRiesgosPlanContingenciaDetalle] (
  [IDAccionContingencia] LONG(4),
  [IDContingencia] LONG(4),
  [CodAccion] TEXT(255),
  [Accion] MEMO,
  [ResponsableAccion] TEXT(255),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [EsUltimaAccion] TEXT(2),
  [Estado] TEXT(255),
  [NombreIcono] TEXT(255)
);
-- Rows: 1817

CREATE TABLE [TbRiesgosPlanContingenciaDetalleReversa] (
  [IDAccionContingencia] LONG(4),
  [IDEdicion] LONG(4),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [Estado] TEXT(255)
);
-- Rows: 27

CREATE TABLE [TbRiesgosPlanContingenciaPpal] (
  [IDContingencia] LONG(4),
  [IDRiesgo] LONG(4),
  [CodContingencia] TEXT(255),
  [DisparadorDelPlan] MEMO,
  [FechaDeActivacion] SHORT_DATE_TIME(8),
  [FechaDesactivacion] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [NombreIcono] TEXT(255)
);
-- Rows: 1028

CREATE TABLE [TbRiesgosPlanMitigacionDetalle] (
  [IDAccionMitigacion] LONG(4),
  [IDMitigacion] LONG(4),
  [CodAccion] TEXT(255),
  [Accion] MEMO,
  [ResponsableAccion] TEXT(255),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [EsUltimaAccion] TEXT(2),
  [Estado] TEXT(255),
  [NombreIcono] TEXT(255)
);
-- Rows: 4221

CREATE TABLE [TbRiesgosPlanMitigacionDetalleReversa] (
  [IDAccionMitigacion] LONG(4),
  [IDEdicion] LONG(4),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [Estado] TEXT(255)
);
-- Rows: 41

CREATE TABLE [TbRiesgosPlanMitigacionPpal] (
  [IDMitigacion] LONG(4),
  [IDRiesgo] LONG(4),
  [CodMitigacion] TEXT(255),
  [DisparadorDelPlan] MEMO,
  [FechaDeActivacion] SHORT_DATE_TIME(8),
  [FechaDesactivacion] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [NombreIcono] TEXT(255)
);
-- Rows: 2843

CREATE TABLE [TbRiesgosPorTipoProyecto] (
  [IDTipoProyecto] LONG(4),
  [IDCatalogoRiesgo] LONG(4)
);
-- Rows: 0

CREATE TABLE [TbRiesgosValoracion] (
  [Impacto] TEXT(255),
  [Vulnerabilidad] TEXT(255),
  [Valoracion] TEXT(255)
);
-- Rows: 25

-- LINKED/SKIPPED: [TbSuministradores] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbSuministradoresSAP] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbTablas] (
  [Nombre] TEXT(255),
  [Orden] LONG(4),
  [Origen] TEXT(255),
  [NombreTablaOrigenID] TEXT(255),
  [NombreID1] TEXT(255),
  [NombreID2] TEXT(255),
  [TablaID] TEXT(255),
  [EsTablaID] TEXT(2),
  [RequiereID] TEXT(2)
);
-- Rows: 41

CREATE TABLE [TbTareas] (
  [IDtarea] LONG(4),
  [IDProyecto] LONG(4),
  [IDRiesgo] LONG(4),
  [CodRiesgo] TEXT(255),
  [TipoTarea] TEXT(255),
  [EstadoTarea] TEXT(255),
  [FechaAccion] SHORT_DATE_TIME(8),
  [FechaJustificacion] SHORT_DATE_TIME(8),
  [Justificacion] MEMO,
  [FechaVisado] SHORT_DATE_TIME(8),
  [FechaRechazo] SHORT_DATE_TIME(8)
);
-- Rows: 36

CREATE TABLE [TbTareasExplicaciones] (
  [NodoTarea] TEXT(255),
  [TituloTarea] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 10

CREATE TABLE [TbUltimoProyecto] (
  [IDUltimoProyecto] LONG(4),
  [IDProyecto] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8)
);
-- Rows: 29

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbValoresPosiblesContingencia] (
  [Valores] TEXT(255)
);
-- Rows: 4

CREATE TABLE [TbValoresPosiblesEstadoRiesgo] (
  [Valores] TEXT(255)
);
-- Rows: 7

CREATE TABLE [TbValoresPosiblesMitigacion] (
  [Valores] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 4

CREATE TABLE [TbValoresPosiblesPlazoCalidadCosteVulnerabilidad] (
  [Valores] TEXT(255)
);
-- Rows: 5

CREATE TABLE [TbValoresPosiblesValoracion] (
  [Valores] TEXT(255)
);
-- Rows: 4

CREATE TABLE [~TMPCLP576781] (
  [ID] LONG(4),
  [IDEdicion] LONG(4),
  [Suministrador] TEXT(255),
  [IDAnexo] LONG(4)
);
-- Rows: 0

CREATE TABLE [~TMPCLP603661] (
  [ID] LONG(4),
  [IDProyecto] LONG(4),
  [Suministrador] TEXT(255),
  [GestionCalidad] TEXT(2)
);
-- Rows: 0

-- Summary: 68 local tables, 16 linked/skipped
