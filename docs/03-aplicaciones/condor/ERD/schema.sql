-- Schema dump: condor_datos.accdb
-- Tables: 25

CREATE TABLE [tbAdjuntos] (
  [idAdjunto] LONG(4),
  [idSolicitud] LONG(4),
  [etapaWF] TEXT(50),
  [nombreArchivo] TEXT(255),
  [fechaSubida] SHORT_DATE_TIME(8),
  [usuarioSubida] TEXT(100),
  [descripcion] MEMO,
  [TipoAccion] TEXT(255)
);
-- Rows: 0

-- LINKED/SKIPPED: [TbCorreosEnviados] (given file does not exist: C:\00repos\datos\Correos_datos.accdb)
CREATE TABLE [tbDatosCDCA] (
  [idDatosCDCA] LONG(4),
  [idSolicitud] LONG(4),
  [numContrato] TEXT(100),
  [refSuministrador] TEXT(100),
  [SuministradorNombreDir] MEMO,
  [refDesviacionesPrevias] TEXT(100),
  [requiereModificacionContrato] BOOLEAN(1),
  [identificacionMaterial] MEMO,
  [numPlanoEspecificacion] MEMO,
  [cantidadPeriodo] TEXT(50),
  [numSerieLote] TEXT(100),
  [causaNC] MEMO,
  [descripcionImpactoNC] MEMO,
  [descripcionImpactoNCCont] MEMO,
  [afectaPrestaciones] BOOLEAN(1),
  [afectaSeguridad] BOOLEAN(1),
  [afectaFiabilidad] BOOLEAN(1),
  [afectaVidaUtil] BOOLEAN(1),
  [afectaMedioambiente] BOOLEAN(1),
  [afectaIntercambiabilidad] BOOLEAN(1),
  [afectaMantenibilidad] BOOLEAN(1),
  [afectaApariencia] BOOLEAN(1),
  [afectaOtros] BOOLEAN(1),
  [impactoCoste] TEXT(50),
  [clasificacionNC] TEXT(50),
  [esSuministradorAD] BOOLEAN(1),
  [identificacionAutoridadDiseno] TEXT(100),
  [efectoFechaEntrega] MEMO,
  [firmaAprobacionRespIngenieriaNombre] TEXT(255),
  [firmaAprobacionRespProduccionNombre] TEXT(255),
  [firmaAprobacionRespCalidadNombre] TEXT(255),
  [firmaAprobacionRespDisenioNombre] TEXT(255),
  [firmaAprobacionRepresentanteSumNombre] TEXT(255),
  [racCodigo] TEXT(50),
  [observacionesRAC] MEMO,
  [racNombre] TEXT(255),
  [racDecision] TEXT(50),
  [decisionFinal] TEXT(50),
  [observacionesFinales] MEMO,
  [NombreFirmanteFinal] TEXT(100)
);
-- Rows: 0

CREATE TABLE [tbDatosCDCASUB] (
  [idDatosCDCASUB] LONG(4),
  [idSolicitud] LONG(4),
  [refSuministrador] TEXT(100),
  [refSubSuministrador] TEXT(100),
  [suministradorPrincipalNombreDir] MEMO,
  [subSuministradorNombreDir] MEMO,
  [refDesviacionesPrevias] TEXT(100),
  [requiereModificacionContrato] BOOLEAN(1),
  [identificacionMaterial] MEMO,
  [numPlanoEspecificacion] TEXT(100),
  [cantidadPeriodo] TEXT(50),
  [numSerieLote] TEXT(100),
  [causaNC] MEMO,
  [descripcionImpactoNC] MEMO,
  [descripcionImpactoNCCont] MEMO,
  [afectaPrestaciones] BOOLEAN(1),
  [afectaSeguridad] BOOLEAN(1),
  [afectaFiabilidad] BOOLEAN(1),
  [afectaVidaUtil] BOOLEAN(1),
  [afectaMedioambiente] BOOLEAN(1),
  [afectaIntercambiabilidad] BOOLEAN(1),
  [afectaMantenibilidad] BOOLEAN(1),
  [afectaApariencia] BOOLEAN(1),
  [afectaOtros] BOOLEAN(1),
  [impactoCoste] TEXT(50),
  [clasificacionNC] TEXT(50),
  [esSubSuministradorAD] BOOLEAN(1),
  [identificacionAutoridadDiseno] TEXT(100),
  [efectoFechaEntrega] MEMO,
  [firmaAprobacionRespIngenieriaNombre] TEXT(255),
  [firmaAprobacionRespProduccionNombre] TEXT(255),
  [firmaAprobacionRespCalidadNombre] TEXT(255),
  [firmaAprobacionRespDisenioNombre] TEXT(255),
  [firmaAprobacionRepresentanteSumNombre] TEXT(255),
  [racCodigo] TEXT(50),
  [observacionesRAC] MEMO,
  [racNombre] TEXT(255),
  [racDecision] TEXT(50),
  [racRechazoMotivos] MEMO,
  [observacionesRACDelegador] MEMO,
  [racNombreDelegador] TEXT(255),
  [decisionFinal] TEXT(50),
  [observacionesFinales] MEMO,
  [NombreFirmanteFinal] TEXT(100)
);
-- Rows: 0

CREATE TABLE [tbDatosPC] (
  [idDatosPC] LONG(4),
  [idSolicitud] LONG(4),
  [refContratoInspeccionOficial] TEXT(100),
  [refSuministrador] TEXT(100),
  [denominacionContrato] MEMO,
  [suministradorNombreDir] MEMO,
  [objetoContrato] MEMO,
  [descripcionMaterialAfectado] MEMO,
  [numPlanoEspecificacion] MEMO,
  [descripcionPropuestaCambio] MEMO,
  [descripcionPropuestaCambioCont] MEMO,
  [motivoCorregirDeficiencias] BOOLEAN(1),
  [motivoMejorarCapacidad] BOOLEAN(1),
  [motivoAumentarNacionalizacion] BOOLEAN(1),
  [motivoMejorarSeguridad] BOOLEAN(1),
  [motivoMejorarFiabilidad] BOOLEAN(1),
  [motivoMejorarCosteEficacia] BOOLEAN(1),
  [motivoOtros] BOOLEAN(1),
  [motivoOtrosDetalle] MEMO,
  [incidenciaCoste] TEXT(50),
  [incidenciaPlazo] TEXT(50),
  [incidenciaSeguridad] BOOLEAN(1),
  [incidenciaFiabilidad] BOOLEAN(1),
  [incidenciaMantenibilidad] BOOLEAN(1),
  [incidenciaIntercambiabilidad] BOOLEAN(1),
  [incidenciaVidaUtilAlmacen] BOOLEAN(1),
  [incidenciaFuncionamientoFuncion] BOOLEAN(1),
  [impactoClasificacion] TEXT(255),
  [CambioAfectaAMaterial] TEXT(255),
  [firmaOficinaTecnicaNombre] TEXT(100),
  [firmaRepSuministradorNombre] TEXT(100),
  [racCodigo] TEXT(50),
  [observacionesRAC] MEMO,
  [racNombre] TEXT(255),
  [racDecision] TEXT(50),
  [racRechazoMotivos] MEMO,
  [obsAprobacionAutoridadDiseno] MEMO,
  [NombreAutoridadDiseno] TEXT(100),
  [decisionFinal] TEXT(50),
  [obsDecisionFinal] MEMO,
  [NombreFirmanteFinal] TEXT(100)
);
-- Rows: 0

CREATE TABLE [tbDatosPCSUB] (
  [idDatosPCSUB] LONG(4),
  [idSolicitud] LONG(4),
  [refContratoInspeccionOficial] TEXT(100),
  [refSubSuministrador] TEXT(100),
  [denominacionContrato] MEMO,
  [SubsuministradorNombreDir] MEMO,
  [objetoContrato] MEMO,
  [descripcionMaterialAfectado] MEMO,
  [numPlanoEspecificacion] MEMO,
  [descripcionPropuestaCambio] MEMO,
  [descripcionPropuestaCambioCont] MEMO,
  [motivoCorregirDeficiencias] BOOLEAN(1),
  [motivoMejorarCapacidad] BOOLEAN(1),
  [motivoAumentarNacionalizacion] BOOLEAN(1),
  [motivoMejorarSeguridad] BOOLEAN(1),
  [motivoMejorarFiabilidad] BOOLEAN(1),
  [motivoMejorarCosteEficacia] BOOLEAN(1),
  [motivoOtros] BOOLEAN(1),
  [motivoOtrosDetalle] MEMO,
  [incidenciaCoste] TEXT(50),
  [incidenciaPlazo] TEXT(50),
  [incidenciaSeguridad] BOOLEAN(1),
  [incidenciaFiabilidad] BOOLEAN(1),
  [incidenciaMantenibilidad] BOOLEAN(1),
  [incidenciaIntercambiabilidad] BOOLEAN(1),
  [incidenciaVidaUtilAlmacen] BOOLEAN(1),
  [incidenciaFuncionamientoFuncion] BOOLEAN(1),
  [impactoClasificacion] TEXT(255),
  [CambioAfectaAMaterial] TEXT(255),
  [firmaOficinaTecnicaSubSuministradorNombre] TEXT(100),
  [firmaRepSubSuministradorNombre] TEXT(100),
  [racCodigo] TEXT(50),
  [observacionesRAC] MEMO,
  [racNombre] TEXT(255),
  [racDecision] TEXT(50),
  [racRechazoMotivos] MEMO,
  [observacionesRACDelegador] MEMO,
  [racNombreDelegador] TEXT(255),
  [obsAprobacionAutoridadDiseno] MEMO,
  [NombreAutoridadDiseno] TEXT(100),
  [decisionFinal] TEXT(50),
  [obsDecisionFinal] MEMO,
  [NombreFirmanteFinal] TEXT(100)
);
-- Rows: 0

CREATE TABLE [tbEstados] (
  [idEstado] LONG(4),
  [nombreEstado] TEXT(50),
  [descripcion] TEXT(255),
  [esEstadoInicial] BOOLEAN(1),
  [esEstadoFinal] BOOLEAN(1),
  [orden] LONG(4)
);
-- Rows: 9

-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesRACS] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesResponsables] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesSuministradores] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [tbHistorialRechazos] (
  [ID] LONG(4),
  [IdSolicitud] LONG(4),
  [FechaRechazo] SHORT_DATE_TIME(8),
  [UsuarioCalidad] TEXT(255),
  [AreaAfectada] TEXT(255),
  [MotivoPrincipal] MEMO,
  [Comentarios] MEMO,
  [EstaResuelto] BOOLEAN(1)
);
-- Rows: 0

CREATE TABLE [tbLogCambios] (
  [idLogCambio] LONG(4),
  [fechaHora] SHORT_DATE_TIME(8),
  [usuario] TEXT(100),
  [suplantadoPor] TEXT(255),
  [tabla] TEXT(50),
  [registro] LONG(4),
  [campo] TEXT(50),
  [valorAnterior] MEMO,
  [valorNuevo] MEMO,
  [tipoOperacion] TEXT(255)
);
-- Rows: 5

CREATE TABLE [tbLogErrores] (
  [idLogError] LONG(4),
  [fechaHora] SHORT_DATE_TIME(8),
  [usuario] TEXT(100),
  [suplantadoPor] TEXT(255),
  [modulo] TEXT(100),
  [procedimiento] TEXT(100),
  [numeroError] LONG(4),
  [descripcionError] MEMO,
  [contexto] MEMO
);
-- Rows: 3

CREATE TABLE [tbLogEstados] (
  [idLogEstado] LONG(4),
  [idSolicitud] LONG(4),
  [idEstadoAnterior] LONG(4),
  [idEstadoNuevo] LONG(4),
  [fechaTransicion] SHORT_DATE_TIME(8),
  [usuarioTransicion] TEXT(255)
);
-- Rows: 0

CREATE TABLE [tbMapeoCampos] (
  [idMapeo] LONG(4),
  [nombrePlantilla] TEXT(50),
  [nombreCampoTabla] TEXT(100),
  [valorAsociado] TEXT(100),
  [nombreCampoWord] TEXT(100),
  [numExtensiones] INT(2)
);
-- Rows: 183

-- LINKED/SKIPPED: [TbNoConformidades] (given file does not exist: C:\00repos\datos\NoConformidades_Datos.accdb)
-- LINKED/SKIPPED: [TbRACS] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [tbRechazos] (
  [idRechazo] LONG(4),
  [idSolicitud] LONG(4),
  [fechaRechazo] SHORT_DATE_TIME(8),
  [motivoPrincipal] TEXT(255),
  [areaAfectada] TEXT(255),
  [comentarios] MEMO,
  [usuarioRechazo] TEXT(255),
  [esActivo] BOOLEAN(1),
  [notasSubsanacion] MEMO,
  [CambiosTecnico] MEMO
);
-- Rows: 0

CREATE TABLE [tbSolicitudes] (
  [idSolicitud] LONG(4),
  [idExpediente] LONG(4),
  [tipoSolicitud] TEXT(20),
  [codigoSolicitud] TEXT(50),
  [idNCAsociada] LONG(4),
  [idEstadoInterno] LONG(4),
  [fechaCreacion] SHORT_DATE_TIME(8),
  [usuarioCreacion] TEXT(100),
  [fechaModificacion] SHORT_DATE_TIME(8),
  [usuarioModificacion] TEXT(100),
  [revisionCalidadEstado] TEXT(20),
  [revisionCalidadComentarios] MEMO
);
-- Rows: 1

-- LINKED/SKIPPED: [TbSuministradores] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [tbTransiciones] (
  [idTransicion] LONG(4),
  [idEstadoOrigen] LONG(4),
  [idEstadoDestino] LONG(4),
  [rolRequerido] TEXT(50)
);
-- Rows: 13

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [tbValidacionRevision] (
  [Id] LONG(4),
  [idSolicitud] LONG(4),
  [ordinal] INT(2),
  [idAdjunto] LONG(4),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Resultado] TEXT(255),
  [Comentarios] MEMO,
  [Usuario] TEXT(255),
  [HashDatos] TEXT(255)
);
-- Rows: 0

-- Summary: 15 local tables, 10 linked/skipped
