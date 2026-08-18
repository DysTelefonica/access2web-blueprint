-- Schema dump: NoConformidades_Datos.accdb
-- Tables: 49

CREATE TABLE [Copia de TbNCARAvisos] (
  [ID] LONG(4),
  [IDAR] LONG(4),
  [IDCorreo15] LONG(4),
  [IDCorreo7] LONG(4),
  [IDCorreo0] LONG(4),
  [Fecha] SHORT_DATE_TIME(8)
);
-- Rows: 9401

CREATE TABLE [TbAnexos] (
  [IDAnexo] LONG(4),
  [IDNoConformidad] LONG(4),
  [TituloAnexo] TEXT(255),
  [DescripcionAnexo] MEMO,
  [URLInicial] TEXT(255),
  [NombreArchivoFinalAnexo] TEXT(255),
  [FechaAnexo] SHORT_DATE_TIME(8)
);
-- Rows: 228

CREATE TABLE [TbAnexosAuditoria] (
  [IDAnexo] LONG(4),
  [IDAuditoria] LONG(4),
  [URLInicial] TEXT(255),
  [NombreArchivo] TEXT(255),
  [FechaAnexo] SHORT_DATE_TIME(8)
);
-- Rows: 8

CREATE TABLE [TbAnexosNCAuditorias] (
  [IDAnexo] LONG(4),
  [IDNoConformidad] LONG(4),
  [URLInicial] TEXT(255),
  [NombreArchivo] TEXT(255),
  [FechaAnexo] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbAuditoriaLog] (
  [IDLog] LONG(4),
  [IDNC] LONG(4),
  [IDAC] LONG(4),
  [IDAR] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Titulo] MEMO,
  [Linea] MEMO
);
-- Rows: 0

CREATE TABLE [TbAuditorias] (
  [IDAuditoria] LONG(4),
  [Tipo] TEXT(255),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8)
);
-- Rows: 7

CREATE TABLE [TbAuxPuntoNorma] (
  [PuntoNorma] TEXT(255)
);
-- Rows: 28

CREATE TABLE [TbCacheIndicadoresAuditoriaDetalle] (
  [ID] LONG(4),
  [IDCacheIndicadorProyecto] LONG(4),
  [IDCacheConfig] LONG(4),
  [Dominio] TEXT(32),
  [Bucket] TEXT(50),
  [TipoFila] TEXT(50),
  [IDEntidad] LONG(4),
  [IDNoConformidad] LONG(4),
  [IDAccionCorrectiva] LONG(4),
  [IDAccionRealizada] LONG(4),
  [IDTarea] LONG(4),
  [ResponsableCalidad] TEXT(255),
  [ResponsableUsuarioRed] TEXT(255),
  [DisplayTitulo] TEXT(255),
  [DisplaySubtitulo] MEMO,
  [FechaSnapshot] SHORT_DATE_TIME(8),
  [ClaveEntidad] TEXT(255),
  [OrigenTabla] TEXT(255),
  [VersionRegla] TEXT(64)
);
-- Rows: 0

CREATE TABLE [TbCacheIndicadoresAuditoriaHeader] (
  [ID] LONG(4),
  [IDCacheConfig] LONG(4),
  [Dominio] TEXT(32),
  [FechaSincronizacion] SHORT_DATE_TIME(8),
  [UsuarioSincronizacion] TEXT(255),
  [Estado] TEXT(50),
  [ErrorUltimaSincronizacion] MEMO
);
-- Rows: 0

CREATE TABLE [TbCacheIndicadoresConfig] (
  [IDCacheConfig] LONG(4),
  [Dominio] TEXT(32),
  [Activo] BOOLEAN(1),
  [VersionRegla] TEXT(64),
  [FechaConfiguracion] SHORT_DATE_TIME(8),
  [UsuarioConfiguracion] TEXT(255)
);
-- Rows: 2

CREATE TABLE [TbCacheIndicadoresProyectoDetalle] (
  [IDCacheDetalle] LONG(4),
  [IDCacheIndicadorProyecto] LONG(4),
  [Bucket] TEXT(64),
  [TipoFila] TEXT(16),
  [IDEntidad] LONG(4),
  [IDNoConformidad] LONG(4),
  [IDAccionCorrectiva] LONG(4),
  [IDAccionRealizada] LONG(4),
  [ResponsableCalidad] TEXT(255),
  [CodigoNoConformidad] TEXT(255),
  [Descripcion] MEMO,
  [Nemotecnico] TEXT(255),
  [Tarea] MEMO,
  [Estado] TEXT(255),
  [Tecnico] TEXT(255),
  [TipoNC] TEXT(255),
  [IDExpediente] LONG(4),
  [NAccion] TEXT(50),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [RequiereControlEficacia] TEXT(50),
  [ResultadoControlEficacia] TEXT(255),
  [FechaSnapshot] SHORT_DATE_TIME(8),
  [IDCacheConfig] LONG(4),
  [Dominio] TEXT(32),
  [ClaveEntidad] TEXT(128),
  [IDTarea] LONG(4),
  [OrigenTabla] TEXT(64),
  [ResponsableUsuarioRed] TEXT(255),
  [DisplayTitulo] TEXT(255),
  [DisplaySubtitulo] MEMO,
  [FechaActualizacionEntidad] SHORT_DATE_TIME(8),
  [VersionRegla] TEXT(64)
);
-- Rows: 0

CREATE TABLE [TbCacheIndicadoresProyectoHeader] (
  [IDCacheIndicadorProyecto] LONG(4),
  [FechaSincronizacion] SHORT_DATE_TIME(8),
  [UsuarioSincronizacion] TEXT(255),
  [Estado] TEXT(50),
  [ErrorUltimaSincronizacion] MEMO,
  [IDCacheConfig] LONG(4),
  [Dominio] TEXT(32),
  [VersionRegla] TEXT(64),
  [MotivoSincronizacion] TEXT(64),
  [IDNoConformidadUltimaSync] LONG(4),
  [FechaUltimaSincronizacionNC] SHORT_DATE_TIME(8),
  [OperadorSync] TEXT(64)
);
-- Rows: 0

CREATE TABLE [TbCacheListadoNC] (
  [IDNoConformidad] LONG(4),
  [Version] LONG(4),
  [CodigoNoConformidad] TEXT(255),
  [IDExpediente] LONG(4),
  [Nemotecnico] TEXT(255),
  [CodExp] TEXT(255),
  [JuridicaExp] TEXT(255),
  [IDTipo] LONG(4),
  [Descripcion] MEMO,
  [Notas] MEMO,
  [Estado] TEXT(100),
  [FechaApertura] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [RequiereControlEficacia] TEXT(10),
  [ControlEficacia] TEXT(255),
  [ResponsableTelefonica] TEXT(255),
  [RESPONSABLECALIDAD] TEXT(255),
  [ACR] TEXT(255),
  [Cerrada] TEXT(10),
  [FechaCache] SHORT_DATE_TIME(8),
  [CacheValida] BOOLEAN(1)
);
-- Rows: 11

CREATE TABLE [TbCacheListadoNCAuditoria] (
  [ID] LONG(4),
  [IDAuditoria] LONG(4),
  [Tipo] TEXT(255),
  [Numero] TEXT(255),
  [Descripcion] MEMO,
  [CAUSARAIZ] MEMO,
  [RESPONSABLEIMPLANTACION] TEXT(255),
  [Estado] TEXT(255),
  [FechaApertura] SHORT_DATE_TIME(8),
  [FECHACIERRE] SHORT_DATE_TIME(8),
  [RequiereControlEficacia] TEXT(25),
  [ControlEficacia] MEMO,
  [Notas] MEMO,
  [Cerrada] TEXT(10),
  [Borrado] BOOLEAN(1),
  [AccionesCorrectivasConcatenadas] MEMO,
  [AccionesRealizadasConcatenadas] MEMO,
  [FechaCache] SHORT_DATE_TIME(8),
  [CacheValida] BOOLEAN(1),
  [Version] LONG(4)
);
-- Rows: 55

CREATE TABLE [TbCacheNCProyecto] (
  [IDNoConformidad] LONG(4),
  [IDCache] LONG(4),
  [Version] INT(2),
  [FechaCache] SHORT_DATE_TIME(8),
  [DatosNC] MEMO,
  [DatosACs] MEMO,
  [DatosARs] MEMO,
  [DatosReplanificaciones] MEMO,
  [DatosRiesgos] MEMO,
  [UsuarioCache] TEXT(50),
  [CacheValida] BOOLEAN(1)
);
-- Rows: 455

CREATE TABLE [TbConexiones] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 0

CREATE TABLE [TbConfiguracion] (
  [ID] LONG(4),
  [CacheHabilitada] BOOLEAN(1),
  [FechaCambioCache] SHORT_DATE_TIME(8),
  [UsuarioCambioCache] TEXT(255),
  [MotivoCambioCache] MEMO
);
-- Rows: 1

CREATE TABLE [TbConsultasPorFechas] (
  [Descripcion] TEXT(255)
);
-- Rows: 8

-- LINKED/SKIPPED: [TbCorreosEnviados] (given file does not exist: C:\00repos\datos\Correos_datos.accdb)
CREATE TABLE [TbDocumentosAuditorias] (
  [IDDocumento] LONG(4),
  [IDNoConformidad] LONG(4),
  [Documento] TEXT(255),
  [NombreAnexo] TEXT(255),
  [IDAccionRealizada] LONG(4),
  [IDAuditoria] LONG(4),
  [IDAuditoriaResultante] LONG(4)
);
-- Rows: 203

CREATE TABLE [TbEstadoCatalogo] (
  [TipoEntidad] TEXT(10),
  [EnumValor] LONG(4),
  [Codigo] TEXT(100),
  [Titulo] TEXT(255),
  [Texto] TEXT(100),
  [Version] LONG(4),
  [Activo] BOOLEAN(1),
  [FechaBootstrap] SHORT_DATE_TIME(8),
  [UsuarioBootstrap] TEXT(255)
);
-- Rows: 21

-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesResponsables] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbHerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 10

CREATE TABLE [TbLog] (
  [IDLog] LONG(4),
  [IDNC] LONG(4),
  [IDAC] LONG(4),
  [IDAR] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Titulo] MEMO,
  [Linea] MEMO
);
-- Rows: 4108

CREATE TABLE [TbLogAuditoria] (
  [IDLog] LONG(4),
  [IDNC] LONG(4),
  [IDAC] LONG(4),
  [IDAR] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Titulo] MEMO,
  [Linea] MEMO
);
-- Rows: 309

CREATE TABLE [TbLogCache] (
  [IDLog] LONG(4),
  [IDNoConformidad] LONG(4),
  [TipoOperacion] TEXT(50),
  [SeccionCache] TEXT(50),
  [Detalles] MEMO,
  [FechaOperacion] SHORT_DATE_TIME(8),
  [Usuario] TEXT(50),
  [DuracionMs] LONG(4),
  [Exito] BOOLEAN(1)
);
-- Rows: 135796

CREATE TABLE [TbNCAccionCorrectivas] (
  [IDAccionCorrectiva] LONG(4),
  [IDNoConformidad] LONG(4),
  [NAccion] LONG(4),
  [AccionCorrectiva] MEMO,
  [FechaAccionCorrectiva] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [FechaInicialMinima] SHORT_DATE_TIME(8),
  [FechaFinalUltima] SHORT_DATE_TIME(8),
  [Notas] MEMO,
  [Responsable] TEXT(255),
  [FechaFinPrevistaUltima] SHORT_DATE_TIME(8)
);
-- Rows: 604

CREATE TABLE [TbNCAccionesRealizadas] (
  [IDAccionRealizada] LONG(4),
  [IDAccionCorrectiva] LONG(4),
  [NAccion] LONG(4),
  [AccionRealizada] MEMO,
  [FechaAccionRealizada] SHORT_DATE_TIME(8),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [Notas] MEMO,
  [Responsable] TEXT(255)
);
-- Rows: 1588

CREATE TABLE [TbNCARAvisos] (
  [ID] LONG(4),
  [IDAR] LONG(4),
  [IDCorreo15] LONG(4),
  [IDCorreo7] LONG(4),
  [IDCorreo0] LONG(4),
  [Fecha] SHORT_DATE_TIME(8)
);
-- Rows: 332

CREATE TABLE [TbNCAuditoriaAccionCorrectivas] (
  [IDAccionCorrectiva] LONG(4),
  [ID] LONG(4),
  [NAccion] LONG(4),
  [AccionCorrectiva] MEMO,
  [FechaAccionCorrectiva] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [FechaInicialMinima] SHORT_DATE_TIME(8),
  [FechaFinalUltima] SHORT_DATE_TIME(8),
  [Notas] MEMO,
  [Responsable] TEXT(255),
  [FechaFinPrevistaUltima] SHORT_DATE_TIME(8)
);
-- Rows: 44

CREATE TABLE [TbNCAuditoriaAccionesRealizadas] (
  [IDAccionRealizada] LONG(4),
  [IDAccionCorrectiva] LONG(4),
  [NAccion] LONG(4),
  [AccionRealizada] MEMO,
  [FechaAccionRealizada] SHORT_DATE_TIME(8),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFinPrevista] SHORT_DATE_TIME(8),
  [FechaFinReal] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [Notas] MEMO,
  [Responsable] TEXT(255)
);
-- Rows: 61

CREATE TABLE [TbNCDocumentos] (
  [IDDocumento] LONG(4),
  [IDNoConformidad] LONG(4),
  [Documento] TEXT(255),
  [NombreAnexo] TEXT(255),
  [IDAccionRealizada] LONG(4),
  [IDNoConformidadResultante] LONG(4)
);
-- Rows: 918

CREATE TABLE [TbNCInformacionRAC] (
  [IDInformacionRAC] LONG(4),
  [IDNoConformidad] LONG(4),
  [Informacion] MEMO,
  [FechaInformacion] SHORT_DATE_TIME(8),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEdicion] SHORT_DATE_TIME(8),
  [UsuarioCrea] TEXT(255),
  [UltimoUsuarioEdita] TEXT(255)
);
-- Rows: 2

CREATE TABLE [TbNCInformacionRACAnexos] (
  [IDAnexoInformacionRAC] LONG(4),
  [IDInformacionRAC] LONG(4),
  [NombreArchivo] TEXT(255),
  [FechaAnexo] SHORT_DATE_TIME(8)
);
-- Rows: 1

CREATE TABLE [TbNoConformidades] (
  [IDNoConformidad] LONG(4),
  [Juridica] TEXT(255),
  [CodigoNoConformidad] TEXT(255),
  [EsNoConformidad] BOOLEAN(1),
  [EXPEDIENTE] TEXT(255),
  [PROYECTO] TEXT(255),
  [VEHICULO] TEXT(255),
  [DESCRIPCION] MEMO,
  [CAUSA] MEMO,
  [ENTIDADRESPONSABLE] TEXT(50),
  [RESPONSABLETELEFONICA] TEXT(50),
  [FECHAAPERTURA] SHORT_DATE_TIME(8),
  [FECHACIERRE] SHORT_DATE_TIME(8),
  [FPREVCIERRE] SHORT_DATE_TIME(8),
  [TIPO] TEXT(255),
  [NOTAS] MEMO,
  [Borrado] BOOLEAN(1),
  [RequiereACR] BOOLEAN(1),
  [ACR] MEMO,
  [MotivoBorrado] MEMO,
  [RequiereControlEficacia] TEXT(255),
  [ControlEficacia] MEMO,
  [FechaControlEficacia] SHORT_DATE_TIME(8),
  [FechaPrevistaControlEficacia] SHORT_DATE_TIME(8),
  [ResultadoControlEficacia] MEMO,
  [ConformeControlEficacia] TEXT(2),
  [RESPONSABLECALIDAD] TEXT(255),
  [IDExpediente] LONG(4),
  [CodExp] TEXT(255),
  [Nemotecnico] TEXT(255),
  [JuridicaExp] TEXT(255),
  [RESPONSABLECALIDADExp] TEXT(255),
  [CausaYAnalisRaiz] MEMO,
  [Tipologia] TEXT(255),
  [IDProyecto] LONG(4),
  [CodigoRiesgo] TEXT(255),
  [DetectadoPor] TEXT(255),
  [ResponsableEjecucion] TEXT(255),
  [ESTADO] TEXT(255),
  [IDTipo] LONG(4),
  [Cerrada] TEXT(2),
  [IDNCAsociada] LONG(4),
  [CodigoNoConformidadAsociada] TEXT(255),
  [CodConcesionAsociada] TEXT(255),
  [MotivoNoRequiereControlEficacia] MEMO
);
-- Rows: 438

CREATE TABLE [TbNoConformidadesAuditoria] (
  [ID] LONG(4),
  [IDAuditoria] LONG(4),
  [FechaApertura] SHORT_DATE_TIME(8),
  [Numero] TEXT(255),
  [DESCRIPCION] MEMO,
  [CAUSARAIZ] MEMO,
  [ACCIONCORRECTIVA] MEMO,
  [CORRECCION] MEMO,
  [FECHACIERRE] SHORT_DATE_TIME(8),
  [FPREVCIERRE] SHORT_DATE_TIME(8),
  [RESPONSABLEIMPLANTACION] TEXT(255),
  [RequiereControlEficacia] TEXT(25),
  [ControlEficacia] MEMO,
  [FechaControlEficacia] SHORT_DATE_TIME(8),
  [FechaPrevistaControlEficacia] SHORT_DATE_TIME(8),
  [ResultadoControlEficacia] MEMO,
  [ConformeControlEficacia] TEXT(2),
  [RequiereAccionCorrectiva] TEXT(2),
  [MotivoNoAccionCorrectiva] MEMO,
  [Tipo] TEXT(255),
  [PuntoNorma] TEXT(255),
  [ESTADO] TEXT(255),
  [Borrado] BOOLEAN(1),
  [MotivoBorrado] MEMO,
  [Notas] MEMO,
  [Cerrada] TEXT(2),
  [MotivoNoRequiereControlEficacia] MEMO
);
-- Rows: 55

CREATE TABLE [TbNoConformidadesIngresoPorLotesDetalle] (
  [IDLoteExcel] LONG(4),
  [CodigoNoConformidad] TEXT(255),
  [EsNoConformidad] BOOLEAN(1),
  [DocumentoDeReferencia] MEMO,
  [EXPEDIENTE] TEXT(255),
  [PROYECTO] TEXT(255),
  [VEHICULO] TEXT(255),
  [DESCRIPCION] MEMO,
  [CAUSA] MEMO,
  [ACCIONCORRECTIVA] MEMO,
  [ACCIONREALIZADA] MEMO,
  [ENTIDADRESPONSABLE] TEXT(50),
  [RESPONSABLETELEFONICA] TEXT(50),
  [FECHAAPERTURA] SHORT_DATE_TIME(8),
  [FECHAPREVISTACIERRE] SHORT_DATE_TIME(8),
  [FECHACIERRE] SHORT_DATE_TIME(8),
  [TIPO] TEXT(255),
  [NOTAS] MEMO,
  [DatosDelRegistro] MEMO
);
-- Rows: 166

CREATE TABLE [TbNoConformidadesIngresoPorLotesPrincipal] (
  [IDLoteExcel] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [URLCompletaOrigen] TEXT(255),
  [NombreArchivoExcel] TEXT(255),
  [FilaDatosInicial] INT(2),
  [Observaciones] MEMO
);
-- Rows: 1

CREATE TABLE [TbNoConformidadesIngresoPorLotesTemporal] (
  [IDTmp] LONG(4),
  [EsNoConformidadTmp] MEMO,
  [DocumentoDeReferenciaTmp] MEMO,
  [EXPEDIENTETmp] MEMO,
  [PROYECTOTmp] MEMO,
  [VEHICULOTmp] MEMO,
  [DESCRIPCIONTmp] MEMO,
  [CAUSATmp] MEMO,
  [ACCIONCORRECTIVATmp] MEMO,
  [ACCIONREALIZADATmp] MEMO,
  [ENTIDADRESPONSABLETmp] MEMO,
  [RESPONSABLETELEFONICATmp] MEMO,
  [FECHAAPERTURATmp] MEMO,
  [FECHAPREVISTACIERRETmp] MEMO,
  [FECHACIERRETmp] MEMO,
  [TIPOTmp] MEMO,
  [NOTASTmp] MEMO,
  [blnPasaCriterioParaGrabar] BOOLEAN(1),
  [DatosDelRegistroTmp] MEMO,
  [ValidacionDatos] MEMO
);
-- Rows: 166

CREATE TABLE [TbReplanificacionesAuditoria] (
  [IDReplanificacion] LONG(4),
  [IDNoConformidad] LONG(4),
  [IDAccionRealizada] LONG(4),
  [FechaReprogramacion] SHORT_DATE_TIME(8),
  [FechaPrevistaAlInicio] SHORT_DATE_TIME(8),
  [FechaPrevistaReplanificada] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 2

CREATE TABLE [TbReplanificacionesProyecto] (
  [IDReplanificacion] LONG(4),
  [IDNoConformidad] LONG(4),
  [IDAccionRealizada] LONG(4),
  [FechaReprogramacion] SHORT_DATE_TIME(8),
  [FechaPrevistaAlInicio] SHORT_DATE_TIME(8),
  [FechaPrevistaReplanificada] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 127

-- LINKED/SKIPPED: [TbRiesgos] (given file does not exist: C:\00repos\datos\Gestion_Riesgos_Datos.accdb)
-- LINKED/SKIPPED: [TbRiesgosNC] (given file does not exist: C:\00repos\datos\Gestion_Riesgos_Datos.accdb)
CREATE TABLE [TbTareasExplicaciones] (
  [NodoTarea] TEXT(255),
  [TituloTarea] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 10

CREATE TABLE [TbTipologia] (
  [CodTipologia] TEXT(2),
  [Tipologia] TEXT(255)
);
-- Rows: 23

CREATE TABLE [TbTiposNCProyectos] (
  [IDTipo] LONG(4),
  [Tipologia] TEXT(255)
);
-- Rows: 13

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- Summary: 42 local tables, 7 linked/skipped
