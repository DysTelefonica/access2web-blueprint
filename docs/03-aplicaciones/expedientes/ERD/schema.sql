-- Schema dump: Expedientes_datos.accdb
-- Tables: 73

CREATE TABLE [Copia de TbExpedientes] (
  [IDExpediente] LONG(4),
  [IDExpedientePadre] LONG(4),
  [Nemotecnico] TEXT(255),
  [Titulo] MEMO,
  [ImporteLicitacion] DOUBLE(8),
  [ImporteContratacion] DOUBLE(8),
  [CodProyecto] TEXT(255),
  [CodExp] TEXT(255),
  [CodExpLargo] TEXT(255),
  [CodS4H] TEXT(255),
  [FechaInicioContrato] SHORT_DATE_TIME(8),
  [FechaFinContrato] SHORT_DATE_TIME(8),
  [FechaFinGarantia] SHORT_DATE_TIME(8),
  [EsAM] TEXT(2),
  [EsLote] TEXT(2),
  [EsBasado] TEXT(255),
  [EsExpediente] TEXT(2),
  [Ordinal] TEXT(255),
  [IdGradoClasificacion] LONG(4),
  [IDOrganoContratacion] LONG(4),
  [IDOficinaPrograma] LONG(4),
  [IDEjercito] LONG(4),
  [AccesoSharepoint] MEMO,
  [Observaciones] MEMO,
  [FechaCreacion] SHORT_DATE_TIME(8),
  [IDUsuarioCreacion] TEXT(255),
  [FechaUltimoCambio] SHORT_DATE_TIME(8),
  [IDUsuarioUltimoCambio] TEXT(255),
  [IDEstado] LONG(4),
  [Ambito] TEXT(10),
  [NPedido] TEXT(255),
  [IDResponsableCalidad] LONG(4),
  [Adjudicado] TEXT(2),
  [EnPeriodoDeAdjudicacion] TEXT(2),
  [Tipo] TEXT(255),
  [TipoInforme] TEXT(255),
  [AGEDYSAplica] TEXT(2),
  [AGEDYSGenerico] TEXT(2),
  [HPSAplica] TEXT(2),
  [CadenaPecal] TEXT(255),
  [Pecal] TEXT(2),
  [POSTAGEDO] TEXT(2),
  [APLICAESTADO] TEXT(2),
  [FECHAINICIOLICITACION] SHORT_DATE_TIME(8),
  [FECHAOFERTA] SHORT_DATE_TIME(8),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [FECHAFIRMACONTRATO] SHORT_DATE_TIME(8),
  [GARANTIAMESES] TEXT(255),
  [FECHACERTIFICACION] SHORT_DATE_TIME(8),
  [FECHAPERDIDA] SHORT_DATE_TIME(8),
  [FECHADESESTIMADA] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [CodigoActividad] TEXT(255),
  [HPSAplicaTareaS4H] TEXT(2)
);
-- Rows: 0

CREATE TABLE [Copia de TbExpedientesConEntidades] (
  [IDExpediente] LONG(4),
  [Clasificacion] TEXT(255),
  [OrganoContratacion] TEXT(255),
  [OficinaPrograma] TEXT(255),
  [Ejercito] TEXT(255),
  [Estado] TEXT(255),
  [ResponsableCalidad] TEXT(255),
  [ResponsableSeguridad] TEXT(255),
  [CadenaPecal] TEXT(255),
  [Pecal] TEXT(2),
  [CadenaContratistas] TEXT(255),
  [CadenaSubContratistas] TEXT(255),
  [CadenaSuministradores] TEXT(255),
  [CadenaComerciales] TEXT(255),
  [CadenaJPs] TEXT(255),
  [CadenaRACs] TEXT(255),
  [CadenaCorreoRACs] TEXT(255),
  [CadenaHitos] TEXT(255),
  [TipoParaLista] TEXT(255),
  [CadenaLugares] TEXT(255),
  [CadenaJuridicas] TEXT(255)
);
-- Rows: 365

CREATE TABLE [ListaPrevia] (
  [ID] DOUBLE(8),
  [IDExpedientePadre] DOUBLE(8),
  [Contratistas] TEXT(255),
  [Nemotécnico] TEXT(255),
  [Título] TEXT(255),
  [ObjetoContrato] TEXT(255),
  [Importe Licitación] TEXT(255),
  [Importe Contratación] DOUBLE(8),
  [Cod Proyecto] TEXT(255),
  [Nº Exp Corto] TEXT(255),
  [Nº Exp Largo] TEXT(255),
  [CodS4H] TEXT(255),
  [Cod Actividad] TEXT(255),
  [Fecha Inicio] SHORT_DATE_TIME(8),
  [Fecha Fin] SHORT_DATE_TIME(8),
  [Garantía (meses)] DOUBLE(8),
  [Fecha Fin Garantía] SHORT_DATE_TIME(8),
  [AM] TEXT(255),
  [Lote] TEXT(255),
  [Basado] TEXT(255),
  [Expediente] TEXT(255),
  [Número] DOUBLE(8),
  [Clasificación] TEXT(255),
  [OrganoContratación] TEXT(255),
  [Oficina de Programa] TEXT(255),
  [Ejército] TEXT(255),
  [AccesoSharepoint] TEXT(255),
  [Observaciones] MEMO,
  [Se Ejecuta en Defensa] TEXT(255),
  [NPedido] TEXT(255),
  [Calidad] TEXT(255),
  [Seguridad] TEXT(255),
  [Para AGEDYS] TEXT(255),
  [AGEDYS Genérico] TEXT(255),
  [Para HPS] TEXT(255),
  [Tipo] TEXT(255),
  [CadenaPecal] TEXT(255),
  [Pecal] TEXT(255),
  [POSTAGEDO] TEXT(255),
  [Aplica Estado] TEXT(255),
  [FPreOferta] TEXT(255),
  [FOferta] TEXT(255),
  [FAdjudicación] SHORT_DATE_TIME(8),
  [FFirma] SHORT_DATE_TIME(8),
  [FCertificación] SHORT_DATE_TIME(8),
  [FPerdida] TEXT(255),
  [FDesestimada] TEXT(255),
  [ESTADOCalculadoTexto] TEXT(255),
  [Jefe Proyecto] TEXT(255),
  [Comerciales] TEXT(255),
  [Racs] TEXT(255),
  [Hitos] TEXT(255)
);
-- Rows: 99

CREATE TABLE [TbAusExpPostAGEDO] (
  [ID] LONG(4)
);
-- Rows: 203

CREATE TABLE [TbAuxEstadosMartina] (
  [ID] LONG(4),
  [EstadoMartina] TEXT(255),
  [Igual] TEXT(2)
);
-- Rows: 143

CREATE TABLE [TbAuxNemotecnico] (
  [Nemotecnico] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCambios] (
  [IDCambio] LONG(4),
  [NombreTabla] TEXT(255),
  [NombreCampoID] TEXT(255),
  [ValorCampoID] LONG(4),
  [NombreCampo] TEXT(255),
  [ValorInicial] TEXT(255),
  [ValorFinal] TEXT(255),
  [FechaCambio] SHORT_DATE_TIME(8),
  [IDUsuarioCambio] LONG(4),
  [Accion] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbComerciales] (
  [IDComercial] LONG(4),
  [Comercial] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 26

-- LINKED/SKIPPED: [TbComunicados] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbConfMostrarEstado] (
  [ID] LONG(4),
  [UsuarioRed] TEXT(255),
  [MostrarEstado] TEXT(2)
);
-- Rows: 23

-- LINKED/SKIPPED: [TbCorreosEnviados] (given file does not exist: C:\00repos\datos\Correos_datos.accdb)
CREATE TABLE [TbCPV] (
  [IDCPV] LONG(4),
  [CPV] TEXT(255),
  [DESCRIPCION] MEMO
);
-- Rows: 73

CREATE TABLE [TbDatosEconomicosExpedientes] (
  [IDExpediente] LONG(4),
  [MontoTotal] DOUBLE(8),
  [CosteExternoPrevisto] DOUBLE(8),
  [CosteInternoPrevisto] DOUBLE(8),
  [GastoMaximoPrevisto] DOUBLE(8),
  [InversionPrevista] DOUBLE(8),
  [Observaciones] TEXT(255)
);
-- Rows: 91

-- LINKED/SKIPPED: [TbDpDInformeCondicionamiento] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbE2EExportBatch] (
  [IDBatch] LONG(4),
  [SessionId] TEXT(100),
  [UsuarioConectado] TEXT(255),
  [Estado] TEXT(50),
  [CreatedAt] SHORT_DATE_TIME(8),
  [StartedAt] SHORT_DATE_TIME(8),
  [CompletedAt] SHORT_DATE_TIME(8),
  [TotalSeleccionados] LONG(4),
  [TotalExportados] LONG(4),
  [ErrorMessage] MEMO
);
-- Rows: 0

CREATE TABLE [TbE2EExportBatchDetalle] (
  [IDBatchDetalle] LONG(4),
  [IDBatch] LONG(4),
  [IDExpediente] LONG(4),
  [OrdinalSeleccion] LONG(4),
  [HashExportado] TEXT(64),
  [Estado] TEXT(50),
  [CreatedAt] SHORT_DATE_TIME(8),
  [ExportedAt] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbE2EExportSeleccionTemp] (
  [IDTemp] LONG(4),
  [UsuarioConectado] TEXT(255),
  [SessionId] TEXT(100),
  [IDExpediente] LONG(4),
  [CreatedAt] SHORT_DATE_TIME(8)
);
-- Rows: 5

CREATE TABLE [TbE2EJsonDestinationUserConfig] (
  [IDConfig] LONG(4),
  [UsuarioRed] TEXT(255),
  [RutaDestino] MEMO,
  [CreatedAt] SHORT_DATE_TIME(8),
  [UpdatedAt] SHORT_DATE_TIME(8)
);
-- Rows: 4

CREATE TABLE [TbEjercitos] (
  [IDEjercito] LONG(4),
  [Ejercito] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 5

CREATE TABLE [TbEstados] (
  [IDEstado] LONG(4),
  [Estado] TEXT(255),
  [DESCRIPCION] MEMO
);
-- Rows: 9

CREATE TABLE [TbExpAgedys] (
  [IdExpediente] LONG(4),
  [Expediente] TEXT(255),
  [CodExpedienteLargo] TEXT(50),
  [OFICINADELPROGRAMA] TEXT(50),
  [CarpetaArchivo] TEXT(11),
  [FirmadelContrato] SHORT_DATE_TIME(8),
  [TITULOEXP] MEMO,
  [IMPORTEEXP] DOUBLE(8),
  [INTECDEF] TEXT(50),
  [RAC] TEXT(60),
  [FechaPrimerModificado] SHORT_DATE_TIME(8),
  [FechaSegundoModificado] SHORT_DATE_TIME(8),
  [FechaFinExp] SHORT_DATE_TIME(8),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Situacion] TEXT(150),
  [Adjudicado] BOOLEAN(1),
  [EnPeriodoDeAdjudicacion] BOOLEAN(1),
  [Comentarios] MEMO,
  [año] LONG(4),
  [Informe] BOOLEAN(1),
  [RistraResponsablesTemp] TEXT(255),
  [Pcal] TEXT(2),
  [PcalTipo] TEXT(50),
  [Generico] TEXT(2),
  [AñosGarantia] DOUBLE(8),
  [PCAP] TEXT(2),
  [IDDocumentoPCAP] LONG(4),
  [PPT] TEXT(2),
  [IDDocumentoPPT] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [FechaEntradaContrato] SHORT_DATE_TIME(8),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [CLASIFICACION] TEXT(255),
  [IDDocumentoContrato] LONG(4),
  [URLRAC] MEMO,
  [URLADJUDICACION] MEMO,
  [IDDocumentoRecepcion] LONG(4),
  [IDDocumentoProrrogas] LONG(4),
  [Margen] DOUBLE(8),
  [TieneVerificacionDisenio] TEXT(2),
  [FechaAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [FechaFinAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [AcuerdoMarco] TEXT(2),
  [IDDocumentoRAC] LONG(4),
  [ContratoBasadoDe] TEXT(255),
  [IDAdjudicacion] LONG(4),
  [ParaHacerPedidos] TEXT(2),
  [ResponsableCalidad] TEXT(255)
);
-- Rows: 116

CREATE TABLE [TbExpAGEDYS1] (
  [IdExpediente] LONG(4),
  [Expediente] TEXT(255),
  [CodExpedienteLargo] TEXT(50),
  [OFICINADELPROGRAMA] TEXT(50),
  [CarpetaArchivo] TEXT(11),
  [FirmadelContrato] SHORT_DATE_TIME(8),
  [TITULOEXP] MEMO,
  [IMPORTEEXP] DOUBLE(8),
  [INTECDEF] TEXT(50),
  [RAC] TEXT(60),
  [FechaPrimerModificado] SHORT_DATE_TIME(8),
  [FechaSegundoModificado] SHORT_DATE_TIME(8),
  [FechaFinExp] SHORT_DATE_TIME(8),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Situacion] TEXT(150),
  [Adjudicado] BOOLEAN(1),
  [EnPeriodoDeAdjudicacion] BOOLEAN(1),
  [Comentarios] MEMO,
  [año] LONG(4),
  [Informe] BOOLEAN(1),
  [RistraResponsablesTemp] TEXT(255),
  [Pcal] TEXT(2),
  [PcalTipo] TEXT(50),
  [Generico] TEXT(2),
  [AñosGarantia] DOUBLE(8),
  [PCAP] TEXT(2),
  [IDDocumentoPCAP] LONG(4),
  [PPT] TEXT(2),
  [IDDocumentoPPT] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [FechaEntradaContrato] SHORT_DATE_TIME(8),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [CLASIFICACION] TEXT(255),
  [IDDocumentoContrato] LONG(4),
  [URLRAC] MEMO,
  [URLADJUDICACION] MEMO,
  [IDDocumentoRecepcion] LONG(4),
  [IDDocumentoProrrogas] LONG(4),
  [Margen] DOUBLE(8),
  [TieneVerificacionDisenio] TEXT(2),
  [FechaAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [FechaFinAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [AcuerdoMarco] TEXT(2),
  [IDDocumentoRAC] LONG(4),
  [ContratoBasadoDe] TEXT(255),
  [IDAdjudicacion] LONG(4),
  [ParaHacerPedidos] TEXT(2),
  [ResponsableCalidad] TEXT(255)
);
-- Rows: 116

CREATE TABLE [TbExpedientes] (
  [IDExpediente] LONG(4),
  [IDExpedientePadre] LONG(4),
  [Nemotecnico] TEXT(255),
  [Titulo] MEMO,
  [ImporteLicitacion] DOUBLE(8),
  [ImporteContratacion] DOUBLE(8),
  [CodProyecto] TEXT(255),
  [CodExp] TEXT(255),
  [CodExpLargo] TEXT(255),
  [CodS4H] TEXT(255),
  [FechaInicioContrato] SHORT_DATE_TIME(8),
  [FechaFinContrato] SHORT_DATE_TIME(8),
  [FechaFinGarantia] SHORT_DATE_TIME(8),
  [EsAM] TEXT(2),
  [EsLote] TEXT(2),
  [EsBasado] TEXT(255),
  [EsExpediente] TEXT(2),
  [Ordinal] TEXT(255),
  [IdGradoClasificacion] LONG(4),
  [IDOrganoContratacion] LONG(4),
  [IDOficinaPrograma] LONG(4),
  [IDEjercito] LONG(4),
  [AccesoSharepoint] MEMO,
  [Observaciones] MEMO,
  [FechaCreacion] SHORT_DATE_TIME(8),
  [IDUsuarioCreacion] TEXT(255),
  [FechaUltimoCambio] SHORT_DATE_TIME(8),
  [IDUsuarioUltimoCambio] TEXT(255),
  [IDEstado] LONG(4),
  [Ambito] TEXT(10),
  [NPedido] TEXT(255),
  [IDResponsableCalidad] LONG(4),
  [Adjudicado] TEXT(2),
  [EnPeriodoDeAdjudicacion] TEXT(2),
  [Tipo] TEXT(255),
  [TipoInforme] TEXT(255),
  [AGEDYSAplica] TEXT(2),
  [AGEDYSGenerico] TEXT(2),
  [HPSAplica] TEXT(2),
  [CadenaPecal] TEXT(255),
  [Pecal] TEXT(2),
  [POSTAGEDO] TEXT(2),
  [FECHAPREOFERTA] SHORT_DATE_TIME(8),
  [APLICAESTADO] TEXT(2),
  [FECHAINICIOLICITACION] SHORT_DATE_TIME(8),
  [FECHAOFERTA] SHORT_DATE_TIME(8),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [FECHAFIRMACONTRATO] SHORT_DATE_TIME(8),
  [GARANTIAMESES] TEXT(255),
  [FECHACERTIFICACION] SHORT_DATE_TIME(8),
  [FECHAPERDIDA] SHORT_DATE_TIME(8),
  [FECHADESESTIMADA] SHORT_DATE_TIME(8),
  [ESTADO] TEXT(255),
  [CodigoActividad] TEXT(255),
  [AplicaTareaS4H] TEXT(2),
  [ContratistaPrincipal] TEXT(2),
  [IDResponsableSeguridad] LONG(4),
  [ObjetoContrato] MEMO,
  [OrdinalE2E] LONG(4),
  [HashActual] TEXT(64),
  [HashUltimaExportacion] TEXT(64)
);
-- Rows: 453

-- LINKED/SKIPPED: [TbExpedientes1] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExpedientes_antes] (
  [IdExpediente] LONG(4),
  [Expediente] TEXT(255),
  [CodExpedienteLargo] TEXT(50),
  [OFICINADELPROGRAMA] TEXT(50),
  [CarpetaArchivo] TEXT(11),
  [FirmadelContrato] SHORT_DATE_TIME(8),
  [TITULOEXP] MEMO,
  [IMPORTEEXP] DOUBLE(8),
  [INTECDEF] TEXT(50),
  [RAC] TEXT(60),
  [FechaPrimerModificado] SHORT_DATE_TIME(8),
  [FechaSegundoModificado] SHORT_DATE_TIME(8),
  [FechaFinExp] SHORT_DATE_TIME(8),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Situacion] TEXT(150),
  [Adjudicado] BOOLEAN(1),
  [EnPeriodoDeAdjudicacion] BOOLEAN(1),
  [Comentarios] MEMO,
  [año] LONG(4),
  [Informe] BOOLEAN(1),
  [RistraResponsablesTemp] TEXT(255),
  [Pcal] TEXT(2),
  [PcalTipo] TEXT(50),
  [Generico] TEXT(2),
  [AñosGarantia] DOUBLE(8),
  [PCAP] TEXT(2),
  [IDDocumentoPCAP] LONG(4),
  [PPT] TEXT(2),
  [IDDocumentoPPT] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [FechaEntradaContrato] SHORT_DATE_TIME(8),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [CLASIFICACION] TEXT(255),
  [IDDocumentoContrato] LONG(4),
  [URLRAC] MEMO,
  [URLADJUDICACION] MEMO,
  [IDDocumentoRecepcion] LONG(4),
  [IDDocumentoProrrogas] LONG(4),
  [Margen] DOUBLE(8),
  [TieneVerificacionDisenio] TEXT(2),
  [FechaAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [FechaFinAvisoVerificacionDisenio] SHORT_DATE_TIME(8),
  [Estado] TEXT(255),
  [AcuerdoMarco] TEXT(2),
  [IDDocumentoRAC] LONG(4),
  [ContratoBasadoDe] TEXT(255),
  [IDAdjudicacion] LONG(4),
  [ParaHacerPedidos] TEXT(2),
  [ResponsableCalidad] TEXT(255)
);
-- Rows: 457

CREATE TABLE [TbExpedientesAnexos] (
  [IDDocumento] LONG(4),
  [IDExpediente] LONG(4),
  [NombreDocumento] TEXT(255)
);
-- Rows: 710

CREATE TABLE [TbExpedientesAnualidades] (
  [IDAnualidad] LONG(4),
  [IDExpediente] LONG(4),
  [Año] INT(2),
  [BIIVA] DOUBLE(8),
  [BIIPSI] DOUBLE(8),
  [BIIGIC] DOUBLE(8),
  [BIEXENTA] DOUBLE(8),
  [IVA] DOUBLE(8),
  [IPSI] DOUBLE(8),
  [IGIC] DOUBLE(8),
  [PeriodoFacturacion] TEXT(255)
);
-- Rows: 174

-- LINKED/SKIPPED: [TbExpedientesAnualidades1] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbExpedientesCadenaContratacion] (
  [ID] LONG(4),
  [IDPadre] LONG(4),
  [IDExpediente] LONG(4),
  [IDSuministrador] LONG(4),
  [AplicaCalidad] TEXT(2),
  [AplicaRiesgos] TEXT(2),
  [AplicaContratosClasificados] TEXT(2),
  [AplicaHPS] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 0

CREATE TABLE [TbExpedientesCodigoCompras] (
  [ID] LONG(4),
  [IDExpediente] LONG(4),
  [CodCompras] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbExpedientesComerciales] (
  [IDComercialExpediente] LONG(4),
  [IDComercial] LONG(4),
  [IDExpediente] LONG(4)
);
-- Rows: 333

CREATE TABLE [TbExpedientesConEntidades] (
  [IDExpediente] LONG(4),
  [Clasificacion] TEXT(255),
  [OrganoContratacion] TEXT(255),
  [OficinaPrograma] TEXT(255),
  [Ejercito] TEXT(255),
  [Estado] TEXT(255),
  [ResponsableCalidad] TEXT(255),
  [ResponsableSeguridad] TEXT(255),
  [CadenaPecal] TEXT(255),
  [Pecal] TEXT(2),
  [CadenaContratistas] TEXT(255),
  [CadenaSubContratistas] TEXT(255),
  [CadenaSuministradores] TEXT(255),
  [CadenaComerciales] TEXT(255),
  [CadenaJPs] TEXT(255),
  [CadenaRACs] TEXT(255),
  [CadenaCorreoRACs] TEXT(255),
  [CadenaHitos] TEXT(255),
  [TipoParaLista] TEXT(255),
  [CadenaLugares] TEXT(255),
  [CadenaJuridicas] TEXT(255)
);
-- Rows: 451

CREATE TABLE [TbExpedientesCPVs] (
  [IDCPVExpediente] LONG(4),
  [IDCPV] LONG(4),
  [IDExpediente] LONG(4)
);
-- Rows: 429

CREATE TABLE [TbExpedientesE2E] (
  [IDExpediente] LONG(4),
  [HashPayload] TEXT(255),
  [Estado] TEXT(20),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaModificacion] SHORT_DATE_TIME(8),
  [UsuarioCreacion] TEXT(100),
  [UsuarioModificacion] TEXT(100)
);
-- Rows: 0

CREATE TABLE [TbExpedientesHitos] (
  [IDHitoExpediente] LONG(4),
  [IDExpediente] LONG(4),
  [Descripcion] TEXT(255),
  [FechaHito] SHORT_DATE_TIME(8),
  [FechaGarantiaHito] SHORT_DATE_TIME(8),
  [Importe] DOUBLE(8)
);
-- Rows: 46

CREATE TABLE [TbExpedientesJefaturas] (
  [IDJefaturaExpediente] LONG(4),
  [IDJefatura] LONG(4),
  [IDExpediente] LONG(4)
);
-- Rows: 0

CREATE TABLE [TbExpedientesJuridicas] (
  [IDExpedienteJuridica] LONG(4),
  [IDExpediente] LONG(4),
  [IDJuridica] LONG(4),
  [IDSuministrador] LONG(4),
  [ContratistaPrincipal] TEXT(2),
  [SubContratista] TEXT(2)
);
-- Rows: 417

CREATE TABLE [TbExpedientesLugaresEjecucion] (
  [IDExpedienteLugarEjecucion] LONG(4),
  [IDExpediente] LONG(4),
  [IDLugarEjecucion] LONG(4)
);
-- Rows: 194

CREATE TABLE [TbExpedientesModificados] (
  [IDExpedienteModificado] LONG(4),
  [IDExpediente] LONG(4),
  [NModificado] TEXT(255),
  [FechaFirmaModificado] SHORT_DATE_TIME(8),
  [FechaFinModificado] SHORT_DATE_TIME(8),
  [Descripcion] MEMO
);
-- Rows: 37

CREATE TABLE [TbExpedientesPECAL] (
  [IDPECALExpediente] LONG(4),
  [IDExpediente] LONG(4),
  [IDPECAL] LONG(4)
);
-- Rows: 366

CREATE TABLE [TbExpedientesRACS] (
  [IDRacExpediente] LONG(4),
  [IDExpediente] LONG(4),
  [IDRAC] LONG(4)
);
-- Rows: 193

CREATE TABLE [TbExpedientesResponsables] (
  [IDExpedienteResponsable] LONG(4),
  [IdExpediente] LONG(4),
  [IdUsuario] LONG(4),
  [CorreoSiempre] TEXT(2),
  [EsJefeProyecto] TEXT(2),
  [esPreventa] TEXT(2)
);
-- Rows: 730

CREATE TABLE [TbExpedientesSuministradores] (
  [IDExpedienteSuministrador] LONG(4),
  [IDExpediente] LONG(4),
  [IDSuministrador] LONG(4),
  [Descripcon] MEMO,
  [ContratistaPrincipal] TEXT(2),
  [SubContratista] TEXT(2),
  [IDPadre] LONG(4)
);
-- Rows: 713

-- LINKED/SKIPPED: [TbFacturasDetalle] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbFacturasPrincipal] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbGestionRiesgos] (given file does not exist: \\datoste\Aplicaciones_dys\Aplicaciones PpD\GESTION RIESGOS\Gestion_Riesgos_Datos.accdb)
CREATE TABLE [TbGradosClasificacion] (
  [IdGradoClasificacion] LONG(4),
  [GradoClasificacion] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 5

-- LINKED/SKIPPED: [TbGTVContratos] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbGTVContratosExpedientes] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbJefaturas] (
  [IDJefatura] LONG(4),
  [Jefatura] TEXT(255),
  [DESCRIPCION] MEMO
);
-- Rows: 3

CREATE TABLE [TbJuridicas] (
  [IDJuridica] LONG(4),
  [Juridica] TEXT(255),
  [DESCRIPCION] MEMO,
  [IDSuministrador] LONG(4)
);
-- Rows: 17

CREATE TABLE [TbLugaresEjecucion] (
  [IDLugarEjecucion] LONG(4),
  [LugarEjecucion] MEMO,
  [Descripcion] MEMO
);
-- Rows: 34

-- LINKED/SKIPPED: [TbNoConformidades] (given file does not exist: C:\00repos\datos\NoConformidades_Datos.accdb)
-- LINKED/SKIPPED: [TbNPedido] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbOficinasPrograma] (
  [IDOficinaPrograma] LONG(4),
  [OficinaPrograma] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 37

CREATE TABLE [TbOrganosContratacion] (
  [IDOrganoContratacion] LONG(4),
  [OrganoContratacion] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 46

CREATE TABLE [TbPECAL] (
  [IDPECAL] LONG(4),
  [PECAL] TEXT(255),
  [DESCRIPCION] MEMO
);
-- Rows: 6

-- LINKED/SKIPPED: [TbProyectos_AGEDYS] (given file does not exist: \\datoste\aplicaciones_dys\Aplicaciones PpD\Proyectos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbProyectos_Riesgos] (given file does not exist: \\datoste\Aplicaciones_dys\Aplicaciones PpD\GESTION RIESGOS\Gestion_Riesgos_Datos.accdb)
-- LINKED/SKIPPED: [TbProyectosEdiciones_Riesgos] (given file does not exist: \\datoste\Aplicaciones_dys\Aplicaciones PpD\GESTION RIESGOS\Gestion_Riesgos_Datos.accdb)
CREATE TABLE [TbRACS] (
  [IDRAC] LONG(4),
  [RAC] TEXT(255),
  [CORREO] TEXT(255),
  [DESCRIPCION] MEMO
);
-- Rows: 33

-- LINKED/SKIPPED: [TbResponsablesExpedientes] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbResponsablesPorRol] (
  [IDResponsablePorRol] LONG(4),
  [IDUsuario] LONG(4),
  [Rol] TEXT(50)
);
-- Rows: 11

-- LINKED/SKIPPED: [TbSolicitudesOfertasPrevias] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbSuministradores] (
  [IDSuministrador] LONG(4),
  [Nombre] TEXT(255),
  [CIF] TEXT(255),
  [DESCRIPCION] MEMO,
  [TramitadoraHPS] TEXT(2),
  [Nemotecnico] TEXT(255),
  [Direccion] MEMO,
  [CP] TEXT(255),
  [Ciudad] TEXT(255),
  [ConsorcioPropio] TEXT(2)
);
-- Rows: 72

-- LINKED/SKIPPED: [TbSuministradoresSAP] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbUltimoCambio] (
  [ID] LONG(4),
  [IDExpediente] LONG(4),
  [FechaCambio] SHORT_DATE_TIME(8),
  [IDUsuarioCambio] LONG(4)
);
-- Rows: 6

-- LINKED/SKIPPED: [TbUsuarios_HPS] (given file does not exist: \\datoste\Aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosHistoricos_HPS] (given file does not exist: \\datoste\Aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb)
-- LINKED/SKIPPED: [TbVisadoFacturas_Nueva] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbVisadosGenerales] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- Summary: 49 local tables, 24 linked/skipped
