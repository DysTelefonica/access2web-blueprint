-- Schema dump: AGEDYS_DATOS.accdb
-- Tables: 139

-- LINKED/SKIPPED: [Salidas] (given file does not exist: C:\00repos\datos\Registro_Datos.accdb)
CREATE TABLE [TbAdjudicacionesEnvioAlSuministrador] (
  [IDEnvio] LONG(4),
  [NCarta] LONG(4),
  [IDSalida] LONG(4),
  [NSalida] TEXT(255),
  [NombreTabla] TEXT(255),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [IDCorreo] LONG(4),
  [NDPD] TEXT(255),
  [emailSuministrador] MEMO
);
-- Rows: 1598

CREATE TABLE [TbAdjudicacionesOfertasCartas] (
  [NCarta] LONG(4),
  [NDPD] TEXT(255),
  [NPedido] TEXT(255),
  [NAcreedor] LONG(4),
  [TEXTOOFERTA] MEMO,
  [IMPORTEOFERTA] DOUBLE(8),
  [FECHACARTA] SHORT_DATE_TIME(8),
  [TEXTORECTIFICATORIO] MEMO,
  [EMAIL] TEXT(255),
  [FechaMarcadaParaNoEnvio] SHORT_DATE_TIME(8),
  [IDAnexo] LONG(4)
);
-- Rows: 2027

CREATE TABLE [TbAmbitoConsultas] (
  [IdAmbitoConsulta] LONG(4),
  [AmbitoConsulta] TEXT(255)
);
-- Rows: 2

-- LINKED/SKIPPED: [TbAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbAuxExp] (
  [IdExpediente] LONG(4),
  [Expediente] TEXT(255),
  [FECHAADJUDICACION] SHORT_DATE_TIME(8),
  [FirmadelContrato] SHORT_DATE_TIME(8),
  [TITULOEXP] MEMO,
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [FechaFinTotal] SHORT_DATE_TIME(8),
  [Adjudicado] BOOLEAN(1),
  [AñosGarantia] DOUBLE(8),
  [Obs] MEMO
);
-- Rows: 40

CREATE TABLE [TbAuxExpedientesUsuario] (
  [IDExpediente] LONG(4),
  [Usuario] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbAuxExpedientesVivos] (
  [IDExpediente] LONG(4)
);
-- Rows: 40

CREATE TABLE [TbAuxGTV] (
  [Usuario] TEXT(50),
  [NContrato] TEXT(50),
  [Descripcion] TEXT(255),
  [ImporteRestante] DOUBLE(8),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [NombreEmpresa] TEXT(255)
);
-- Rows: 40

CREATE TABLE [TbAuxIDFactura] (
  [IDFactura] LONG(4)
);
-- Rows: 4923

CREATE TABLE [TbAuxPilar] (
  [NPEDIDO] TEXT(50),
  [FECHASRM] SHORT_DATE_TIME(8),
  [SRM] TEXT(50)
);
-- Rows: 56

CREATE TABLE [TbAuxPilar1] (
  [NDOCAntiguo] TEXT(255),
  [NDOCNuevo] TEXT(255)
);
-- Rows: 18

CREATE TABLE [TbAuxPrado] (
  [NºFactura] TEXT(50),
  [DPD] TEXT(50),
  [Fecha] SHORT_DATE_TIME(8),
  [DOCUMENTO] TEXT(50),
  [SRM] TEXT(50),
  [FECHASRM] SHORT_DATE_TIME(8),
  [Empresa] TEXT(255),
  [Cantidad] DOUBLE(8),
  [Visado Técnico] TEXT(255),
  [Cesta_Contrato] TEXT(50),
  [Cesta_z01] TEXT(50)
);
-- Rows: 468

CREATE TABLE [TbAvalesAnexo] (
  [CodInternoAval] LONG(4),
  [URLAval] TEXT(255)
);
-- Rows: 32

CREATE TABLE [TbAvalesCancelacion] (
  [CodInternoAval] LONG(4),
  [FechaNotificacionClienteLiberalizacionAval] SHORT_DATE_TIME(8),
  [FechaOrdenCancelacionAval] SHORT_DATE_TIME(8),
  [FechaRecogidaAvalCGD] SHORT_DATE_TIME(8),
  [FechaDevolucionAvalTGestiona] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 187

CREATE TABLE [TbAvalesDatosGenerales] (
  [NombreContactoTGestiona] TEXT(255),
  [CorreoContactoTGestiona] TEXT(255),
  [TlfnoContactoTGestiona] LONG(4),
  [NombreJefeContactoTGestiona] TEXT(255),
  [TlfnoJefeContactoTGestiona] LONG(4),
  [FaxTGestiona] LONG(4),
  [DireccionSolicitante] TEXT(255),
  [GerenciaSolicitante] TEXT(255),
  [TlfnoResponsableSolicitud] LONG(4),
  [NombreResponsableSolicitud] TEXT(255),
  [FaxSolicitante] LONG(4),
  [EntidadAvalada] TEXT(255),
  [Cliente] TEXT(255),
  [AFavorDe] TEXT(255),
  [Organismo] TEXT(255),
  [Dependencia] TEXT(255),
  [CIFCliente] TEXT(50),
  [ConceptoAval] TEXT(255),
  [NombreContactoCliente] TEXT(255),
  [TlfnoContactoCliente] LONG(4),
  [DireccionContactoCliente] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbAvalesDatosGeneralesNuevos] (
  [FaxTGestiona] LONG(4),
  [NombreSolicitante] TEXT(255),
  [TlfnoSolicitante] LONG(4),
  [DireccionSolicitante] TEXT(255),
  [emailSolicitante] TEXT(255),
  [NombreContactoArea] TEXT(255),
  [TlfnoContactoArea] LONG(4),
  [emailContactoArea] TEXT(255),
  [NombreResponsableSolicitud] TEXT(255),
  [emailResponsable] TEXT(255),
  [EntidadAvalada] TEXT(255),
  [Cliente] TEXT(255),
  [AFavorDe] TEXT(255),
  [CIFCliente] TEXT(50),
  [Organismo] TEXT(255),
  [Dependencia] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbAvalesFacturasClienteVinculo] (
  [codAval] LONG(4),
  [NFactura] TEXT(50)
);
-- Rows: 6

CREATE TABLE [TbAvalesInicio] (
  [CodInternoAval] LONG(4),
  [OrdenAnual] TEXT(3),
  [año] TEXT(50),
  [NombreContactoTGestiona] TEXT(255),
  [CorreoContactoTGestiona] TEXT(255),
  [TlfnoContactoTGestiona] LONG(4),
  [NombreJefeContactoTGestiona] TEXT(255),
  [TlfnoJefeContactoTGestiona] LONG(4),
  [FaxTGestiona] LONG(4),
  [NombreSolicitante] TEXT(255),
  [emailSolicitante] TEXT(255),
  [TlfnoSolicitante] LONG(4),
  [DireccionSolicitante] TEXT(255),
  [GerenciaSolicitante] TEXT(255),
  [NombreContactoArea] TEXT(255),
  [TlfnoContactoArea] LONG(4),
  [emailContactoArea] TEXT(255),
  [TlfnoResponsableSolicitud] LONG(4),
  [NombreResponsableSolicitud] TEXT(255),
  [emailResponsable] TEXT(255),
  [FaxSolicitante] LONG(4),
  [EntidadAvalada] TEXT(255),
  [Cliente] TEXT(255),
  [AFavorDe] TEXT(255),
  [Organismo] TEXT(255),
  [Dependencia] TEXT(255),
  [CIFCliente] TEXT(50),
  [ConceptoAval] TEXT(255),
  [NombreContactoCliente] TEXT(255),
  [TlfnoContactoCliente] LONG(4),
  [DireccionContactoCliente] TEXT(255),
  [ClaseAval] TEXT(50),
  [IDExpediente] LONG(4),
  [TextoExpediente] TEXT(255),
  [FechaPrevistaLiberizacion] SHORT_DATE_TIME(8),
  [FechaVencimiento] SHORT_DATE_TIME(8),
  [FechaInicioGestionAval] SHORT_DATE_TIME(8),
  [ImporteAval] DOUBLE(8),
  [TecnicoQuePideElAval] TEXT(255),
  [TlfnoAreaContratante] LONG(4),
  [AreaContratante] TEXT(255),
  [CODEXPLARGO] TEXT(50)
);
-- Rows: 185

CREATE TABLE [TbAvalesInicioNuevo] (
  [CodInternoAval] LONG(4),
  [OrdenAnual] TEXT(3),
  [año] TEXT(50),
  [SOCIEDAD_AVALADA] TEXT(50),
  [OTRAS_SOCIEDADES_AVALADAS] MEMO,
  [SOLICITUD_DIRECCION_Y_AREA] MEMO,
  [SOLICITUD_SOLICITANTE_NOMBRE] TEXT(255),
  [SOLICITUD_SOLICITANTE_TELEFONO] TEXT(50),
  [SOLICITUD_SOLICITANTE_EMAIL] TEXT(255),
  [SOLICITUD_PERSONA_CONTACTO_AREA_NOMBRE] TEXT(255),
  [SOLICITUD_PERSONA_CONTACTO_AREA_TELEFONO] TEXT(50),
  [SOLICITUD_PERSONA_CONTACTO_AREA_EMAIL] TEXT(255),
  [SOLICITUD_PERSONA_RESPONABLE_AREA_NOMBRE] TEXT(255),
  [SOLICITUD_PERSONA_RESPONSABLE_AREA_TELEFONO] TEXT(50),
  [SOLICITUD_PERSONA_RESPONSABLE_AREA_EMAIL] TEXT(255),
  [DATOS_AVAL_CLIENTE_BENEFICIARIO] MEMO,
  [DATOS_AVAL_CIF_NIF] TEXT(50),
  [DATOS_AVAL_AREA_CONTRATANTE_CONTACTO] MEMO,
  [DATOS_AVAL_FIRMAS_LEGITIMADAS] TEXT(50),
  [DATOS_AVAL_INTERVENIDO] TEXT(50),
  [DATOS_AVAL_SEGUN_MODELO] TEXT(255),
  [DATOS_AVAL_IMPORTE] DOUBLE(8),
  [DATOS_AVAL_DIVISA] TEXT(50),
  [DATOS_AVAL_CLASE] TEXT(50),
  [DATOS_AVAL_FECHA_PREVISTA_LIBERALIZACION] SHORT_DATE_TIME(8),
  [DATOS_AVAL_TEXTO_CONCEPTO] MEMO,
  [DATOS_AVAL_IC_LUGAR_RECOGIDA] TEXT(50),
  [DATOS_AVAL_IC_FECHA] SHORT_DATE_TIME(8),
  [DATOS_AVAL_IC_SIADE] TEXT(50),
  [DATOS_AVAL_IC_SOLICITANTE] TEXT(255),
  [IDEXPEDIENTE] LONG(4),
  [TRAMITADO] TEXT(50),
  [MOTIVO_NO_TRAMITACION] MEMO,
  [FechaInicioGestionAval] SHORT_DATE_TIME(8)
);
-- Rows: 56

CREATE TABLE [TbAvalesObtencionAvalYDeposito] (
  [CodInternoAval] LONG(4),
  [FechaRecogidaPrevista] SHORT_DATE_TIME(8),
  [DireccionEntidadBancaria] TEXT(255),
  [Entidad Bancaria] TEXT(255),
  [FechaRecogidaReal] SHORT_DATE_TIME(8),
  [NAval] TEXT(50),
  [FechaDeposito] SHORT_DATE_TIME(8),
  [NREGISTROCGD] TEXT(50),
  [FechaEntregaAlCliente] SHORT_DATE_TIME(8)
);
-- Rows: 132

CREATE TABLE [TbAvalesSolicitudSIADE] (
  [CodInternoAval] LONG(4),
  [FechaSolicitudSIADE] SHORT_DATE_TIME(8),
  [SIADE] TEXT(50),
  [FechaAprovacion] SHORT_DATE_TIME(8)
);
-- Rows: 151

CREATE TABLE [TbBloqueos] (
  [CODPROYECTOS] TEXT(50),
  [BLOQUEADO] BOOLEAN(1),
  [BLOQUEANTE] TEXT(50),
  [OBSERVACIONES] MEMO
);
-- Rows: 2254

CREATE TABLE [TbCalendarioProvisiones] (
  [IdProvision] LONG(4),
  [FechaInicialProvision] SHORT_DATE_TIME(8),
  [FechaFinalProvision] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 125

CREATE TABLE [TbCalendarioProvisionesDetalle] (
  [IDProvisionDetalle] LONG(4),
  [IDProvision] LONG(4),
  [NContrato] TEXT(50),
  [Posicion] TEXT(50),
  [Cesta] TEXT(50),
  [NPedido] TEXT(50),
  [ImporteSolicitado] DOUBLE(8),
  [ImporteAdjudicado] DOUBLE(8),
  [Pte_Facturar] DOUBLE(8),
  [Expediente] TEXT(50),
  [Responsables] MEMO,
  [Peticionario] TEXT(255),
  [NDPD] TEXT(50),
  [Empresa] TEXT(255),
  [FNPedido] SHORT_DATE_TIME(8),
  [FechaCesta] SHORT_DATE_TIME(8)
);
-- Rows: 247

CREATE TABLE [TbCestasAnexosOrdinaria] (
  [IDCestaAnexo] LONG(4),
  [Cesta] TEXT(255),
  [IDTipoAnexoOrdinaria] LONG(4),
  [URLLocal] TEXT(255),
  [URLFinal] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCestasTiposDocumentos] (
  [ID] LONG(4),
  [NombreDocumento] TEXT(255),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaBaja] SHORT_DATE_TIME(8)
);
-- Rows: 0

-- LINKED/SKIPPED: [TbComerciales] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbComunicados] (
  [IDComunicado] LONG(4),
  [Titulo] TEXT(255),
  [Emisor] TEXT(50),
  [TextoComunicado] MEMO,
  [URLDocumentoAmplicacion] TEXT(255),
  [FechaComunicado] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbComunicadosUsuario] (
  [IDComunicado] LONG(4),
  [strUsuario] TEXT(50),
  [Mostrar] TEXT(2)
);
-- Rows: 0

CREATE TABLE [TbConexiones] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 24

CREATE TABLE [TbConexionesAGEDYSECO] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 9

CREATE TABLE [TbConexionesAGEDYSTEC] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 15

-- LINKED/SKIPPED: [TbConexionesRegistro] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbConsultasDetalle] (
  [IDConsulta] LONG(4),
  [IDCategoria] LONG(4),
  [NombreConsulta] TEXT(255),
  [SQLConsulta] MEMO,
  [SQLSimple] MEMO,
  [DescripcionConsulta] MEMO,
  [Compleja] LONG(4),
  [TbAux] TEXT(255)
);
-- Rows: 23

CREATE TABLE [TbConsultasOpciones] (
  [IdElemento] LONG(4),
  [NombreNodo] TEXT(255),
  [ClavePadre] TEXT(50),
  [ClaveNodo] TEXT(50),
  [DescripcionNodo] MEMO,
  [SQLNodo] MEMO,
  [Observaciones] TEXT(255),
  [ImagenNodo] INT(2)
);
-- Rows: 6

CREATE TABLE [TbConsultasPrincipal] (
  [IDCategoria] LONG(4),
  [NombreCategoria] TEXT(255),
  [DescripcionCategoria] MEMO
);
-- Rows: 7

CREATE TABLE [TbCorreoDiferidoNotas] (
  [IDEvento] LONG(4),
  [IDCorreo] LONG(4),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaNota] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 117

-- LINKED/SKIPPED: [TbCorreosEnviados] (given file does not exist: C:\00repos\datos\Correos_datos.accdb)
-- LINKED/SKIPPED: [TbCPV] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbCriteriosSuministrador] (
  [IDCriterio] LONG(4),
  [Criterio] TEXT(255),
  [Condiciones] TEXT(255)
);
-- Rows: 6

CREATE TABLE [TbCSS] (
  [IDCSS] LONG(4),
  [Descripcion] TEXT(255),
  [HTML] MEMO
);
-- Rows: 1

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

CREATE TABLE [TbDireccionesSuministradores] (
  [Suministrador] TEXT(50),
  [LineaDireccion1] TEXT(255),
  [LineaDireccion2] TEXT(255),
  [LineaDireccion3] TEXT(255),
  [LineaDireccion4] TEXT(255),
  [LineaDireccion5] TEXT(255)
);
-- Rows: 1

-- LINKED/SKIPPED: [TbDocExpProceso] (given file does not exist: C:\00repos\datos\AGEDO20_Datos.accdb)
-- LINKED/SKIPPED: [TbDocumentos] (given file does not exist: C:\00repos\datos\AGEDO20_Datos.accdb)
-- LINKED/SKIPPED: [TbDocumentosID] (given file does not exist: C:\00repos\datos\AGEDO20_Datos.accdb)
CREATE TABLE [TbDpDAsuntoDefensa] (
  [CODPROYECTO] TEXT(255),
  [AsuntoDefensa] BOOLEAN(1)
);
-- Rows: 277

CREATE TABLE [TbDPDCartaSolicitudOfertaDatos] (
  [NDPD] TEXT(50),
  [NAcreedor] LONG(4),
  [ATTPARACARTA] TEXT(255),
  [DESCRIPCIONPARACARTA] TEXT(255),
  [ARCHIVOPARACARTA] TEXT(50),
  [PETICIONARIOPARACARTA] TEXT(255),
  [DESTINOPARACARTA] TEXT(255),
  [CLASEOBRAPARACARTA] TEXT(255),
  [MOTIVOOBRAPARACARTA] TEXT(255),
  [GENERALPARACARTA] MEMO,
  [URLEXCELEQUIPAMIENTOPARACARTA] TEXT(255),
  [TEXTOEQUIPAMIENTOPARACARTA] MEMO,
  [NOMBREPERSONACONTACTO1PARACARTA] TEXT(255),
  [NOMBREPERSONACONTACTO2PARACARTA] TEXT(255),
  [TELPERSONACONTACTO1PARACARTA] TEXT(50),
  [TELPERSONACONTACTO2PARACARTA] TEXT(50),
  [FINICIAL] SHORT_DATE_TIME(8),
  [FFINAL] SHORT_DATE_TIME(8),
  [FechaCarta] SHORT_DATE_TIME(8),
  [NombreDireccionDys] TEXT(255),
  [DIRDys] TEXT(255),
  [CPDYS] TEXT(50),
  [FORMAPAGODIAS] TEXT(50),
  [CODUNIDADRESPONSABLE] TEXT(50),
  [TEXTOTRAMITACIONFACTURAS] MEMO,
  [GERENCIA] TEXT(255),
  [TELGERENCIA] TEXT(255),
  [FAXGERENCIA] TEXT(255),
  [EMAILGERENCIA] TEXT(255),
  [PedidoPropiedadDelCliente] BOOLEAN(1),
  [TipoMaterial] TEXT(50),
  [TipoDPD] TEXT(255),
  [UsuarioQueRegistra] TEXT(50),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [FechaComunicacionSecretaria] SHORT_DATE_TIME(8),
  [TEXTOREQUISITOSOTAN] MEMO,
  [RB01] TEXT(255),
  [RB02] TEXT(255),
  [RS01] TEXT(255)
);
-- Rows: 2334

CREATE TABLE [TbDpDCartaSolicitudOfertaDatosFijos] (
  [NombreDireccionDys] TEXT(255),
  [DIRDys] TEXT(255),
  [CPDYS] TEXT(50),
  [FORMAPAGODIAS] TEXT(50),
  [CODUNIDADRESPONSABLE] TEXT(50),
  [TEXTOTRAMITACIONFACTURAS] MEMO,
  [GERENCIA] TEXT(255),
  [TELGERENCIA] TEXT(255),
  [FAXGERENCIA] TEXT(255),
  [EMAILGERENCIA] TEXT(255),
  [TEXTOREQUISITOSOTAN] MEMO
);
-- Rows: 1

CREATE TABLE [TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos] (
  [ParaDefensa] BOOLEAN(1),
  [TipoMaterial] TEXT(50),
  [TipoDPD] TEXT(255),
  [Pagina3Nombre] TEXT(50),
  [RequiereRB01] BOOLEAN(1),
  [RequiereRB02] BOOLEAN(1),
  [RequiereRS01] BOOLEAN(1),
  [ClaseObra] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 25

CREATE TABLE [TbDPDComiteOfertas] (
  [CODDPD] TEXT(255),
  [COMITEOFERTA] TEXT(255),
  [FECHACOMITEOFERTA] SHORT_DATE_TIME(8),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistro] TEXT(255)
);
-- Rows: 8

CREATE TABLE [TbDpDDocAnexada] (
  [IDDocumento] LONG(4),
  [NDPD] TEXT(50),
  [TITULODOC] TEXT(255),
  [TIPODOC] TEXT(50),
  [URLDOC] TEXT(255),
  [DESCRIPCIONDOC] MEMO,
  [QUIENLOGENERA] TEXT(50),
  [NRegistroCompletoSalida] TEXT(50),
  [NRegistroCompletoEntrada] TEXT(50),
  [ObservacionesDocAnexada] MEMO,
  [URLFINAL] MEMO,
  [FechaAnexion] SHORT_DATE_TIME(8),
  [NCartaAdjudicacion] TEXT(50)
);
-- Rows: 11975

CREATE TABLE [TbDpDInformeCondicionamiento] (
  [CodProyecto] TEXT(50),
  [CorreoComprador] TEXT(255),
  [RequiereInformeCond] TEXT(50),
  [FechaRequiereInformeCond] SHORT_DATE_TIME(8),
  [UsuarioQuitaObligacionInformeCond] TEXT(50),
  [FechaQuitaObligacionInformeCond] SHORT_DATE_TIME(8),
  [URLCausasCondicionamiento] TEXT(255),
  [FechaCausasEscritas] SHORT_DATE_TIME(8),
  [FechaCartaCondRealizada] SHORT_DATE_TIME(8),
  [CorreoDireccionPara] TEXT(255),
  [CorreoDireccionCC] TEXT(255),
  [CorreoDireccionAsunto] TEXT(255),
  [CorreoDireccionCuerpo] MEMO,
  [URLCartaCondicionamientoRealizada] TEXT(255),
  [URLCartaAutorizacion] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 1134

CREATE TABLE [TbDPDOfertaSuministrador] (
  [NDPD] TEXT(50),
  [FECHAREALIZACIONCARTASOLICITUD] SHORT_DATE_TIME(8),
  [FECHACONVERSIONAPDFSOLICITUD] SHORT_DATE_TIME(8),
  [FECHARECEPCIONOFERTASUMINISTRADOR] SHORT_DATE_TIME(8),
  [OBSERVACIONESSOLICITUDOFERTA] MEMO,
  [OBSERVACIONESRECEPCIONOFERTA] MEMO
);
-- Rows: 2899

CREATE TABLE [TbDPDPlantillasParaCartaSolicitudDeOFerta] (
  [RequisitoBasico] TEXT(255),
  [NombrePlantilla] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 9

CREATE TABLE [TbDpDPrepCartaOfAcepDetalle] (
  [NCarta] LONG(4),
  [NPEDIDO] TEXT(50),
  [NDPD] TEXT(50),
  [TEXTOOFERTA] TEXT(255),
  [IMPORTEOFERTA] DOUBLE(8),
  [FECHACARTA] SHORT_DATE_TIME(8),
  [NAcreedor] LONG(4)
);
-- Rows: 1148

CREATE TABLE [TbDpDPrepCartaOfAcepPrincipal] (
  [NCarta] LONG(4),
  [NAcreedor] LONG(4),
  [DireccionCarta] TEXT(255),
  [CPCarta] TEXT(50),
  [PoblacionCarta] TEXT(255),
  [CodExpediente] TEXT(50),
  [ATT] TEXT(255),
  [NREFERENCIA] TEXT(50),
  [AnexadaCarta] BOOLEAN(1),
  [NDPD] TEXT(255)
);
-- Rows: 860

-- LINKED/SKIPPED: [TbEjercitos] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbEntradas_RES] (given file does not exist: C:\00repos\datos\Registro_Ent_Salida_Datos.accdb)
-- LINKED/SKIPPED: [TbEstados] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExpedientes_1] (
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
  [ResponsableCalidad] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 463

-- LINKED/SKIPPED: [TbExpedientesAnexos] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
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
-- Rows: 223

-- LINKED/SKIPPED: [TbExpedientesAnualidades1] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExpedientesAutorizaciones] (
  [IDAutorizacion] LONG(4),
  [IDExpediente] LONG(4),
  [Motivo] MEMO,
  [NombreAdjunto] TEXT(50),
  [MargenAnterior] DOUBLE(8),
  [MargenNuevo] DOUBLE(8),
  [AutorizadoPor] TEXT(255),
  [FechaAutorizacion] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbExpedientesCodigoCompras] (
  [ID] LONG(4),
  [IDExpediente] LONG(4),
  [CodCompras] TEXT(255)
);
-- Rows: 0

-- LINKED/SKIPPED: [TbExpedientesComerciales] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesCPVs] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExpedientesDocRequeridos] (
  [IDDocRequerido] LONG(4),
  [IDExpediente] LONG(4),
  [NombreDocumento] TEXT(255),
  [FechaPrevista] SHORT_DATE_TIME(8),
  [FechaRealizado] SHORT_DATE_TIME(8),
  [IDDocumento] LONG(4)
);
-- Rows: 158

-- LINKED/SKIPPED: [TbExpedientesJefaturas] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesJuridicas] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesLugaresEjecucion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesPECAL] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbExpedientesProrrogas] (
  [IDProrroga] LONG(4),
  [IDExpediente] LONG(4),
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [BIIVA] DOUBLE(8),
  [BIIPSI] DOUBLE(8),
  [BIIGIC] DOUBLE(8),
  [BIEXENTA] DOUBLE(8),
  [IVA] DOUBLE(8),
  [IPSI] DOUBLE(8),
  [IGIC] DOUBLE(8),
  [PeriodoFacturacion] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 20

-- LINKED/SKIPPED: [TbExpedientesRACS] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesResponsables] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbFacturaInformeMatFalsicado] (
  [IDMF] LONG(4),
  [NFactura] TEXT(255),
  [NPedido] TEXT(255),
  [PH] TEXT(255),
  [EM] TEXT(255),
  [EQ] TEXT(255),
  [ET] TEXT(255),
  [DOC] TEXT(255),
  [DD] TEXT(2),
  [DDRetrasoPlazos] TEXT(2),
  [DDDañosEmbalaje] TEXT(2),
  [DDInsuficienteDoc] TEXT(2),
  [DDIncumplimientoReq] TEXT(2),
  [DDDañoEquipo] TEXT(2),
  [DDFaltanRegistros] TEXT(2),
  [DDEntregaIncompleta] TEXT(2),
  [EF] TEXT(2),
  [EFIR] TEXT(2),
  [EFIC] TEXT(2),
  [Observaciones] MEMO,
  [InspectorNombre] TEXT(255),
  [JPNombre] TEXT(255),
  [FechaInspeccion] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbFacturasCliente] (
  [NFACTURA] TEXT(50),
  [CODEXP] TEXT(50),
  [FECHAEMISION] SHORT_DATE_TIME(8),
  [BASEIMPONIBLE] DOUBLE(8),
  [BASEEXENTA] DOUBLE(8),
  [TIPOIMPOSITIVO] DOUBLE(8),
  [TIPOFACTURA] TEXT(50),
  [OBSERVACIONES] MEMO,
  [IDExpediente] LONG(4),
  [URLFactura] TEXT(255),
  [Grabador] TEXT(255),
  [FechaGrabacion] SHORT_DATE_TIME(8)
);
-- Rows: 926

CREATE TABLE [TbFacturasDetalle] (
  [IDFactura] LONG(4),
  [NPEDIDO] TEXT(50),
  [NFactura] TEXT(50),
  [FechaContabilizacion] SHORT_DATE_TIME(8),
  [FechaFactura] SHORT_DATE_TIME(8),
  [ImporteFactura] DOUBLE(8),
  [ExpAdjudicado] TEXT(50),
  [Observaciones] MEMO,
  [NDOCUMENTO] TEXT(50),
  [SOLPED] TEXT(50),
  [FechaAceptacion] SHORT_DATE_TIME(8),
  [URLFACTURA] TEXT(255),
  [Grabador] TEXT(255),
  [FechaGrabacion] SHORT_DATE_TIME(8),
  [NDPD] TEXT(255),
  [IDDocumentoIV02] LONG(4)
);
-- Rows: 5658

CREATE TABLE [TbFacturasDetalleNoDYS] (
  [IDFacturaNoDYS] LONG(4),
  [NPEDIDO] TEXT(50),
  [NFactura] TEXT(50),
  [FechaContabilizacion] SHORT_DATE_TIME(8),
  [FechaFactura] SHORT_DATE_TIME(8),
  [ImporteFactura] DOUBLE(8),
  [Observaciones] MEMO,
  [NDOCUMENTO] TEXT(50),
  [NAcreedorSAP] LONG(4),
  [NombreArchivo] TEXT(255)
);
-- Rows: 513

CREATE TABLE [TbFacturasDetalleSinPedido] (
  [IDFactura] LONG(4),
  [NPEDIDO] TEXT(50),
  [NFactura] TEXT(50),
  [FechaFactura] SHORT_DATE_TIME(8),
  [ImporteFactura] DOUBLE(8),
  [ExpAdjudicado] TEXT(50),
  [Observaciones] MEMO,
  [NDOCUMENTO] TEXT(50),
  [URLFACTURA] TEXT(255),
  [Grabador] TEXT(255),
  [FechaGrabacion] SHORT_DATE_TIME(8),
  [FechaNPedidoPorTecnico] SHORT_DATE_TIME(8),
  [FechaGrabacionNPedidoTSAP] SHORT_DATE_TIME(8),
  [ObservacionesTecnicas] MEMO,
  [ObservacionesEconomicas] MEMO,
  [UsuarioRedDestinatario] TEXT(255),
  [NAcreedor] LONG(4)
);
-- Rows: 2

CREATE TABLE [TbFacturasPrincipal] (
  [NPedido] TEXT(50),
  [Estado] TEXT(50),
  [SOLPED] TEXT(50),
  [FechaEstado] SHORT_DATE_TIME(8),
  [NDPD] TEXT(255)
);
-- Rows: 2604

-- LINKED/SKIPPED: [TbGradosClasificacion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbGTV] (
  [IDGTV] LONG(4),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [SOLPED] TEXT(50),
  [IMPORTESOLICITADO] DOUBLE(8),
  [DESCRIPCIONGTV] TEXT(255),
  [FECHAINICIAL] SHORT_DATE_TIME(8),
  [FECHAFINAL] SHORT_DATE_TIME(8),
  [CODARTICULO] TEXT(50),
  [ObservacionesGTV] MEMO,
  [URLSolicitud] TEXT(255),
  [FECHASRM] SHORT_DATE_TIME(8),
  [SRM] TEXT(255),
  [FECHAGRABACION] SHORT_DATE_TIME(8)
);
-- Rows: 350

CREATE TABLE [TbGTV_Temp] (
  [Id] LONG(4),
  [NContratoTemp] TEXT(50),
  [PosicionTemp] LONG(4),
  [DescripcionGTVTemp] MEMO,
  [NAcreedorTemp] LONG(4),
  [NombreSuministradorTemp] TEXT(255),
  [ImporteInicialTemp] DOUBLE(8),
  [ImporteRestanteTemp] DOUBLE(8),
  [FechaInicialTemp] SHORT_DATE_TIME(8),
  [FechaFinalTemp] SHORT_DATE_TIME(8),
  [ObservacionesTemp] MEMO
);
-- Rows: 395

CREATE TABLE [TbGTVContratos] (
  [IDGTVContrato] LONG(4),
  [IDGTV] LONG(4),
  [NCONTRATOGTV] TEXT(255),
  [POSICION] LONG(4),
  [IMPORTEADJUDICADO] DOUBLE(8),
  [NAcreedor] LONG(4),
  [URLOferta] TEXT(255),
  [CODOFERTA] TEXT(255)
);
-- Rows: 593

CREATE TABLE [TbGTVContratosExpedientes] (
  [ID] LONG(4),
  [IDGTVContrato] LONG(4),
  [IDExpediente] LONG(4)
);
-- Rows: 594

CREATE TABLE [TbHerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 10

-- LINKED/SKIPPED: [TbJefaturas] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbJuridicas] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbLog] (
  [ID] LONG(4),
  [IDGTV] LONG(4),
  [Titulo] TEXT(255),
  [Accion] MEMO,
  [NDPD] TEXT(255),
  [FechaAccion] SHORT_DATE_TIME(8)
);
-- Rows: 19580

-- LINKED/SKIPPED: [TbLugaresEjecucion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbNPedido] (
  [CODPPD] TEXT(50),
  [CODCONTRATOGTV] TEXT(255),
  [TIPOPEDIDO] TEXT(50),
  [SOLPED] TEXT(50),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [NPEDIDO] TEXT(50),
  [FECHABORRADA] SHORT_DATE_TIME(8),
  [FECHACOMUNICACIONMARILO] SHORT_DATE_TIME(8),
  [FECHANPEDIDO] SHORT_DATE_TIME(8),
  [USUARIO] TEXT(50),
  [FECHACREACION] SHORT_DATE_TIME(8),
  [FECHAMODIFICACION] TEXT(50),
  [IMPORTEADJUDICADO] DOUBLE(8),
  [NAcreedorSAP] LONG(4),
  [EMPRESAADJUDICATARIA] TEXT(255),
  [OBSERVACIONES] TEXT(255),
  [FECHADESPACHOPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDYS] SHORT_DATE_TIME(8),
  [FECHALIBERACIONTSAP] SHORT_DATE_TIME(8),
  [FECHAINICIOTSAP] SHORT_DATE_TIME(8),
  [REQUIERECARTADIRECTOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTADIRECTOR] SHORT_DATE_TIME(8),
  [FECHACARTADIRECTORREALIZADA] SHORT_DATE_TIME(8),
  [REQUIERECARTANPEDIDOSUMINISTRADOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTASUMINISTRADOR] SHORT_DATE_TIME(8),
  [FECHACARTASUMINISTRADORREALIZADA] SHORT_DATE_TIME(8),
  [FECHAANEXOCARTADIRECTORTSAP] SHORT_DATE_TIME(8),
  [FECHAFINTSAP] SHORT_DATE_TIME(8),
  [FECHARECHAZOSOLPED] SHORT_DATE_TIME(8),
  [FECHARECHAZONPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDIRECTOR] SHORT_DATE_TIME(8),
  [SRM] TEXT(50),
  [FECHASRM] SHORT_DATE_TIME(8),
  [FECHAGRABACIONSRM] SHORT_DATE_TIME(8),
  [ESORDINARIO] TEXT(2)
);
-- Rows: 2743

CREATE TABLE [TbNPedidoAntes] (
  [CODPPD] TEXT(50),
  [CODCONTRATOGTV] TEXT(255),
  [TIPOPEDIDO] TEXT(50),
  [SOLPED] TEXT(50),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [NPEDIDO] TEXT(50),
  [FECHABORRADA] SHORT_DATE_TIME(8),
  [FECHACOMUNICACIONMARILO] SHORT_DATE_TIME(8),
  [FECHANPEDIDO] SHORT_DATE_TIME(8),
  [USUARIO] TEXT(50),
  [FECHACREACION] SHORT_DATE_TIME(8),
  [FECHAMODIFICACION] TEXT(50),
  [IMPORTEADJUDICADO] DOUBLE(8),
  [IDEmpresaAdjudicataria] LONG(4),
  [EMPRESAADJUDICATARIA] TEXT(255),
  [OBSERVACIONES] TEXT(255),
  [FECHADESPACHOPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDYS] SHORT_DATE_TIME(8),
  [FECHALIBERACIONTSAP] SHORT_DATE_TIME(8),
  [FECHAINICIOTSAP] SHORT_DATE_TIME(8),
  [REQUIERECARTADIRECTOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTADIRECTOR] SHORT_DATE_TIME(8),
  [FECHACARTADIRECTORREALIZADA] SHORT_DATE_TIME(8),
  [REQUIERECARTANPEDIDOSUMINISTRADOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTASUMINISTRADOR] SHORT_DATE_TIME(8),
  [FECHACARTASUMINISTRADORREALIZADA] SHORT_DATE_TIME(8),
  [FECHAANEXOCARTADIRECTORTSAP] SHORT_DATE_TIME(8),
  [FECHAFINTSAP] SHORT_DATE_TIME(8),
  [FECHARECHAZOSOLPED] SHORT_DATE_TIME(8),
  [FECHARECHAZONPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDIRECTOR] SHORT_DATE_TIME(8)
);
-- Rows: 1499

CREATE TABLE [TbNPedidosHistorica] (
  [CODPPD] TEXT(50),
  [CODCONTRATOGTV] TEXT(255),
  [TIPOPEDIDO] TEXT(50),
  [SOLPED] TEXT(50),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [NPEDIDO] TEXT(50),
  [FECHABORRADA] SHORT_DATE_TIME(8),
  [FECHACOMUNICACIONMARILO] SHORT_DATE_TIME(8),
  [FECHANPEDIDO] SHORT_DATE_TIME(8),
  [USUARIO] TEXT(50),
  [FECHACREACION] SHORT_DATE_TIME(8),
  [FECHAMODIFICACION] TEXT(50),
  [IMPORTEADJUDICADO] DOUBLE(8),
  [NAcreedorSAP] LONG(4),
  [EMPRESAADJUDICATARIA] TEXT(255),
  [OBSERVACIONES] TEXT(255),
  [FECHADESPACHOPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDYS] SHORT_DATE_TIME(8),
  [FECHALIBERACIONTSAP] SHORT_DATE_TIME(8),
  [FECHAINICIOTSAP] SHORT_DATE_TIME(8),
  [REQUIERECARTADIRECTOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTADIRECTOR] SHORT_DATE_TIME(8),
  [FECHACARTADIRECTORREALIZADA] SHORT_DATE_TIME(8),
  [REQUIERECARTANPEDIDOSUMINISTRADOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTASUMINISTRADOR] SHORT_DATE_TIME(8),
  [FECHACARTASUMINISTRADORREALIZADA] SHORT_DATE_TIME(8),
  [FECHAANEXOCARTADIRECTORTSAP] SHORT_DATE_TIME(8),
  [FECHAFINTSAP] SHORT_DATE_TIME(8),
  [FECHARECHAZOSOLPED] SHORT_DATE_TIME(8),
  [FECHARECHAZONPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDIRECTOR] SHORT_DATE_TIME(8)
);
-- Rows: 1

CREATE TABLE [TbNuevosProductos] (
  [Cod] TEXT(255),
  [Articulo] TEXT(255),
  [FechaObsoleto] SHORT_DATE_TIME(8)
);
-- Rows: 407

-- LINKED/SKIPPED: [TbOficinasPrograma] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbOrdinariasTipoAnexos] (
  [IDTipoAnexo] LONG(4),
  [Nombre] TEXT(255),
  [URLRecurso] TEXT(255),
  [RecursoInterno] TEXT(2),
  [ParaMenosDe10K] TEXT(2),
  [ParaMasDe10K] TEXT(2)
);
-- Rows: 8

-- LINKED/SKIPPED: [TbOrganosContratacion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbPECAL] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbProductos] (
  [Codigo] TEXT(50),
  [Nombre] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 408

CREATE TABLE [TbProductos_antes] (
  [Codigo] TEXT(50),
  [Nombre] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 424

CREATE TABLE [TbProductosObsoletos] (
  [CODARTICULO] TEXT(50)
);
-- Rows: 25

CREATE TABLE [TbProductosParaBusqueda] (
  [Codigo] TEXT(255),
  [PalabraClave] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 63

CREATE TABLE [TbProyectos] (
  [ID_PROYECTOS] LONG(4),
  [CODPROYECTOS] TEXT(50),
  [NAcreedorSAP] LONG(4),
  [email] TEXT(255),
  [CODARTICULO] TEXT(50),
  [DESCRIPCION] TEXT(255),
  [EXPEDIENTE] TEXT(50),
  [PETICIONARIO] TEXT(50),
  [PETICIONARIOREAL] TEXT(50),
  [FECHAPETICION] SHORT_DATE_TIME(8),
  [OFERTARECIBIDAFECHA] SHORT_DATE_TIME(8),
  [FECHARECEPCIONECONOMICA] SHORT_DATE_TIME(8),
  [OBSERVACIONESECONOMICAS] MEMO,
  [OBSERVACIONESCALIDAD] TEXT(255),
  [CODOFERTASUMINISTRADOR] TEXT(50),
  [CRITERIOCOMPRAS] TEXT(255),
  [IMPORTESINIVA] DOUBLE(8),
  [OBSERVACIONES] MEMO,
  [FREGISTRO] SHORT_DATE_TIME(8),
  [ELIMINADO] BOOLEAN(1),
  [FELIMINADO] SHORT_DATE_TIME(8),
  [FINICIONECESIDAD] SHORT_DATE_TIME(8),
  [FFINNECESIDAD] SHORT_DATE_TIME(8),
  [CODCONTRATOGTV] TEXT(255),
  [POSICIONCONTRATOGTV] LONG(4),
  [TIPOPEDIDO] TEXT(50),
  [IDExpediente] LONG(4),
  [RistraResponsablesExpTemp] TEXT(255),
  [FechaFinAgendaTecnica] SHORT_DATE_TIME(8),
  [EstadoCorto] TEXT(255),
  [EstadoLargo] TEXT(255),
  [FechaPreparadoParaVisadoOferta] SHORT_DATE_TIME(8),
  [FechaVisadoOferta] SHORT_DATE_TIME(8),
  [FechaRechazoOferta] SHORT_DATE_TIME(8),
  [ELIMINADOOBSERVACIONES] MEMO
);
-- Rows: 2971

CREATE TABLE [TbProyectosDocumentosParaPedidoOrdinario] (
  [NDPD] TEXT(255),
  [Cesta] TEXT(255),
  [IDTipoDocumento] LONG(4),
  [URLFinal] TEXT(255),
  [FechaSolicitado] SHORT_DATE_TIME(8),
  [FechaEntregado] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 0

CREATE TABLE [TbProyectosHistoricos] (
  [ID_PROYECTOS] LONG(4),
  [CODPROYECTOS] TEXT(50),
  [NAcreedorSAP] LONG(4),
  [CODCENTRO] TEXT(50),
  [NOMBRECENTRO] TEXT(50),
  [CODARTICULO] TEXT(50),
  [DEPENDENCIA] TEXT(50),
  [DESCRIPCION] TEXT(255),
  [EXPEDIENTE] TEXT(50),
  [PETICIONARIO] TEXT(50),
  [FECHAPETICION] SHORT_DATE_TIME(8),
  [OFERTARECIBIDA] BOOLEAN(1),
  [OFERTARECIBIDAFECHA] SHORT_DATE_TIME(8),
  [FECHARECEPCIONECONOMICA] SHORT_DATE_TIME(8),
  [OBSERVACIONESECONOMICAS] MEMO,
  [OFERTAACEPTADA] BOOLEAN(1),
  [OFERTAACEPTADAFECHA] SHORT_DATE_TIME(8),
  [OBRAREPAROS] BOOLEAN(1),
  [REPAROSSUBSANADOS] BOOLEAN(1),
  [FECHACERTIFICADOCONFORMIDAD] SHORT_DATE_TIME(8),
  [TEXTOCERTIFICADOCONFORMIDAD] TEXT(255),
  [OBSERVACIONESCALIDAD] TEXT(255),
  [FECHAREPAROSSUBSANADOS] SHORT_DATE_TIME(8),
  [FECHAOBRAENTREGASINREPAROS] SHORT_DATE_TIME(8),
  [CODOFERTASUMINISTRADOR] TEXT(50),
  [CRITERIOCOMPRAS] TEXT(255),
  [IMPORTESINIVA] DOUBLE(8),
  [FACTURAPAGADA] BOOLEAN(1),
  [OBSERVACIONES] MEMO,
  [FREGISTRO] SHORT_DATE_TIME(8),
  [CENTRO] TEXT(50),
  [ELIMINADO] BOOLEAN(1),
  [FELIMINADO] SHORT_DATE_TIME(8),
  [FINICIONECESIDAD] SHORT_DATE_TIME(8),
  [FFINNECESIDAD] SHORT_DATE_TIME(8),
  [CONTINUAENVERSIONDOS] BOOLEAN(1),
  [AntesDeVersion2] BOOLEAN(1),
  [AÑOPRESUPUESTO] LONG(4),
  [FIRMADELTECNICO] TEXT(255),
  [CODCONTRATOGTV] TEXT(255),
  [TIPOPEDIDO] TEXT(50),
  [AnexoOfertaSuministrador] TEXT(255),
  [FechaAnexoOfertaSuministrador] SHORT_DATE_TIME(8),
  [IDSuministrador] LONG(4)
);
-- Rows: 2047

-- LINKED/SKIPPED: [TbRACS] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbRequisitosSTG] (
  [NDPD] TEXT(50),
  [A0] TEXT(50),
  [A1] TEXT(2),
  [A2] TEXT(2),
  [A3] TEXT(2),
  [A4] TEXT(2),
  [B00] TEXT(50),
  [B01] TEXT(2),
  [B02] TEXT(2),
  [B03] TEXT(2),
  [B1] TEXT(50),
  [B2] TEXT(50),
  [B21] TEXT(2),
  [B22] TEXT(2),
  [B23] TEXT(2),
  [B24] TEXT(2),
  [B25] TEXT(2),
  [B26] TEXT(2),
  [C0] TEXT(50),
  [C1] TEXT(2),
  [C2] TEXT(2),
  [C3] TEXT(2),
  [D0] TEXT(50),
  [D1] TEXT(2),
  [D2] TEXT(2),
  [D3] TEXT(2),
  [D4] TEXT(2),
  [E0] TEXT(50),
  [E1] TEXT(2),
  [E2] TEXT(2),
  [E3] TEXT(2),
  [F0] TEXT(50),
  [F1] TEXT(2),
  [F2] TEXT(2),
  [F3] TEXT(2),
  [F4] TEXT(2),
  [F5] TEXT(2),
  [F6] TEXT(2),
  [F7] TEXT(2),
  [G0] TEXT(50),
  [G1] TEXT(2),
  [G2] TEXT(2),
  [G3] TEXT(2),
  [H0] TEXT(50),
  [I0] TEXT(50),
  [I1] TEXT(2),
  [I2] TEXT(2),
  [I3] TEXT(2),
  [I4] TEXT(2),
  [J0] TEXT(50),
  [J1] TEXT(2),
  [J2] TEXT(2),
  [J3] TEXT(2),
  [J4] TEXT(2),
  [J5] TEXT(2),
  [K0] TEXT(50),
  [K1] TEXT(2),
  [K2] TEXT(2),
  [K3] TEXT(2),
  [K4] TEXT(2),
  [L0] TEXT(50),
  [L1] TEXT(255),
  [L2] TEXT(255),
  [L3] TEXT(255),
  [L4] TEXT(2),
  [M0] TEXT(50),
  [M1] TEXT(2),
  [M2] TEXT(2),
  [M3] TEXT(2),
  [M4] TEXT(2),
  [M5] TEXT(2),
  [M6] TEXT(2),
  [M7] TEXT(255)
);
-- Rows: 227

CREATE TABLE [TbResponsablesExpedientes] (
  [IdExpediente] LONG(4),
  [IdUsuario] LONG(4),
  [CorreoSiempre] TEXT(2)
);
-- Rows: 996

-- LINKED/SKIPPED: [TbSalidas_RES] (given file does not exist: C:\00repos\datos\Registro_Ent_Salida_Datos.accdb)
CREATE TABLE [TbSolicitudesEnvioAlSuministrador] (
  [IDEnvio] LONG(4),
  [IDSO] LONG(4),
  [NDPD] TEXT(255),
  [NSalida] TEXT(255),
  [IDSalida] LONG(4),
  [NombreTabla] TEXT(255),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [IDCorreo] LONG(4),
  [emailSuministrador] MEMO
);
-- Rows: 1698

CREATE TABLE [TbSolicitudesOfertasPrevias] (
  [IDSO] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [DPD] TEXT(255),
  [NAcreedor] LONG(4),
  [IDCorreo] LONG(4),
  [IDSalida] LONG(4),
  [NombreTabla] TEXT(255),
  [NombreArchivo] TEXT(255),
  [UsuarioRegistra] TEXT(255),
  [FechaModificacion] SHORT_DATE_TIME(8),
  [email] TEXT(255)
);
-- Rows: 2237

-- LINKED/SKIPPED: [TbSolicitudesOfertasPrevias1] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbSTGParametrosParaFormulario] (
  [NombreMarco] TEXT(255),
  [Letras] TEXT(255),
  [Nombre] TEXT(255),
  [NumeroCasillas] BYTE(1)
);
-- Rows: 15

CREATE TABLE [TbSuministradoresDireccion] (
  [IdSuministrador] LONG(4),
  [LineaDireccion1] TEXT(255),
  [LineaDireccion2] TEXT(255),
  [LineaDireccion3] TEXT(255),
  [LineaDireccion4] TEXT(255)
);
-- Rows: 24

CREATE TABLE [TbSuministradoresListaFavoritos] (
  [UsuarioRed] TEXT(255),
  [NombreLista] TEXT(255),
  [NAcreedor] TEXT(255),
  [email] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbSuministradoresNotasReferencia] (
  [Nota] TEXT(50),
  [PuntuacionMinima] LONG(4),
  [PuntuacionMaxima] LONG(4)
);
-- Rows: 4

CREATE TABLE [TbSuministradoresPuntuacion] (
  [IDPuntuacion] LONG(4),
  [IDSuministradorSAP] LONG(4),
  [NPedido] TEXT(50),
  [NFactura] TEXT(50),
  [NDPD] TEXT(50),
  [PuntuacionCalidad] INT(2),
  [PuntuacionServicio] INT(2),
  [Usuario] TEXT(50),
  [FechaPuntuacion] SHORT_DATE_TIME(8)
);
-- Rows: 2450

CREATE TABLE [TbSuministradoresPuntuacionConDetalle] (
  [IDPuntuacion] LONG(4),
  [IDVisado] LONG(4),
  [IDSuministradorSAP] LONG(4),
  [NPedido] TEXT(50),
  [NFactura] TEXT(50),
  [NDPD] TEXT(50),
  [CA] INT(2),
  [CB] INT(2),
  [SA] LONG(4),
  [SB] INT(2),
  [SC] INT(2),
  [SD] INT(2),
  [SE] INT(2),
  [SF] INT(2),
  [SG] INT(2),
  [SH] INT(2),
  [PuntuacionCalidad] TEXT(255),
  [PuntuacionServicio] TEXT(255),
  [Usuario] TEXT(50),
  [FechaPuntuacion] SHORT_DATE_TIME(8),
  [ObservacionesPuntuacion] MEMO,
  [NDOCUMENTO] TEXT(50)
);
-- Rows: 5551

CREATE TABLE [TbSuministradoresSAP] (
  [IDSuministrador] LONG(4),
  [NifComunitario] TEXT(50),
  [CP] TEXT(50),
  [Poblacion] TEXT(255),
  [Suministrador] TEXT(255),
  [AcreedorSAP] LONG(4),
  [Direccion] TEXT(255),
  [email] TEXT(255),
  [ConceptoBusqueda] MEMO,
  [PerteneceAlGrupo] TEXT(2)
);
-- Rows: 50038

CREATE TABLE [TbTareasTecnicasExplicaciones] (
  [NodoTarea] TEXT(255),
  [TituloTarea] TEXT(255),
  [Explicacion] MEMO
);
-- Rows: 8

CREATE TABLE [TbTemp_PedidosEstado] (
  [NDPD] TEXT(50),
  [EstadoPedido] TEXT(50)
);
-- Rows: 2992

CREATE TABLE [TbTempDPDHTML] (
  [NDPD] TEXT(255),
  [PETICIONARIO] TEXT(255),
  [FECHAPETICION] SHORT_DATE_TIME(8),
  [DESCRIPCION] MEMO,
  [EXPEDIENTE] TEXT(255),
  [IMPORTESOLICITADO] MONEY(8),
  [SUMINISTRADORSOLICITADO] TEXT(255),
  [SUMINISTRADORSOLICITADOEMAIL] TEXT(255),
  [SUMINISTRADORSOLICITADOCIF] TEXT(255),
  [TIPOPEDIDO] TEXT(255),
  [RequiereInformeCond] TEXT(2),
  [CODCONTRATOGTV] TEXT(255),
  [POSICIONCONTRATOGTV] TEXT(255),
  [URLCausasCondicionamiento] TEXT(255),
  [FechaCausasEscritas] SHORT_DATE_TIME(8),
  [SOFechaRealiza] SHORT_DATE_TIME(8),
  [SOFechaEnvioSecretaria] SHORT_DATE_TIME(8),
  [SOFechaEnvio] SHORT_DATE_TIME(8),
  [ROFechaRealiza] SHORT_DATE_TIME(8),
  [ROFechaVisado] SHORT_DATE_TIME(8),
  [ROFechaRechazo] SHORT_DATE_TIME(8),
  [ROObservacionesRechazo] MEMO,
  [ROObservacionesVisado] MEMO,
  [FechaFinAgendaTecnica] SHORT_DATE_TIME(8),
  [FECHARECEPCIONECONOMICA] SHORT_DATE_TIME(8),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [SOLPED] TEXT(255),
  [FECHANPEDIDO] SHORT_DATE_TIME(8),
  [NPEDIDO] TEXT(255),
  [IMPORTEADJUDICADO] MONEY(8),
  [SUMINISTRADORADJUDICADO] TEXT(255),
  [SUMINISTRADORADJUDICADOEMAIL] TEXT(255),
  [SUMINISTRADORADJUDICADOCIF] TEXT(255),
  [ADFechaRealiza] SHORT_DATE_TIME(8),
  [ADFechaEnvioSecretaria] SHORT_DATE_TIME(8),
  [FechaMarcadaParaNoEnvio] SHORT_DATE_TIME(8),
  [ADFechaEnvio] SHORT_DATE_TIME(8),
  [IMPORTEPORFACTURAR] MONEY(8),
  [ESTADOPEDIDO] TEXT(255),
  [FELIMINADO] SHORT_DATE_TIME(8),
  [IMPORTERESTANTEENCONTRATO] MONEY(8)
);
-- Rows: 193

CREATE TABLE [TbTipoDocumento] (
  [TipoDocumento] TEXT(255),
  [Unico] BOOLEAN(1),
  [SiglaParaDocumentacion] TEXT(3),
  [TituloDocumentoPorDefecto] TEXT(50),
  [ordinal] LONG(4)
);
-- Rows: 18

CREATE TABLE [TbTiposImpositivos] (
  [TipoImpositivo] DOUBLE(8),
  [Descripcion] TEXT(255)
);
-- Rows: 8

CREATE TABLE [TbTiposPedido] (
  [TipoPedido] TEXT(50),
  [DescripcionTipoPedido] TEXT(255)
);
-- Rows: 3

CREATE TABLE [TbTmp_DPDs] (
  [NDPD_Temp] TEXT(50),
  [EstadoDPD_Temp] TEXT(50)
);
-- Rows: 2385

CREATE TABLE [TbTmp_Facturas] (
  [IDFacturaTmp] LONG(4),
  [NFacturaTmp] TEXT(50),
  [NPedidoTmp] TEXT(50),
  [EstadoTmp] TEXT(50),
  [ObservacionesTmp] MEMO,
  [CodExpTmp] TEXT(50),
  [CodDPDTemp] TEXT(50),
  [DescripcionDPDTemp] TEXT(255),
  [URLFacturaTemp] TEXT(255),
  [NDocumentoTemp] TEXT(255),
  [ImporteFacturaTemp] TEXT(50),
  [RistraResponsablesExpTemp] TEXT(255),
  [PeticionarioTemp] TEXT(255),
  [PeticionarioLargoTemp] TEXT(255),
  [FechaEnvioCorreo] SHORT_DATE_TIME(8),
  [GrabadorTemp] TEXT(255),
  [FechaGrabacionTemp] SHORT_DATE_TIME(8),
  [FechaFacturaTemp] SHORT_DATE_TIME(8),
  [CadenaIDCorreosNotasDeCalidad] TEXT(255),
  [NAcreedorSAP] LONG(4),
  [ObsRechazoTecnico] MEMO,
  [ObsRechazoCalidad] MEMO,
  [FVISADOECONOMICOTemp] SHORT_DATE_TIME(8),
  [FRECHAZOECONOMICOTemp] SHORT_DATE_TIME(8),
  [FVISADOCALIDADTemp] SHORT_DATE_TIME(8),
  [FRECHAZOCALIDADTemp] SHORT_DATE_TIME(8),
  [FVISADOTECNICOTemp] SHORT_DATE_TIME(8),
  [FRECHAZOTECNICOTemp] SHORT_DATE_TIME(8),
  [FECHAPETICIONTemp] SHORT_DATE_TIME(8)
);
-- Rows: 4986

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbUsuariosEnvioTareas] (
  [Usuario] TEXT(255),
  [DiasEnvioTareas] TEXT(20)
);
-- Rows: 21

CREATE TABLE [TbVisadoFacturas] (
  [NDPD] TEXT(50),
  [NDOCUMENTOFACTURA] TEXT(50),
  [FVISADOTECNICO] SHORT_DATE_TIME(8),
  [USUARIOQUEVISA] TEXT(50),
  [OBSERVACIONESVISADOFACTURA] MEMO,
  [FRECHAZOTECNICO] SHORT_DATE_TIME(8),
  [USUARIOQUERECHAZA] TEXT(50),
  [OBSERVACIONESRECHAZO] MEMO,
  [NFactura] TEXT(50),
  [NPEDIDO] TEXT(50),
  [UsuarioCalidad] TEXT(50),
  [FechaVisadoCalidad] SHORT_DATE_TIME(8),
  [ObservacionesVisadoCalidad] MEMO,
  [DevueltoPorCalidad] BOOLEAN(1)
);
-- Rows: 1307

CREATE TABLE [TbVisadoFacturas_Nueva] (
  [IDVisado] LONG(4),
  [IDFactura] LONG(4),
  [NPEDIDO] TEXT(50),
  [NFactura] TEXT(50),
  [FVISADOTECNICO] SHORT_DATE_TIME(8),
  [FRECHAZOTECNICO] SHORT_DATE_TIME(8),
  [USUARIOTECNICOVISA] TEXT(50),
  [USUARIOTECNICORECHAZA] TEXT(50),
  [OBSERVACIONESTECNICASVISA] MEMO,
  [OBSERVACIONESTECNICASRECHAZA] MEMO,
  [FVISADOCALIDAD] SHORT_DATE_TIME(8),
  [FRECHAZOCALIDAD] SHORT_DATE_TIME(8),
  [USUARIOCALIDADVISA] TEXT(50),
  [USUARIOCALIDADRECHAZA] TEXT(50),
  [OBSERVACIONESCALIDADVISA] MEMO,
  [OBSERVACIONESCALIDADRECHAZA] MEMO,
  [OBSERVACIONESCALIDADPARAECONOMIA] MEMO,
  [FVISADOECONOMICO] SHORT_DATE_TIME(8),
  [FRECHAZOECONOMICO] SHORT_DATE_TIME(8),
  [USUARIOECONOMICOVISA] TEXT(50),
  [USUARIOECONOMICORECHAZA] TEXT(50),
  [OBSERVACIONESECONOMICOVISA] MEMO,
  [OBSERVACIONESECONOMICORECHAZA] MEMO,
  [NDPD] TEXT(255),
  [NDOCUMENTO] TEXT(255)
);
-- Rows: 5571

CREATE TABLE [TbVisadosGenerales] (
  [NDPD] TEXT(50),
  [SOUsuarioRealiza] TEXT(50),
  [SOFechaRealiza] SHORT_DATE_TIME(8),
  [SOFechaVisado] SHORT_DATE_TIME(8),
  [SOFechaActualiza] SHORT_DATE_TIME(8),
  [SOUsuarioVisado] TEXT(50),
  [SOObservacionesVisado] MEMO,
  [SOFechaEnvioSuministrador] SHORT_DATE_TIME(8),
  [SOIDCorreoAlSuministrador] LONG(4),
  [SOFechaRechazo] SHORT_DATE_TIME(8),
  [SOUsuarioRechazo] TEXT(50),
  [SOObservacionesRechazo] MEMO,
  [SOFechaEnvioSecretaria] SHORT_DATE_TIME(8),
  [SORevisadoNoEnvio] TEXT(2),
  [ROUsuarioRealiza] TEXT(50),
  [ROFechaRealiza] SHORT_DATE_TIME(8),
  [ROFechaVisado] SHORT_DATE_TIME(8),
  [ROFechaActualiza] SHORT_DATE_TIME(8),
  [ROUsuarioVisado] TEXT(50),
  [ROObservacionesVisado] MEMO,
  [ROFechaRechazo] SHORT_DATE_TIME(8),
  [ROUsuarioRechazo] TEXT(50),
  [ROObservacionesRechazo] MEMO,
  [ADUsuarioRealiza] TEXT(50),
  [ADFechaRealiza] SHORT_DATE_TIME(8),
  [ADFechaVisado] SHORT_DATE_TIME(8),
  [ADFechaActualiza] SHORT_DATE_TIME(8),
  [ADUsuarioVisado] TEXT(50),
  [ADObservacionesVisado] MEMO,
  [ADFechaEnvioSuministrador] SHORT_DATE_TIME(8),
  [ADIDCorreoAlSuministrador] LONG(4),
  [ADFechaRechazo] SHORT_DATE_TIME(8),
  [ADUsuarioRechazo] TEXT(50),
  [ADObservacionesRechazo] MEMO,
  [ADFechaEnvioSecretaria] SHORT_DATE_TIME(8),
  [ADRevisadoNoEnvio] TEXT(2),
  [EstadoVisadoSO] TEXT(255),
  [EstadoVisadoRO] TEXT(255),
  [EstadoVisadoAD] TEXT(255)
);
-- Rows: 2278

CREATE TABLE [~TMPCLP122541] (
  [F1] DOUBLE(8),
  [F2] SHORT_DATE_TIME(8),
  [F3] DOUBLE(8),
  [F4] TEXT(255)
);
-- Rows: 1

CREATE TABLE [~TMPCLP175201] (
  [CODPPD] TEXT(50),
  [CODCONTRATOGTV] TEXT(255),
  [TIPOPEDIDO] TEXT(50),
  [SOLPED] TEXT(50),
  [FECHASOLPED] SHORT_DATE_TIME(8),
  [NPEDIDO] TEXT(50),
  [FECHABORRADA] SHORT_DATE_TIME(8),
  [FECHACOMUNICACIONMARILO] SHORT_DATE_TIME(8),
  [FECHANPEDIDO] SHORT_DATE_TIME(8),
  [USUARIO] TEXT(50),
  [FECHACREACION] SHORT_DATE_TIME(8),
  [FECHAMODIFICACION] TEXT(50),
  [IMPORTEADJUDICADO] DOUBLE(8),
  [NAcreedorSAP] LONG(4),
  [EMPRESAADJUDICATARIA] TEXT(255),
  [OBSERVACIONES] TEXT(255),
  [FECHADESPACHOPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDYS] SHORT_DATE_TIME(8),
  [FECHALIBERACIONTSAP] SHORT_DATE_TIME(8),
  [FECHAINICIOTSAP] SHORT_DATE_TIME(8),
  [REQUIERECARTADIRECTOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTADIRECTOR] SHORT_DATE_TIME(8),
  [FECHACARTADIRECTORREALIZADA] SHORT_DATE_TIME(8),
  [REQUIERECARTANPEDIDOSUMINISTRADOR] BOOLEAN(1),
  [FECHACOMUNICACIONCARTASUMINISTRADOR] SHORT_DATE_TIME(8),
  [FECHACARTASUMINISTRADORREALIZADA] SHORT_DATE_TIME(8),
  [FECHAANEXOCARTADIRECTORTSAP] SHORT_DATE_TIME(8),
  [FECHAFINTSAP] SHORT_DATE_TIME(8),
  [FECHARECHAZOSOLPED] SHORT_DATE_TIME(8),
  [FECHARECHAZONPEDIDO] SHORT_DATE_TIME(8),
  [FECHALIBERACIONDIRECTOR] SHORT_DATE_TIME(8),
  [SRM] TEXT(50),
  [FECHASRM] SHORT_DATE_TIME(8),
  [FECHAGRABACIONSRM] SHORT_DATE_TIME(8),
  [ESORDINARIO] TEXT(2)
);
-- Rows: 2419

-- Summary: 104 local tables, 35 linked/skipped
