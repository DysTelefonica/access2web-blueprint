-- Schema dump: GestionContratos_datos.accdb
-- Tables: 20

CREATE TABLE [AnexoContra] (
  [Nexpediente] TEXT(255),
  [NombreAnexo] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 149

CREATE TABLE [AnexoDeri] (
  [Nexpe] TEXT(255),
  [ExpeLote] TEXT(255),
  [ExpeDeri] TEXT(255),
  [NombreAnexo] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 265

CREATE TABLE [AnexoFases] (
  [id] LONG(4),
  [NExpediente] TEXT(255),
  [NExpcorto] TEXT(255),
  [NExplote] TEXT(255),
  [NExpder] TEXT(255),
  [Fase] TEXT(255),
  [TipoDAO] TEXT(255),
  [Nombre] TEXT(255)
);
-- Rows: 1067

CREATE TABLE [AnexoLotes] (
  [Nexpe] TEXT(255),
  [ExpeLote] TEXT(255),
  [NombreAnexo] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 49

CREATE TABLE [AnexoSA] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoSA] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 5

CREATE TABLE [AnexoSCCC] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoCCC] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 82

CREATE TABLE [AnexoSCS] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoCS] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 1

CREATE TABLE [AnexoSFS] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoFS] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 11

CREATE TABLE [AnexoSFT] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoFT] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 20

CREATE TABLE [AnexoSGC] (
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [AnexoGC] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 2

CREATE TABLE [Contratos] (
  [Id] LONG(4),
  [Contrato] TEXT(255),
  [JuridicaContrata] TEXT(255),
  [TipoExpediente] TEXT(255),
  [NumeroExpediente] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [Jefe de Proyecto] TEXT(255),
  [Clasificacion] TEXT(255),
  [OrganoContratacion] TEXT(255),
  [FechaFirma] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [Pendientecontrato] BOOLEAN(1),
  [Estado] TEXT(255),
  [NombreJS] TEXT(255),
  [EmailJS] TEXT(255),
  [NombreInspector] TEXT(255),
  [EmailInspector] TEXT(255),
  [comercial] TEXT(255),
  [EJuridica] TEXT(255),
  [NEjecutamosNosotros] BOOLEAN(1)
);
-- Rows: 114

CREATE TABLE [Derivados] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [RefLote] TEXT(255),
  [RefDer] TEXT(255),
  [FechaFirmaDe] SHORT_DATE_TIME(8),
  [FechaFinDe] SHORT_DATE_TIME(8),
  [Contacto] TEXT(255),
  [Observaciones] TEXT(255),
  [Nderivado] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 157

CREATE TABLE [FaseEje] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [Nexplote] TEXT(255),
  [Nexpder] TEXT(255),
  [CCC] BOOLEAN(1),
  [ClausulaSeguridad] BOOLEAN(1),
  [GiaClasi] BOOLEAN(1),
  [FirmaContratacion] BOOLEAN(1),
  [ObFirmaContratacion] TEXT(255),
  [FirmaTelefonica] BOOLEAN(1),
  [ObFirmaTelefonica] TEXT(255),
  [EnvioContratacion] BOOLEAN(1),
  [FechaEnvioContratacion] SHORT_DATE_TIME(8),
  [EnvioASI] BOOLEAN(1),
  [FechaEnvioASI] SHORT_DATE_TIME(8),
  [SolicitudAcceso] BOOLEAN(1),
  [AutorizacionAcceso] BOOLEAN(1),
  [FechaAutorizacion] SHORT_DATE_TIME(8),
  [SolicitadoSubcontrata] BOOLEAN(1),
  [Observaciones] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255),
  [EnviadoJP] BOOLEAN(1),
  [FechaEnviadoJP] SHORT_DATE_TIME(8)
);
-- Rows: 282

CREATE TABLE [FaseOferta] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [SolicitudPliegoASI] BOOLEAN(1),
  [ObSolicitudPliegoASI] TEXT(255),
  [CertificadoHSEMdeASI] BOOLEAN(1),
  [ObCertificadoHSEMdeASI] TEXT(255),
  [DeclaracionHSEMdeJSSP] BOOLEAN(1),
  [ObDeclaracionHSEMdeJSSP] TEXT(255),
  [CertificadoHPSdeASI] BOOLEAN(1),
  [ObCertificadoHPSdeASI] TEXT(255),
  [DeclaracionHPSdeJSSP] BOOLEAN(1),
  [ObDeclaracionHPSdeJSSP] TEXT(255),
  [NP1] BOOLEAN(1),
  [ObNP1] TEXT(255),
  [NP2] BOOLEAN(1),
  [ObNP2] TEXT(255),
  [ComentariosGenerales] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 225

CREATE TABLE [Lotes] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [NEXPLOTE] TEXT(255),
  [NLOTE] TEXT(255),
  [FechaFirmaLo] SHORT_DATE_TIME(8),
  [FechaFinLo] SHORT_DATE_TIME(8),
  [Contacto] TEXT(255),
  [Observaciones] TEXT(255),
  [NombreLote] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 33

CREATE TABLE [Subcontrata] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [NexpLote] TEXT(255),
  [Nexpder] TEXT(255),
  [Nombre] TEXT(255),
  [CCC] BOOLEAN(1),
  [ClausulaSeguridad] BOOLEAN(1),
  [GiaClasi] BOOLEAN(1),
  [FirmaSubContrata] BOOLEAN(1),
  [ObFirmaSubContrata] TEXT(255),
  [FirmaTelefonica] BOOLEAN(1),
  [ObFirmaTelefonica] TEXT(255),
  [EnvioASI] BOOLEAN(1),
  [FechaEnvioASI] SHORT_DATE_TIME(8),
  [SolicitudAcceso] BOOLEAN(1),
  [NumeroExpedienteCorto] TEXT(255)
);
-- Rows: 58

CREATE TABLE [TbConsultaContratosGlobal] (
  [id] TEXT(255),
  [contrato] TEXT(255),
  [JuridicaContrata] TEXT(255),
  [TipoExpediente] TEXT(255),
  [NumeroExpediente] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [JefedeProyecto] TEXT(255),
  [Clasificacion] TEXT(255),
  [OrganoContratacion] TEXT(255),
  [FechaFirma] TEXT(255),
  [FechaFin] TEXT(255),
  [Pendientecontrato] TEXT(255),
  [Estado] TEXT(255),
  [NombreJS] TEXT(255),
  [EmailJS] TEXT(255),
  [NombreInspector] TEXT(255),
  [EmailInspector] TEXT(255),
  [comercial] TEXT(255),
  [EJuridica] TEXT(255),
  [IDLote] TEXT(255),
  [NEXPLOTE] TEXT(255),
  [NLOTE] TEXT(255),
  [FechaFirmaLo] TEXT(255),
  [FechaFinLo] TEXT(255),
  [Contacto] TEXT(255),
  [Observaciones_Lote] TEXT(255),
  [NombreLote] TEXT(255),
  [Id_C] TEXT(255),
  [RefLote_C] TEXT(255),
  [RefDer_C] TEXT(255),
  [FechaFirmaDe_C] TEXT(255),
  [FechaFinDe_C] TEXT(255),
  [Contacto_C] TEXT(255),
  [Observaciones_C] TEXT(255),
  [Nderivado_C] TEXT(255),
  [Id_L] TEXT(255),
  [RefLote_L] TEXT(255),
  [RefDer_L] TEXT(255),
  [FechaFirmaDe_L] TEXT(255),
  [FechaFinDe_L] TEXT(255),
  [Contacto_L] TEXT(255),
  [Observaciones_L] TEXT(255),
  [Nderivado_L] TEXT(255),
  [id_Feje] TEXT(255),
  [CCC] TEXT(255),
  [CSeg] TEXT(255),
  [GiaClasi] TEXT(255),
  [FirCont] TEXT(255),
  [FirTele] TEXT(255),
  [EnvCon] TEXT(255),
  [EnvAsi] TEXT(255),
  [AnexoVIII] TEXT(255),
  [AutAcc] TEXT(255),
  [AutSub] TEXT(255),
  [Obser_Feje] TEXT(255),
  [NContrata] TEXT(255),
  [NExploteContrata] TEXT(255),
  [NExpderContrata] TEXT(255),
  [CCCcontrata] TEXT(255),
  [CSegcontrata] TEXT(255),
  [GiaClasicontrata] TEXT(255),
  [FirSubConcontrata] TEXT(255),
  [FirTelecontrata] TEXT(255),
  [EnvAsicontrata] TEXT(255),
  [AnexoVIIIcontrata] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCorreosEnviados] (
  [IDCorreo] LONG(4),
  [Originador] TEXT(255),
  [Destinatarios] MEMO,
  [DestinatariosConCopia] MEMO,
  [DestinatariosConCopiaOculta] TEXT(255),
  [Asunto] TEXT(255),
  [Cuerpo] MEMO,
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaGrabacion] SHORT_DATE_TIME(8)
);
-- Rows: 204

CREATE TABLE [TbParaMarioloExpedientes] (
  [NExpediente] MEMO,
  [NExpedienteLargo] MEMO,
  [NExpedienteCorto] MEMO
);
-- Rows: 67

CREATE TABLE [TContratos] (
  [Id] LONG(4),
  [NumeroExpediente] TEXT(255),
  [NumeroExpedienteCorto] TEXT(255),
  [Contrato] TEXT(255),
  [JuridicaContrata] TEXT(255),
  [TipoExpediente] TEXT(255),
  [ReferenciaExpediente] TEXT(255),
  [Jefe de Proyecto] TEXT(255),
  [Clasificacion] TEXT(255),
  [OrganoContratacion] TEXT(255),
  [FechaFirma] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [Pendientecontrato] BOOLEAN(1),
  [Estado] TEXT(255),
  [NombreJS] TEXT(255),
  [EmailJS] TEXT(255),
  [NombreInspector] TEXT(255),
  [EmailInspector] TEXT(255),
  [comercial] TEXT(255),
  [EJuridica] TEXT(255),
  [NEXPLOTE] TEXT(255),
  [RefDer] TEXT(255),
  [Ncontrata] TEXT(255),
  [NEjecutamosNosotros] BOOLEAN(1)
);
-- Rows: 50

-- Summary: 20 local tables, 0 linked/skipped
