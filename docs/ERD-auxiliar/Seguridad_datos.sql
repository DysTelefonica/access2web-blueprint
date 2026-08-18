-- Schema dump: Seguridad_datos.accdb
-- Tables: 36

CREATE TABLE [Tb0Adjuntos] (
  [IDAdjunto] LONG(4),
  [IDRegistro] LONG(4),
  [IDTipoRegistro] LONG(4),
  [FechaAdjunto] SHORT_DATE_TIME(8),
  [NombreAdjuntoLocal] TEXT(255),
  [NombreAdjunto] TEXT(255),
  [AdjuntoEncriptado] TEXT(2)
);
-- Rows: 785

CREATE TABLE [Tb0AdjuntosIndexaciones] (
  [IDAdjunto] LONG(4),
  [Hash] TEXT(255),
  [FechaIndexacion] SHORT_DATE_TIME(8),
  [TextoAdjunto] MEMO
);
-- Rows: 1

CREATE TABLE [Tb0Conexiones] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 3

CREATE TABLE [Tb0HerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 4

CREATE TABLE [Tb0RegistrosPendientes] (
  [IDRegistroPendiente] LONG(4),
  [IDRegistro] LONG(4),
  [IDTipoRegistro] LONG(4),
  [FechaPendiente] SHORT_DATE_TIME(8),
  [FechaPrevisto] SHORT_DATE_TIME(8),
  [FechaCompletado] SHORT_DATE_TIME(8),
  [DescripcionPendiente] MEMO,
  [CorreoDestinatarios] MEMO
);
-- Rows: 2

CREATE TABLE [Tb0RegistrosRelacionados] (
  [IDRegistroRelacion] LONG(4),
  [IDTipoRegistroARelacionar] LONG(4),
  [IDRegistroARelacionar] LONG(4),
  [IDTipoRegistroRelacionado] LONG(4),
  [IDRegistroRelacionado] LONG(4)
);
-- Rows: 133

CREATE TABLE [Tb0TipoRegistros] (
  [IDTipoRegistro] LONG(4),
  [TipoRegistro] TEXT(255),
  [NombreCampoSubTipoRegistro] TEXT(255),
  [AliasTipoRegistro] TEXT(255),
  [NombreTabla] TEXT(255),
  [NombreID] TEXT(255),
  [Descripción] MEMO,
  [NombreFormulario] TEXT(255),
  [NombreCampoAsunto] TEXT(255),
  [NombreCampoNotas] TEXT(255),
  [NombreCampoFecha] TEXT(255),
  [AnchoNombre] DOUBLE(8),
  [Acabado] TEXT(2)
);
-- Rows: 21

CREATE TABLE [Tb0TipoRegistrosDescripcion] (
  [IDTipoRegistro] LONG(4),
  [Descripcion] MEMO
);
-- Rows: 21

CREATE TABLE [Tb0TipoRegistroUso] (
  [IDTipoRegistro] LONG(4),
  [Usuario] TEXT(255),
  [Ordinal] INT(2)
);
-- Rows: 0

CREATE TABLE [Tb0Vinculo] (
  [ID] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbAcreditacionSistema] (
  [IDAcreditacionSistema] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoAcreditacionSistema] TEXT(255),
  [FechaAcreditacionSistema] SHORT_DATE_TIME(8),
  [RemitenteAcreditacionSistema] TEXT(255),
  [DestinatarioAcreditacionSistema] TEXT(255),
  [FechaEntradaAcreditacionSistema] SHORT_DATE_TIME(8),
  [FechaSalidaAcreditacionSistema] SHORT_DATE_TIME(8),
  [NotasAcreditacionSistema] MEMO,
  [NumeroReferenciaAcreditacionSistema] TEXT(255)
);
-- Rows: 28

CREATE TABLE [TbAuditorias] (
  [IDAuditoria] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [FechaAuditoria] SHORT_DATE_TIME(8),
  [AsuntoAuditoria] TEXT(255),
  [NReferenciaAuditoria] TEXT(255),
  [DescripcionAuditoria] MEMO,
  [NotasAuditoria] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbBuzonTE] (
  [IDBuzonTE] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoBuzonTE] TEXT(255),
  [FechaBuzonTE] SHORT_DATE_TIME(8),
  [RemitenteBuzonTE] TEXT(255),
  [DestinatarioBuzonTE] TEXT(255),
  [NotasBuzonTE] TEXT(255)
);
-- Rows: 12

CREATE TABLE [TbBuzonTEOficio] (
  [IDBuzonTEOficio] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [ReferenciaBuzonTEOficio] TEXT(255),
  [FechaBuzonTEOficio] SHORT_DATE_TIME(8),
  [AsuntoBuzonTEOficio] TEXT(255),
  [DestinatarioBuzonTEOficio] TEXT(255),
  [DescripcionBuzonTEOficio] MEMO,
  [NotasBuzonTEOficio] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCarta] (
  [IDCarta] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [RemitenteCarta] TEXT(255),
  [DestinatarioCarta] TEXT(255),
  [FechaCarta] SHORT_DATE_TIME(8),
  [AsuntoCarta] TEXT(255),
  [NReferenciaCarta] TEXT(255),
  [DescripcionCarta] MEMO,
  [FirmaCarta] TEXT(255),
  [NotasCarta] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCertificados] (
  [IDCertificado] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoCertificado] TEXT(255),
  [ReferenciaCertificado] TEXT(255),
  [FechaCertificado] SHORT_DATE_TIME(8),
  [EmisorAutorCertificado] TEXT(255),
  [DestinatarioCertificado] TEXT(255),
  [NumeroASICertificado] TEXT(255),
  [UbicacionCertificado] TEXT(255),
  [SoporteCertificado] TEXT(255),
  [NotasCertificado] MEMO
);
-- Rows: 0

CREATE TABLE [TbCorrespondencia] (
  [IDCorrespondencia] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [ReferenciaCorrespondencia] TEXT(255),
  [AsuntoCorrespondencia] TEXT(255),
  [TipoCorrespondenciaAcreditacionSistema] TEXT(255),
  [RemitenteCorrespondencia] TEXT(255),
  [DestinatarioCorrespondencia] TEXT(255),
  [NumeroASICorrespondencia] TEXT(255),
  [FechaCorrespondencia] SHORT_DATE_TIME(8),
  [ModoEntradaCorrespondencia] TEXT(255),
  [ModoSalidaCorrespondencia] TEXT(255),
  [UbicacionCorrespondencia] TEXT(255),
  [SoporteCorrespondencia] TEXT(255),
  [NotasCorrespondencia] MEMO
);
-- Rows: 106

CREATE TABLE [TbCripto] (
  [IDCripto] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [FechaCripto] SHORT_DATE_TIME(8),
  [AsuntoCripto] TEXT(255),
  [NReferenciaCripto] TEXT(255),
  [NotasCripto] MEMO
);
-- Rows: 18

CREATE TABLE [TbENSServiciosTelefonica] (
  [IDENS] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [ResponsableProyecto] TEXT(255),
  [FechaENS] SHORT_DATE_TIME(8),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoENS] TEXT(255),
  [NotasENS] MEMO
);
-- Rows: 8

CREATE TABLE [TbFichaContratista] (
  [IDFichaContratista] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoFichaContratista] TEXT(255),
  [FechaFichaContratista] SHORT_DATE_TIME(8),
  [NotasFichaContratista] MEMO
);
-- Rows: 38

CREATE TABLE [TbFormacion] (
  [IDFormacion] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [FechaFormacion] SHORT_DATE_TIME(8),
  [AsuntoFormacion] TEXT(255),
  [NotasFormacion] MEMO
);
-- Rows: 9

CREATE TABLE [TbHSEMHSES] (
  [IDHSEMHSES] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [SubTipoElevacionGradoHSEM] TEXT(255),
  [GradoElevacionGradoHSEM] TEXT(255),
  [RemitenteElevacionGradoHSEM] TEXT(255),
  [DestinatarioElevacionGradoHSEM] TEXT(255),
  [FechaElevacionGradoHSEM] SHORT_DATE_TIME(8),
  [UbicacionElevacionGradoHSEM] TEXT(255),
  [SoporteElevacionGradoHSEM] TEXT(255),
  [NotasElevacionGradoHSEM] MEMO,
  [AsuntoHSEMHSES] TEXT(255)
);
-- Rows: 77

CREATE TABLE [TbInformes] (
  [IDInforme] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoInforme] TEXT(255),
  [ReferenciaInformeSeguimiento] TEXT(255),
  [RemitenteInformeSeguimiento] TEXT(255),
  [DestinatarioInformeSeguimiento] TEXT(255),
  [NumeroASIInformeSeguimiento] TEXT(255),
  [FechaInformeSeguimiento] SHORT_DATE_TIME(8),
  [UbicacionInformeSeguimiento] TEXT(255),
  [SoporteInformeSeguimiento] TEXT(255),
  [NotasInformeSeguimiento] MEMO
);
-- Rows: 0

CREATE TABLE [TbInspeccionesAuditoriasSeguridad] (
  [IDInspeccionesAuditoriasSeguridad] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [FechaInspeccion] SHORT_DATE_TIME(8),
  [AsuntoInspeccionesAuditoriasSeguridad] TEXT(255),
  [NReferenciaInspeccion] TEXT(255),
  [NotasInspecciones] MEMO
);
-- Rows: 0

CREATE TABLE [TbInventarioDocumentacionClasificada] (
  [IDInventarioDocumentacionClasificada] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoInventarioDocumentacionClasificada] TEXT(255),
  [FechaListadoDocumentacionClasificada] SHORT_DATE_TIME(8),
  [RemitenteListadoDocumentacionClasificada] TEXT(255),
  [DestinatarioListadoDocumentacionClasificada] TEXT(255),
  [NotasListadoDocumentacionClasificada] MEMO
);
-- Rows: 8

CREATE TABLE [TbLog] (
  [IDLog] LONG(4),
  [Titulo] TEXT(255),
  [IDRegistro] LONG(4),
  [IDTipoRegistro] LONG(4),
  [Linea] MEMO,
  [Fecha] SHORT_DATE_TIME(8),
  [Usuario] TEXT(255)
);
-- Rows: 2260

CREATE TABLE [TbMatClasificadoReciboEntrega] (
  [IDMatClasificadoReciboEntrega] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoMatClasificadoReciboEntrega] TEXT(255),
  [NumeroMatClasificadoReciboEntrega] TEXT(255),
  [FechaMatClasificadoReciboEntrega] SHORT_DATE_TIME(8),
  [DestinatarioMatClasificadoReciboEntrega] TEXT(255),
  [NotasMatClasificadoReciboEntrega] MEMO,
  [ExpedienteMatClasificadoReciboEntrega] MEMO,
  [Clasificacion] TEXT(255)
);
-- Rows: 101

CREATE TABLE [TbMatClasificadoReciboTransporte] (
  [IDMatClasificadoReciboTransporte] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoMatClasificadoReciboTransporte] TEXT(255),
  [NumeroMatClasificadoReciboTransporte] TEXT(255),
  [FechaMatClasificadoReciboTransporte] SHORT_DATE_TIME(8),
  [DestinatarioMatClasificadoReciboTransporte] TEXT(255),
  [RemitenteMatClasificadoReciboTransporte] TEXT(255),
  [ExpedienteMatClasificadoReciboTransporte] MEMO,
  [NotasMatClasificadoReciboTransporte] MEMO,
  [Clasificacion] TEXT(255)
);
-- Rows: 55

CREATE TABLE [TbMatClasificadoSIntegracion] (
  [IDSMatClasificadoSIntegracion] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoMatClasificadoSIntegracion] TEXT(255),
  [NumeroReciboMatClasificadoSIntegracion] TEXT(255),
  [FechaEnvioMatClasificadoSIntegracion] SHORT_DATE_TIME(8),
  [FechaRecibidoMatClasificadoSIntegracion] SHORT_DATE_TIME(8),
  [RefDelContratoMatClasificadoSIntegracion] TEXT(255),
  [NotasMatClasificadoSIntegracion] MEMO,
  [Clasificacion] TEXT(255)
);
-- Rows: 74

CREATE TABLE [TbNombramientoJS] (
  [IDNombramientoJS] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoNombramientoJS] TEXT(255),
  [FechaNombramientoJS] SHORT_DATE_TIME(8),
  [NumeroReferenciaNombramientoJS] TEXT(255),
  [RemitenteNombramientoJS] TEXT(255),
  [DestinatarioNombramientoJS] TEXT(255),
  [NotasNombramientoJS] MEMO
);
-- Rows: 17

CREATE TABLE [TbNormativas] (
  [IDNormativa] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [FechaNormativa] SHORT_DATE_TIME(8),
  [AsuntoNormativa] TEXT(255),
  [ClasificacionAuditoria] TEXT(255),
  [NRefNormativa] TEXT(255),
  [NotasNormativa] MEMO,
  [SoporteNormativa] TEXT(50)
);
-- Rows: 7

CREATE TABLE [TbPlanesSeguridad] (
  [IDPlanSeguridad] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoPlanSeguridad] TEXT(255),
  [ReferenciaPlanSeguridad] TEXT(255),
  [FechaPlanSeguridad] SHORT_DATE_TIME(8),
  [UbicacionPlanSeguridad] TEXT(255),
  [SoportePlanSeguridad] TEXT(255),
  [NotasPlanSeguridad] MEMO
);
-- Rows: 8

CREATE TABLE [TbPoderesSeguridad] (
  [IDPoderesSeguridad] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoPoderesSeguridad] TEXT(255),
  [ReferenciaPoderesSeguridad] TEXT(255),
  [FechaPoderesSeguridad] SHORT_DATE_TIME(8),
  [UbicacionPoderesSeguridad] TEXT(255),
  [SoportePoderesSeguridad] TEXT(255),
  [NotasPoderesSeguridad] MEMO
);
-- Rows: 4

CREATE TABLE [TbRecibosMP] (
  [IDRecibosMP] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [ExpedienteRecibosMP] TEXT(255),
  [AsuntoRecibosMP] TEXT(255),
  [RemitenteRecibosMP] TEXT(255),
  [DestinatarioRecibosMP] TEXT(255),
  [ClasificacionRecibosMP] TEXT(255),
  [FechaRecibosMP] SHORT_DATE_TIME(8),
  [UbicacionRecibosMP] TEXT(255),
  [SoporteRecibosMP] TEXT(255),
  [NotasRecibosMP] MEMO
);
-- Rows: 0

CREATE TABLE [TbSeguridadEdificio] (
  [IDSeguridadEdificio] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoSeguridadEdificio] TEXT(255),
  [FechaSeguridadEdificio] SHORT_DATE_TIME(8),
  [ReferenciaSeguridadEdificio] TEXT(255),
  [NumeroASISeguridadEdificio] TEXT(255),
  [UbicacionSeguridadEdificio] TEXT(255),
  [SoporteSeguridadEdificio] TEXT(255),
  [NotasSeguridadEdificio] MEMO
);
-- Rows: 52

CREATE TABLE [TbSolDocClas] (
  [IDSolDocClas] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistra] TEXT(255),
  [JuridicaTdE] TEXT(2),
  [JuridicaTME] TEXT(2),
  [JuridicaTSOL] TEXT(2),
  [JuridicaTIS] TEXT(2),
  [AsuntoSolDocClas] TEXT(255),
  [FechaSolicitudSolDocClas] SHORT_DATE_TIME(8),
  [EmpresaSolicitaSolDocClas] TEXT(255),
  [SolicitanteSolDocClas] TEXT(255),
  [NotasSolDocClas] MEMO
);
-- Rows: 5

-- Summary: 36 local tables, 0 linked/skipped
