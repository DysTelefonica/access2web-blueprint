-- Schema dump: AGEDO20_Datos.accdb
-- Tables: 68

CREATE TABLE [Entradas] (
  [Año] TEXT(4),
  [NEntrada] TEXT(4),
  [FEntrada] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [Procedencia] TEXT(50),
  [Clase] TEXT(1),
  [Expediente] TEXT(50),
  [Entrega] TEXT(255),
  [Extracto] TEXT(100),
  [Archivo] TEXT(3),
  [Anexo] TEXT(255),
  [DPD] TEXT(50),
  [Remitente] TEXT(255),
  [ArchivadoPor] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 14125

CREATE TABLE [Salidas] (
  [Año] TEXT(4),
  [NSalida] TEXT(4),
  [FSalida] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [Destino] TEXT(50),
  [Clase] TEXT(1),
  [Extracto] TEXT(100),
  [Anexo] TEXT(255),
  [DPD] TEXT(50),
  [ArchivadoPor] TEXT(50),
  [Expediente] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 11609

CREATE TABLE [TbAuxCodRepetidos] (
  [IDRepetido] TEXT(255)
);
-- Rows: 170

CREATE TABLE [TbCartaDetalle] (
  [IDCarta] LONG(4),
  [CartaDescripcion] TEXT(255),
  [CartaTexto] MEMO,
  [CartaFecha] TEXT(255),
  [CartaATT] MEMO,
  [CartaEMPLAZAMIENTO1] MEMO,
  [CartaEMPLAZAMIENTO2] MEMO,
  [CartaDIRECCION] TEXT(255),
  [CartaPOBLACION] TEXT(255),
  [CartaCP] TEXT(255),
  [CartaNPAPEL] LONG(4),
  [CartaNCDS] LONG(4),
  [CartaURL] MEMO,
  [FEnvioSecretaria] SHORT_DATE_TIME(8)
);
-- Rows: 7

CREATE TABLE [TbCartas] (
  [IDCarta] LONG(4),
  [OrdinalMP14] LONG(4),
  [SGPMC_ASI] TEXT(255),
  [Año] LONG(4),
  [IDProceso] LONG(4)
);
-- Rows: 20

CREATE TABLE [TbCorreoAcciones] (
  [IDCorreo] LONG(4),
  [IDProceso] LONG(4),
  [IDIntegracion] LONG(4)
);
-- Rows: 910

CREATE TABLE [TbCorreosAvisos] (
  [IDCorreoAviso] LONG(4),
  [Destinatarios] MEMO,
  [DestinatariosConCopiaOculta] TEXT(255),
  [Asunto] TEXT(255),
  [Cuerpo] MEMO,
  [FechaEnvio] SHORT_DATE_TIME(8),
  [Observaciones] MEMO,
  [FechaGrabacion] SHORT_DATE_TIME(8)
);
-- Rows: 50

CREATE TABLE [TbDocExpCartaDestinatarios] (
  [IDDestinatarioCartas] LONG(4),
  [NombreDestino] TEXT(255),
  [CartaATT] TEXT(50),
  [CartaEMPLAZAMIENTO1] TEXT(255),
  [CartaEMPLAZAMIENTO2] TEXT(255),
  [CartaDIRECCION] TEXT(50),
  [CartaPOBLACION] TEXT(255),
  [CartaCP] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 24

CREATE TABLE [TbDocExpCodigosReservado] (
  [IDCodigo] LONG(4),
  [Documento] TEXT(255),
  [IDProceso] LONG(4),
  [NRefInterno] TEXT(255)
);
-- Rows: 7

CREATE TABLE [TbDocExpEntregaMatClasificadoDetalle] (
  [IDCartaEntrega] LONG(4),
  [NREGISTROCENTRAL] TEXT(255),
  [NRefInterno] TEXT(255),
  [NumReferencia] TEXT(255),
  [FechaDocumento] SHORT_DATE_TIME(8),
  [Clasificacion] TEXT(255),
  [Asunto] TEXT(255),
  [SOPORTE] TEXT(255),
  [NUMERACION] TEXT(255)
);
-- Rows: 24

CREATE TABLE [TbDocExpEntregaMatClasificadoPpal] (
  [IDCartaEntrega] LONG(4),
  [IDProceso] LONG(4),
  [NSalida] TEXT(255),
  [Fecha] TEXT(255),
  [ClasificacionMaxima] TEXT(255),
  [UsuarioEntrega] TEXT(255),
  [ProgramaDestinatario] TEXT(255),
  [DireccionDestinatarioLinea1] TEXT(255),
  [DireccionDestinatarioLinea2] TEXT(255),
  [CPDestinatario] TEXT(255),
  [CiudadDestinatario] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbDocExpEstadosDefinicion] (
  [Estado] TEXT(50),
  [Definicion] TEXT(255)
);
-- Rows: 16

CREATE TABLE [TbDocExpProceso] (
  [IDProceso] LONG(4),
  [TituloDocPrincipal] TEXT(255),
  [CodExp] TEXT(50),
  [Centro] TEXT(50),
  [TipoDocPrincipal] TEXT(2),
  [AreaDocPrincipal] TEXT(1),
  [VersionableDocPrincipal] TEXT(2),
  [ClasificacionDocPrincipal] TEXT(50),
  [VersionDocPrincipal] TEXT(50),
  [CodDocPrincipal] TEXT(50),
  [URLCarpetaLocal] MEMO,
  [URLCarpeta] MEMO,
  [OBST1] MEMO,
  [NombreProceso] TEXT(50),
  [USUARIOTECNICO] TEXT(50),
  [FREALIZACIONT1] SHORT_DATE_TIME(8),
  [FACTUALIZACIONT1] SHORT_DATE_TIME(8),
  [MODIFICACIONCONTADOR1] INT(2),
  [CartaNPAPEL] INT(2),
  [CartaNCDS] INT(2),
  [USUARIOCALIDAD1] TEXT(50),
  [FVISADO1] SHORT_DATE_TIME(8),
  [OBSVISADO1] MEMO,
  [FRECHAZO1] SHORT_DATE_TIME(8),
  [RECHAZOSCONTADOR1] INT(2),
  [OBSRECHAZO1] MEMO,
  [ACCIONCORRECTIVA1] MEMO,
  [URLDocRechazo1] MEMO,
  [CodDocAsociado1] TEXT(50),
  [CodDocAsociado2] TEXT(50),
  [CodDocAsociado3] TEXT(50),
  [TituloDocAsociado1] TEXT(255),
  [TituloDocAsociado2] TEXT(255),
  [TituloDocAsociado3] TEXT(255),
  [TipoDocAsociado1] TEXT(2),
  [TipoDocAsociado2] TEXT(2),
  [TipoDocAsociado3] TEXT(2),
  [ModeloDocAsociado1] INT(2),
  [ModeloDocAsociado2] INT(2),
  [ModeloDocAsociado3] INT(2),
  [AreaDocAsociado1] TEXT(1),
  [AreaDocAsociado2] TEXT(1),
  [AreaDocAsociado3] TEXT(1),
  [URLPlantillaDocAsociado1] MEMO,
  [URLPlantillaDocAsociado2] MEMO,
  [URLPlantillaDocAsociado3] MEMO,
  [ComentariosDocAsociado1] MEMO,
  [ComentariosDocAsociado2] MEMO,
  [ComentariosDocAsociado3] MEMO,
  [DocAsociado1ParaTecnico] TEXT(2),
  [DocAsociado2ParaTecnico] TEXT(2),
  [DocAsociado3ParaTecnico] TEXT(2),
  [URLDocPrincipal] MEMO,
  [FApruebaDocPrincipal] SHORT_DATE_TIME(8),
  [FRevisaDocPrincipal] SHORT_DATE_TIME(8),
  [RevisaDocPrincipal] TEXT(255),
  [RevisaDocPrincipalCargo] TEXT(255),
  [ApruebaDocPrincipal] TEXT(255),
  [ApruebaDocPrincipalCargo] TEXT(255),
  [OBST2] MEMO,
  [FREALIZACIONT2] SHORT_DATE_TIME(8),
  [FACTUALIZACIONT2] SHORT_DATE_TIME(8),
  [MODIFICACIONCONTADOR2] INT(2),
  [USUARIOMONTAJET] TEXT(255),
  [USUARIOCALIDAD2] TEXT(50),
  [URLDocPrincipalFirmado] MEMO,
  [URLDocAsociado1] MEMO,
  [URLDocAsociado2] MEMO,
  [URLDocAsociado3] MEMO,
  [FVISADO2] SHORT_DATE_TIME(8),
  [OBSVISADO2] MEMO,
  [FRECHAZO2] SHORT_DATE_TIME(8),
  [RECHAZOSCONTADOR2] INT(2),
  [OBSRECHAZO2] MEMO,
  [ACCIONCORRECTIVA2] MEMO,
  [CartaFecha] SHORT_DATE_TIME(8),
  [CartaATT] MEMO,
  [CartaEMPLAZAMIENTO1] MEMO,
  [CartaEMPLAZAMIENTO2] MEMO,
  [CartaDIRECCION] TEXT(255),
  [CartaPOBLACION] TEXT(255),
  [CartaCP] TEXT(255),
  [CartaPARACORREO] TEXT(2),
  [CartaURL] MEMO,
  [CartaUsuarioFirma] TEXT(255),
  [FREALIZACIONT3] SHORT_DATE_TIME(8),
  [FACTUALIZACIONT3] SHORT_DATE_TIME(8),
  [MODIFICACIONCONTADOR3] INT(2),
  [CorreoCD] TEXT(255),
  [CadenaArchivosCD] MEMO,
  [FechaEnvioCorreoCD] SHORT_DATE_TIME(8),
  [FechaRealizacionCD] SHORT_DATE_TIME(8),
  [UsuarioRealizaCD] TEXT(255),
  [USUARIOSALIDAT] TEXT(255),
  [USUARIOCALIDAD3] TEXT(50),
  [FVISADO3] SHORT_DATE_TIME(8),
  [OBSVISADO3] MEMO,
  [FRECHAZO3] SHORT_DATE_TIME(8),
  [RECHAZOSCONTADOR3] INT(2),
  [OBSRECHAZO3] MEMO,
  [ACCIONCORRECTIVA3] MEMO,
  [FEnvioSecretaria] SHORT_DATE_TIME(8),
  [URLCarpetaAGEDODocPrincipal] MEMO,
  [URLCarpetaAGEDODocAsociado1] MEMO,
  [URLCarpetaAGEDODocAsociado2] MEMO,
  [URLCarpetaAGEDODocAsociado3] MEMO,
  [RegistroSalidaDocPrincipal] TEXT(50),
  [IDDocumentoDocPrincipal] LONG(4),
  [IDDocumentoLDocAsociado1] LONG(4),
  [IDDocumentoLDocAsociado2] LONG(4),
  [IDDocumentoLDocAsociado3] LONG(4),
  [FREVISION] SHORT_DATE_TIME(8),
  [FFINPROCESO] SHORT_DATE_TIME(8),
  [Estado] TEXT(50),
  [CARTAEnvios] MEMO,
  [ObservacionesGenerales] MEMO,
  [FECHABORRADOALMACEN] SHORT_DATE_TIME(8)
);
-- Rows: 372

CREATE TABLE [TbDocExpProcesoArchivosParaAGEDO] (
  [IDProceso] LONG(4),
  [ArchivosParaDocPrincipal] MEMO,
  [ArchivosParaDocAsociado1] MEMO,
  [ArchivosParaDocAsociado2] MEMO,
  [ArchivosParaDocAsociado3] MEMO
);
-- Rows: 120

CREATE TABLE [TbDocExpProcesoAutorizados] (
  [IDProceso] LONG(4),
  [Usuario] TEXT(255)
);
-- Rows: 947

CREATE TABLE [TbDocExpProcesoID] (
  [IDProceso] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 438

CREATE TABLE [TbDocExpProcesoRegistroTArchivos] (
  [IDProceso] LONG(4),
  [URLFinal] MEMO,
  [URLLocal] MEMO
);
-- Rows: 810

CREATE TABLE [TbDocumentos] (
  [IDDocumento] LONG(4),
  [FechaAlta] SHORT_DATE_TIME(8),
  [CodExp] TEXT(50),
  [Codigo] TEXT(50),
  [Titulo] TEXT(255),
  [Tipo] TEXT(50),
  [Area] TEXT(50),
  [Edicion] LONG(4),
  [FEdicion] SHORT_DATE_TIME(8),
  [FENTRADAENVIGOR] SHORT_DATE_TIME(8),
  [Archivo] TEXT(50),
  [ResponsableEdicion] TEXT(255),
  [URLCarpetaArchivosLocales] MEMO,
  [CadenaNombreArchivos] MEMO,
  [URLCarpetaAGEDO] MEMO,
  [Centro] TEXT(50),
  [Clasificacion] TEXT(50),
  [Registrador] TEXT(50),
  [RegistroEntrada] TEXT(50),
  [RegistroSalida] TEXT(50),
  [Versionable] TEXT(2),
  [UltimaVersion] TEXT(2),
  [Estado] TEXT(50),
  [FechaCaducidad] SHORT_DATE_TIME(8),
  [CLASE] TEXT(50),
  [Observaciones] MEMO,
  [FechaDeComienzoNuevaVersion] SHORT_DATE_TIME(8),
  [IDDocumentoNuevaVersion] LONG(4),
  [IDDocumentoRevision] LONG(4),
  [NCPARAINDICADOR] TEXT(255),
  [NCABIERTACERRADA] TEXT(255),
  [AdjuntoEncriptado] TEXT(2)
);
-- Rows: 6510

CREATE TABLE [TbDocumentosAccesibilidad] (
  [IDDocumento] LONG(4),
  [URLAGEDO] MEMO,
  [CadenaNombreArchivos] MEMO,
  [Accesible] TEXT(255)
);
-- Rows: 4657

CREATE TABLE [TbDocumentosAlarmas] (
  [IDAlarma] LONG(4),
  [Destinatarios] MEMO,
  [IDDocumento] LONG(4),
  [FechaCaducidad] SHORT_DATE_TIME(8),
  [IDDocumentoRevision] LONG(4),
  [IDDocumentoNuevaVersion] LONG(4),
  [MarcadoComoUltimaVersion] TEXT(2),
  [ObservacionCierre] MEMO,
  [FechaCumplimentada] SHORT_DATE_TIME(8)
);
-- Rows: 122

CREATE TABLE [TbDocumentosArea] (
  [Area] TEXT(1),
  [Descripcion] TEXT(255)
);
-- Rows: 16

CREATE TABLE [TbDocumentosCaducidades] (
  [IDDocumento] LONG(4),
  [DiasParaCaducar] LONG(4)
);
-- Rows: 53

CREATE TABLE [TbDocumentosClasificacion] (
  [Clasificacion] TEXT(50)
);
-- Rows: 4

CREATE TABLE [TbDocumentosID] (
  [IDDocumento] LONG(4),
  [UsuarioCrea] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaEliminacion] SHORT_DATE_TIME(8),
  [UsuarioElimina] TEXT(255)
);
-- Rows: 3280

CREATE TABLE [TbDocumentosTipo] (
  [Tipo] TEXT(2),
  [Modelo] INT(2),
  [Descripcion] TEXT(255),
  [DocumentoTecnico] TEXT(50),
  [Versionable] TEXT(2),
  [EsFormato] TEXT(50),
  [NombrePlantilla] MEMO,
  [TituloDoc] TEXT(255),
  [URLPlantillaNoEnProcuccion] MEMO,
  [AplicaANormativa] TEXT(2),
  [AplicaADocExpedientes] TEXT(2),
  [AdmiteComentarios] TEXT(2),
  [AplicaNC] TEXT(2),
  [FechaObsoleto] SHORT_DATE_TIME(8)
);
-- Rows: 88

CREATE TABLE [TbEncuestaInternaEnunciados] (
  [TipoEncuesta] TEXT(255),
  [Apartado] TEXT(255),
  [Enunciado] MEMO
);
-- Rows: 110

CREATE TABLE [TbEncuestasInternasAplicaciones] (
  [Aplicacion] TEXT(255),
  [Encuesta] TEXT(3)
);
-- Rows: 3

CREATE TABLE [TbEncuestasInternasControl] (
  [IDControl] LONG(4),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [Completo] TEXT(2),
  [Enviadas] INT(2),
  [Recibidas] INT(2),
  [FechaLoteCerrado] SHORT_DATE_TIME(8)
);
-- Rows: 30

CREATE TABLE [TbEncuestasInternasDetalle] (
  [ID] TEXT(255),
  [Apartado] TEXT(255),
  [Resultado] MEMO,
  [Enunciado] MEMO
);
-- Rows: 3515

CREATE TABLE [TbEncuestasInternasEnviosRecepciones] (
  [IDControl] LONG(4),
  [Destinatario] TEXT(50),
  [Encuesta] TEXT(3),
  [Enviador] TEXT(255),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [IDEncuesta] TEXT(255)
);
-- Rows: 255

CREATE TABLE [TbEncuestasInternasGrupos] (
  [Grupo] TEXT(255)
);
-- Rows: 3

CREATE TABLE [TbEncuestasInternasGrupoUsuarios] (
  [Grupo] TEXT(255),
  [Usuario] TEXT(255)
);
-- Rows: 20

CREATE TABLE [TbEncuestasInternasPpal] (
  [ID] TEXT(255),
  [Encuesta] TEXT(3),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [IDControl] LONG(4),
  [Usuario] TEXT(255)
);
-- Rows: 96

CREATE TABLE [TbEncuestasSatisfaccion] (
  [IDEncuesta] LONG(4),
  [IDExpediente] LONG(4),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [UltimaEncuesta] TEXT(2),
  [NombreContacto] TEXT(255),
  [emailContacto] TEXT(255),
  [Asunto] TEXT(255),
  [Mensaje] MEMO,
  [IDDocumento] LONG(4),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Satisfaccion] TEXT(255),
  [Cumplimiento] TEXT(255),
  [Rapidez] TEXT(255),
  [Atencion] TEXT(255),
  [Formacion] TEXT(255),
  [Documento] TEXT(255),
  [Impresion] TEXT(255),
  [Sugerencias] MEMO,
  [Observaciones] MEMO
);
-- Rows: 52

CREATE TABLE [TbEncuestasSatisfaccionEnvio] (
  [IDEncuesta] LONG(4),
  [IDExpediente] LONG(4),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [UltimaEncuesta] TEXT(2),
  [Observaciones] MEMO,
  [IDDocumento] LONG(4)
);
-- Rows: 49

CREATE TABLE [TbEncuestasSatisfaccionRecepcion] (
  [IDEncuesta] LONG(4),
  [FechaRecepcion] SHORT_DATE_TIME(8),
  [Satisfaccion] DOUBLE(8),
  [Cumplimiento] DOUBLE(8),
  [Rapidez] DOUBLE(8),
  [Atencion] DOUBLE(8),
  [Formacion] DOUBLE(8),
  [Documento] DOUBLE(8),
  [Impresion] DOUBLE(8),
  [NombreCompleto] TEXT(255),
  [CargoRango] TEXT(255),
  [Direccion] TEXT(255),
  [Telefono] TEXT(255),
  [email] TEXT(255),
  [EmpresaOrganismo] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 35

CREATE TABLE [TbExpedientesActuales] (
  [CodExp] TEXT(255)
);
-- Rows: 135

CREATE TABLE [TbExpedientesContactos] (
  [IDExp] LONG(4),
  [Contacto] TEXT(255),
  [NombreCompleto] TEXT(255),
  [CargoRango] TEXT(255),
  [Direccion] TEXT(255),
  [Telefono] TEXT(255),
  [email] TEXT(255),
  [EmpresaOrganismo] TEXT(255)
);
-- Rows: 2

CREATE TABLE [TbFicherosNoAlcanzables] (
  [IDDocumento] LONG(4),
  [URLCarpetaAGEDOOriginal] MEMO
);
-- Rows: 28

CREATE TABLE [TbFormatos] (
  [IDFormato] LONG(4),
  [TIPOCOMPLETO] TEXT(255),
  [CODIGOFORMATO] TEXT(255),
  [AREAFORMATO] TEXT(1),
  [EQUIPO] TEXT(255),
  [CodExp] TEXT(255),
  [CENTRO] TEXT(255),
  [NDPD] TEXT(255),
  [FINICIO] TEXT(255),
  [FFIN] TEXT(255),
  [LUGAR] TEXT(255),
  [NOMBRERESPONSABLE] TEXT(255),
  [CARGORESPONSABLE] TEXT(255),
  [SUMINISTRADOR] TEXT(255),
  [CLIENTE] TEXT(255),
  [Aplicacion] TEXT(55),
  [Reparos] TEXT(2),
  [Notas] MEMO,
  [FechaRegistro] SHORT_DATE_TIME(8),
  [UsuarioRegistro] TEXT(255),
  [URLPlantilla] MEMO,
  [PuedeTenerObservacionesCalidad] TEXT(2),
  [NotasCalidad] MEMO,
  [IDDocumento] LONG(4),
  [TITULO] TEXT(255),
  [PA] TEXT(2),
  [FASE] TEXT(50),
  [IDDOCREVISADO] LONG(4)
);
-- Rows: 424

CREATE TABLE [tblCentros] (
  [codigo] TEXT(50),
  [CODLETRA] TEXT(50),
  [Centro] TEXT(255),
  [Provincia] TEXT(255),
  [Oficial] BOOLEAN(1),
  [TIPO DE CENTRO] TEXT(255),
  [EJERCITO] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 488

CREATE TABLE [TbLog] (
  [ID] LONG(4),
  [Titulo] TEXT(255),
  [Accion] MEMO,
  [IDProceso] LONG(4),
  [IDDocumento] LONG(4),
  [IDExpediente] LONG(4),
  [FechaAccion] SHORT_DATE_TIME(8)
);
-- Rows: 5660

CREATE TABLE [TbMatClasificadoEntradaDetalle] (
  [IDEntradaDetalle] LONG(4),
  [IDEntrada] LONG(4),
  [NumeroRegistroCentral] TEXT(255),
  [NEntrada] TEXT(255),
  [Descripcion] TEXT(255),
  [SoporteFisico] TEXT(255),
  [IDDocumento] LONG(4)
);
-- Rows: 2

CREATE TABLE [TbMatClasificadoEntradaDetalleAux] (
  [IDEntradaDetalle] LONG(4),
  [IDEntrada] LONG(4),
  [NumeroRegistroCentral] TEXT(255),
  [UsuarioConectado] TEXT(255),
  [NumeroRegistroInterno] TEXT(255),
  [Descripcion] TEXT(255),
  [SoporteFisico] TEXT(255),
  [IDDocumento] LONG(4)
);
-- Rows: 0

CREATE TABLE [TbMatClasificadoEntradaPrincipal] (
  [IDEntrada] LONG(4),
  [DescripcionEntrada] TEXT(255),
  [FechaRecibo] SHORT_DATE_TIME(8),
  [Programa] TEXT(255),
  [ReciboCompulsadoNombre] TEXT(255),
  [Remitente] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbMatClasificadoEntregaDetalle] (
  [IDEntregaDetalle] LONG(4),
  [IDEntrega] LONG(4),
  [NumeroRegistroCentral] TEXT(255),
  [NumeroRegistroInterno] TEXT(255),
  [NumeroReferencia] TEXT(255),
  [FechaDocumento] SHORT_DATE_TIME(8),
  [Clasificacion] TEXT(255),
  [Descripcion] TEXT(255),
  [SoporteFisico] TEXT(255),
  [Numeracion] TEXT(255)
);
-- Rows: 47

CREATE TABLE [TbMatClasificadoEntregaDetalleAux] (
  [IDEntregaDetalle] LONG(4),
  [IDEntrega] LONG(4),
  [NumeroRegistroCentral] TEXT(255),
  [UsuarioConectado] TEXT(255),
  [NumeroRegistroInterno] TEXT(255),
  [NumeroReferencia] TEXT(255),
  [FechaDocumento] SHORT_DATE_TIME(8),
  [Clasificacion] TEXT(255),
  [Descripcion] TEXT(255),
  [SoporteFisico] TEXT(255),
  [Numeracion] TEXT(255)
);
-- Rows: 22

CREATE TABLE [TbMatClasificadoEntregaPrincipal] (
  [IDEntrega] LONG(4),
  [DescripcionEntrega] TEXT(255),
  [IDIntegracion] LONG(4),
  [CartaATT] TEXT(255),
  [CartaAsunto] TEXT(255),
  [CartaNRef] TEXT(255),
  [CartaTexto] MEMO,
  [CartaDescripcion] TEXT(255),
  [CartaFecha] SHORT_DATE_TIME(8),
  [CartaNPAPEL] LONG(4),
  [CartaUsuarioJefeSeguridad] TEXT(255),
  [CartaNCDS] LONG(4),
  [MatNSalida] TEXT(255),
  [MatFecha] SHORT_DATE_TIME(8),
  [MatClasificacionMaxima] TEXT(255),
  [MatPrograma] TEXT(255),
  [MatExpediente] MEMO,
  [MatProgramaDirLinea1] TEXT(255),
  [MatProgramaDirLinea2] TEXT(255),
  [MatUsuarioEntrega] TEXT(255),
  [MatUsuarioCargoEntrega] TEXT(255),
  [CartaURL] MEMO,
  [FEnvioSecretaria] SHORT_DATE_TIME(8),
  [ReciboCompulsadoNombre] TEXT(255)
);
-- Rows: 24

CREATE TABLE [TbMatClasificadoIntegracionDetalle] (
  [IDIntegracionDetalle] LONG(4),
  [IDIntegracion] LONG(4),
  [NumeroReferencia] TEXT(255),
  [SoporteFisico] TEXT(255),
  [Numeracion] TEXT(255),
  [NumeroRegistroInterno] TEXT(255),
  [FechaDocumento] SHORT_DATE_TIME(8),
  [Clasificacion] TEXT(255),
  [Descripcion] TEXT(255),
  [NumeroRegistroCentral] TEXT(255)
);
-- Rows: 59

CREATE TABLE [TbMatClasificadoIntegracionDetalleAux] (
  [IDIntegracionDetalle] LONG(4),
  [IDIntegracion] LONG(4),
  [NumeroReferencia] TEXT(255),
  [SoporteFisico] TEXT(255),
  [Numeracion] TEXT(255),
  [UsuarioConectado] TEXT(255),
  [NumeroRegistroInterno] TEXT(255),
  [FechaDocumento] SHORT_DATE_TIME(8),
  [Clasificacion] TEXT(255),
  [Descripcion] TEXT(255),
  [NumeroRegistroCentral] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbMatClasificadoIntegracionPrincipal] (
  [IDIntegracion] LONG(4),
  [DescripcionSolicitud] TEXT(255),
  [EstadoMatClasificado] TEXT(3),
  [ClasificacionMaxima] TEXT(255),
  [Expediente] TEXT(255),
  [TITULOEXP] TEXT(255),
  [Programa] TEXT(255),
  [FechaSolicitud] SHORT_DATE_TIME(8),
  [Resuelta] TEXT(2),
  [IDProceso] LONG(4),
  [ReciboCompulsadoNombre] TEXT(255)
);
-- Rows: 29

CREATE TABLE [TbMatClasificadoTransporteDetalle] (
  [IDTransporte] LONG(4),
  [TranspNSalida] TEXT(255),
  [MatClasificacionMaximaPorEntrega] TEXT(50),
  [CartaFecha] SHORT_DATE_TIME(8)
);
-- Rows: 16

CREATE TABLE [TbMatClasificadoTransporteDetalleAux] (
  [IDTransporte] LONG(4),
  [TranspNSalida] TEXT(255),
  [UsuarioConectado] TEXT(255),
  [MatClasificacionMaximaPorEntrega] TEXT(50),
  [CartaFecha] SHORT_DATE_TIME(8)
);
-- Rows: 16

CREATE TABLE [TbMatClasificadoTransporteEntrega] (
  [IDTransporte] LONG(4),
  [IDEntrega] LONG(4)
);
-- Rows: 0

CREATE TABLE [TbMatClasificadoTransportePrincipal] (
  [IDTransporte] LONG(4),
  [TranspNumeroReferencia] TEXT(255),
  [TranspFecha] SHORT_DATE_TIME(8),
  [TranspClasificacionMaxima] TEXT(255),
  [TranspContenido] MEMO,
  [TranspJefeSeguridad] TEXT(255),
  [TranspReciboCompulsadoNombre] TEXT(50)
);
-- Rows: 13

CREATE TABLE [TbNC] (
  [CodigoNC] TEXT(255),
  [AREA] TEXT(1),
  [TIPONC] TEXT(2),
  [IDDocumentoOrigen] LONG(4),
  [TITULO] TEXT(255),
  [FAPERTURA] SHORT_DATE_TIME(8),
  [ESNOCONFORMIDAD] TEXT(2),
  [DESCRIPCION] MEMO,
  [ANALISISCAUSAS] MEMO,
  [FCIERRE] SHORT_DATE_TIME(8),
  [RESPONSABLEAPERTURA] TEXT(255),
  [FFIRMACLIENTE] SHORT_DATE_TIME(8),
  [FFIRMASUMINSTRADOR] SHORT_DATE_TIME(8),
  [FFIRMAAUDITOR] SHORT_DATE_TIME(8),
  [OBSERVACIONES] MEMO,
  [NCPARAINDICADOR] TEXT(2)
);
-- Rows: 24

CREATE TABLE [TbNCAccionesPropuestas] (
  [CODIGOAP] TEXT(255),
  [CodigoNC] TEXT(255),
  [ACCIONPROPUESTA] MEMO,
  [RESPONSABLEAP] TEXT(255),
  [FAP] SHORT_DATE_TIME(8),
  [FIMPLANTACIONAP] SHORT_DATE_TIME(8),
  [FCIERREAP] SHORT_DATE_TIME(8)
);
-- Rows: 32

CREATE TABLE [TbNCAccionesRealizadas] (
  [CODIGOAR] TEXT(255),
  [CodigoNC] TEXT(255),
  [ACCIONREALIZADA] MEMO,
  [RESPONSABLEAR] TEXT(255),
  [FAR] SHORT_DATE_TIME(8),
  [FIMPLANTACIONAR] SHORT_DATE_TIME(8),
  [FCIERREAR] SHORT_DATE_TIME(8)
);
-- Rows: 70

CREATE TABLE [TbNCAnexos] (
  [IDAnexo] TEXT(255),
  [CodigoNC] TEXT(255),
  [DescripcionAnexo] TEXT(255),
  [FormatoNC] TEXT(2),
  [URLAnexo] MEMO,
  [FechaAnexo] SHORT_DATE_TIME(8)
);
-- Rows: 23

CREATE TABLE [TbParaAvisos] (
  [ID] LONG(4),
  [IDProceso] LONG(4),
  [IDDocumento] LONG(4),
  [IDAlarmaRevision] LONG(4),
  [CodExp] TEXT(50),
  [Dias] LONG(4),
  [Usuario] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbPlanificaciones] (
  [IDPlanificacion] LONG(4),
  [IDDocumento] LONG(4),
  [FechaPrevista] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [MotivoReprogramacion] MEMO,
  [DiasDeAvisoPrevio] LONG(4)
);
-- Rows: 0

CREATE TABLE [TbRevisiones] (
  [IDRevision] LONG(4),
  [Descripcion] TEXT(255),
  [FechaRevision] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 7

CREATE TABLE [TbRevisionesAlarmas] (
  [IDAlarmaRevision] LONG(4),
  [IDRevision] LONG(4),
  [FAlarma] SHORT_DATE_TIME(8),
  [IDRevisionNueva] LONG(4),
  [FechaCierreAlarma] SHORT_DATE_TIME(8),
  [AsuntoCorreo] TEXT(255),
  [CuerpoCorreo] MEMO,
  [DestinatariosCorreo] TEXT(255),
  [FechaEnvioCorreo] SHORT_DATE_TIME(8)
);
-- Rows: 7

CREATE TABLE [TbSugerenciasPrincipal] (
  [IDSugerencia] LONG(4),
  [CodigoSugerencia] TEXT(255),
  [Descripcion] MEMO,
  [UsuarioEmisor] TEXT(255),
  [UsuarioReceptor] TEXT(255),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8)
);
-- Rows: 49

CREATE TABLE [TbSugerenciasRespuestas] (
  [IDSugerencia] LONG(4),
  [IDRespuesta] LONG(4),
  [CodigoRespuesta] TEXT(255),
  [Respuesta] MEMO,
  [UsuarioRespuesta] TEXT(255),
  [FechaRespuesta] SHORT_DATE_TIME(8)
);
-- Rows: 54

-- LINKED/SKIPPED: [tbUsuarios] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbVersionCliente] (
  [VersionCliente] TEXT(50)
);
-- Rows: 1

-- Summary: 66 local tables, 2 linked/skipped
