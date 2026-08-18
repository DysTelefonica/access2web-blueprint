-- Schema dump: Gestion_Brass_Gestion_Datos.accdb
-- Tables: 57

CREATE TABLE [Tb0FiltroGestion] (
  [ID] TEXT(255),
  [Fecha] TEXT(255),
  [Datos] MEMO,
  [Usuario] TEXT(255),
  [Visto] TEXT(2)
);
-- Rows: 6551

CREATE TABLE [TbActividades] (
  [IDActividad] TEXT(50),
  [IDEvento] TEXT(50),
  [ALIASTECNICO] TEXT(255),
  [Actividad] TEXT(50),
  [FechaAlta] SHORT_DATE_TIME(8),
  [HorasLaborables] DOUBLE(8),
  [HorasExtras] DOUBLE(8),
  [Ubicacion] MEMO,
  [DESCRIPCION] MEMO,
  [IDEVENTOGENERADO] TEXT(50),
  [IDExportacion] TEXT(50),
  [CodImportacion] TEXT(50),
  [IDFacturacion] LONG(4),
  [Observaciones] MEMO
);
-- Rows: 25639

CREATE TABLE [TbAnexos] (
  [IDAnexo] LONG(4),
  [IDEvento] TEXT(50),
  [IDActividad] TEXT(50),
  [IDMaterial] TEXT(50),
  [IDMaterialSeguimiento] TEXT(50),
  [IDSubContratacion] LONG(4),
  [IDGasto] LONG(4),
  [Titulo] TEXT(255),
  [Descripcion] MEMO,
  [NombreArchivo] TEXT(255)
);
-- Rows: 7583

CREATE TABLE [TbAuxActividad] (
  [NODO] TEXT(255),
  [BUI] TEXT(255),
  [SUBSISTEMA] TEXT(255),
  [IDEQUIPO] TEXT(255),
  [IDEVENTO] TEXT(255),
  [TIPOTECNICO] TEXT(255),
  [ALIASTECNICO] TEXT(255),
  [TIPOACTIVIDAD] TEXT(255),
  [FECHAACTIVIDAD] TEXT(255),
  [HORASLABORABLES] TEXT(50),
  [HORASEXTRAS] TEXT(50),
  [UBICACION] TEXT(255),
  [DESCRIPCION] MEMO,
  [CODREGISTRODIFERENCIA] MEMO
);
-- Rows: 0

CREATE TABLE [TbAuxEventos] (
  [NODO] TEXT(255),
  [BUI] TEXT(255),
  [SUBSISTEMA] TEXT(255),
  [EQUIPO] TEXT(255),
  [IDEVENTO] TEXT(255),
  [PMPR] TEXT(255),
  [TIPOEVENTO] TEXT(255),
  [CRITICIDAD] TEXT(255),
  [TIPOTECNICO] TEXT(255),
  [ORIGINADOR] TEXT(255),
  [FECHAALTA] SHORT_DATE_TIME(8),
  [HORAALTA] SHORT_DATE_TIME(8),
  [TIEMPORESPUESTA] SHORT_DATE_TIME(8),
  [FECHAFIN] SHORT_DATE_TIME(8),
  [HORAFIN] SHORT_DATE_TIME(8),
  [CAUSAFIN] MEMO,
  [CONTACTO] TEXT(255),
  [IDEVENTOGENERADO] MEMO,
  [DESCRIPCION] MEMO,
  [CODREGISTRODIFERENCIA] MEMO
);
-- Rows: 0

CREATE TABLE [TbAuxManteniminetosPreventivosCalendario] (
  [BUI] TEXT(50),
  [IDEQUIPO] LONG(4),
  [NOMBREEQUIPO] TEXT(255),
  [MESPROG] TEXT(50),
  [SEMANAPROG] TEXT(50),
  [ANIO] TEXT(50),
  [FECHASREPARACIONENANIO] MEMO,
  [IDEVENTO] MEMO,
  [PARTESMTOREALIZADOENANIO] MEMO
);
-- Rows: 93

CREATE TABLE [TbAuxMateriales] (
  [NODO] TEXT(255),
  [BUI] TEXT(255),
  [SUBSISTEMA] TEXT(255),
  [EQUIPO] TEXT(255),
  [IDEVENTO] TEXT(255),
  [TIPOACCION] TEXT(255),
  [PN] TEXT(255),
  [NS] TEXT(255),
  [MATERIAL] TEXT(255),
  [FECHAENTREGA] TEXT(50),
  [COSTE] TEXT(50),
  [PrecioSinIVA] TEXT(50),
  [REPARADOPOR] TEXT(255),
  [CODREGISTRODIFERENCIA] MEMO
);
-- Rows: 0

CREATE TABLE [TbAuxPlanificacion] (
  [IDPlanificacion] LONG(4),
  [Equipo] TEXT(255),
  [FechaPrevista] SHORT_DATE_TIME(8),
  [DiasRestan] LONG(4),
  [BUI] TEXT(255),
  [SUBSISTEMA] TEXT(255),
  [CadenaIDEventosPosibles] MEMO,
  [CadenaIDEventosUsados] MEMO
);
-- Rows: 81

CREATE TABLE [TbBUI] (
  [BUI] TEXT(255)
);
-- Rows: 18

CREATE TABLE [TbBuiIDEvento] (
  [BUI] TEXT(50),
  [CODBUIIDVENTO] TEXT(50)
);
-- Rows: 11

CREATE TABLE [TbCausaFin] (
  [CAUSAFIN] TEXT(50)
);
-- Rows: 4

CREATE TABLE [TbCodActividad] (
  [CODIGO] TEXT(50),
  [ACTIVIDADNOPROGRAMADA] TEXT(255)
);
-- Rows: 12

CREATE TABLE [TbCriticidad] (
  [CRITICIDAD] INT(2),
  [DESCRIPCION] MEMO
);
-- Rows: 3

CREATE TABLE [TbEquipos] (
  [IDEquipo] LONG(4),
  [NODO] TEXT(50),
  [BUI] TEXT(50),
  [SUBSISTEMA] TEXT(50),
  [Equipo] TEXT(255),
  [Descripcion] MEMO,
  [FechaObsoleto] SHORT_DATE_TIME(8)
);
-- Rows: 806

CREATE TABLE [TbEquiposCalibrables] (
  [IDEquipoCalibrable] LONG(4),
  [Equipo] TEXT(255),
  [Modelo] TEXT(255),
  [NS] TEXT(50),
  [BUI] TEXT(50),
  [Observaciones] MEMO,
  [FechaBajaParaCalibracion] SHORT_DATE_TIME(8)
);
-- Rows: 27

CREATE TABLE [TbEquiposCalibrablesFechas] (
  [IDCalibracion] LONG(4),
  [IDEquipoCalibrable] LONG(4),
  [FechaCalibrado] SHORT_DATE_TIME(8),
  [FechaFinCalibrado] SHORT_DATE_TIME(8),
  [URLDocumentosAnexados] TEXT(255),
  [FechaComunicacionCliente] SHORT_DATE_TIME(8)
);
-- Rows: 94

CREATE TABLE [TbEquiposMedida] (
  [IDEquipoMedida] LONG(4),
  [Nombre] TEXT(255),
  [NS] TEXT(255),
  [PN] TEXT(255),
  [FechaInicioServicio] SHORT_DATE_TIME(8),
  [FechaFinServicio] SHORT_DATE_TIME(8),
  [Marca] TEXT(255),
  [Modelo] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 14

CREATE TABLE [TbEquiposMedidaCalibraciones] (
  [IDCalibracion] LONG(4),
  [IDEquipoMedida] LONG(4),
  [FechaCalibracion] SHORT_DATE_TIME(8),
  [FechaFinCalibracion] SHORT_DATE_TIME(8),
  [EmpresaCalibradora] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 28

CREATE TABLE [TbEventos] (
  [IDEvento] TEXT(50),
  [NODO] TEXT(50),
  [BUI] TEXT(50),
  [SUBSISTEMA] TEXT(255),
  [IDEquipo] LONG(4),
  [PMPR] TEXT(50),
  [TIPOEVENTO] TEXT(50),
  [CRITICIDAD] INT(2),
  [ALIASTECNICO] TEXT(50),
  [ORIGINADOR] TEXT(50),
  [FECHAALTAEVENTO] SHORT_DATE_TIME(8),
  [HORAINICIALEVENTO] SHORT_DATE_TIME(8),
  [FechaFinal] SHORT_DATE_TIME(8),
  [HORAFINALEVENTO] SHORT_DATE_TIME(8),
  [TIEMPORESPUESTAEVENTO] SHORT_DATE_TIME(8),
  [CAUSAFIN] MEMO,
  [DESCRIPCION] MEMO,
  [CONTACTO] TEXT(255),
  [FechaRegistroAlta] SHORT_DATE_TIME(8),
  [FechaRegistroModificacion] SHORT_DATE_TIME(8),
  [Franqueado] BOOLEAN(1),
  [IDExportacion] TEXT(50),
  [CodImportacion] TEXT(50),
  [IDParte] TEXT(255),
  [Notas] MEMO,
  [FechaEnInformeRAC] SHORT_DATE_TIME(8),
  [FechaRecepcionNotificacion] SHORT_DATE_TIME(8),
  [FechaInicioContactoCliente] SHORT_DATE_TIME(8),
  [IncidenciaAveriaOReparacion] BOOLEAN(1),
  [FechaInicioTiempoAdquisicion] SHORT_DATE_TIME(8),
  [FechaFinTiempoAdquisicion] SHORT_DATE_TIME(8),
  [TipoReparacion] TEXT(50),
  [Urgente] BOOLEAN(1),
  [EventoConServicioAfectado] BOOLEAN(1),
  [FechaRestablecimientoServicio] SHORT_DATE_TIME(8),
  [TipoRepInsitu] BOOLEAN(1),
  [TipoRepNoSMT] BOOLEAN(1),
  [TipoRepValvulas] BOOLEAN(1)
);
-- Rows: 5773

CREATE TABLE [TbEventosEquipoMedida] (
  [IDEventoEquipoMedida] LONG(4),
  [IDEvento] TEXT(255),
  [IDEquipoMedida] LONG(4),
  [IDCalibracion] LONG(4)
);
-- Rows: 302

CREATE TABLE [TbEventosEquipoMedida_antes] (
  [IDEventoEquipoMedida] LONG(4),
  [IDEvento] TEXT(255),
  [IDEquipoMedida] LONG(4)
);
-- Rows: 26

CREATE TABLE [TbFacturaActividadesInvolucradas] (
  [IDFactura] LONG(4),
  [IDActividad] TEXT(50),
  [NODO] TEXT(50),
  [BUI] TEXT(50),
  [SUBSISTEMA] TEXT(255),
  [EQUIPO] TEXT(255),
  [IDEVENTO] TEXT(50),
  [TIPOTECNICO] TEXT(50),
  [ALIASTECNICO] TEXT(50),
  [PERFIL] TEXT(50),
  [PrecioHoraLaborablePerfil] DOUBLE(8),
  [PrecioHoraExtraPerfil] DOUBLE(8),
  [TIPOACTIVIDAD] TEXT(255),
  [DESCRIPCION] MEMO,
  [FECHAALTA] SHORT_DATE_TIME(8),
  [HORASLABORABLES] DOUBLE(8),
  [HORASEXTRAS] DOUBLE(8),
  [UBICACION] TEXT(50)
);
-- Rows: 25107

CREATE TABLE [TbFacturaEventosInvolucrados] (
  [IDFactura] LONG(4),
  [IDEvento] TEXT(50),
  [NODO] TEXT(50),
  [BUI] TEXT(50),
  [SUBSISTEMA] TEXT(255),
  [EQUIPO] TEXT(255),
  [PMPR] TEXT(50),
  [TIPOEVENTO] TEXT(50),
  [CRITICIDAD] FLOAT(4),
  [TIPOTECNICO] TEXT(50),
  [ORIGINADOR] TEXT(50),
  [FECHAALTA] SHORT_DATE_TIME(8),
  [HORAALTA] SHORT_DATE_TIME(8),
  [TIEMPORESPUESTA] SHORT_DATE_TIME(8),
  [FECHAFIN] SHORT_DATE_TIME(8),
  [HORAFIN] SHORT_DATE_TIME(8),
  [CAUSAFIN] MEMO,
  [CONTACTO] TEXT(255),
  [IDEVENTOSGENERADOS] TEXT(255),
  [DESCRIPCION] MEMO,
  [ParaRegularizar] TEXT(255)
);
-- Rows: 5514

CREATE TABLE [TbFacturaGastosInvolucrados] (
  [IDFactura] LONG(4),
  [IDGasto] LONG(4),
  [TipoGasto] TEXT(255),
  [ImporteUnitario] DOUBLE(8),
  [NumeroUnidades] DOUBLE(8),
  [FechaImputacionGasto] SHORT_DATE_TIME(8),
  [ALIAS] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 519

CREATE TABLE [TbFacturaMaterialesInvolucrados] (
  [IDFactura] LONG(4),
  [IDMaterial] TEXT(50),
  [NODO] TEXT(50),
  [BUI] TEXT(50),
  [SUBSISTEMA] TEXT(255),
  [EQUIPO] TEXT(255),
  [IDEvento] TEXT(50),
  [TIPOACCION] TEXT(255),
  [Material] TEXT(255),
  [PN] TEXT(50),
  [NS] TEXT(50),
  [DESCRIPCIONEQUIPO] TEXT(255),
  [FECHAENTREGA] SHORT_DATE_TIME(8),
  [PRECIOSINIVA] DOUBLE(8),
  [ReparadoPor] TEXT(255)
);
-- Rows: 1986

CREATE TABLE [TbFacturaPrincipal] (
  [IDFactura] LONG(4),
  [FechaFinalEventos] SHORT_DATE_TIME(8),
  [FechaFactura] SHORT_DATE_TIME(8),
  [Titulo1] TEXT(255),
  [Titulo2] TEXT(255),
  [Titulo3] TEXT(255),
  [TipoImpositivo] DOUBLE(8),
  [Recargo] DOUBLE(8),
  [PrecioHoras] DOUBLE(8),
  [PrecioSuministro] DOUBLE(8),
  [PrecioSuministroConRecargo] DOUBLE(8),
  [PrecioDietasYGastos] DOUBLE(8),
  [PrecioIncurrido] DOUBLE(8),
  [PrecioSubcontrataciones] DOUBLE(8),
  [PrecioSubcontratacionesConRecargo] DOUBLE(8),
  [PrecioBaseImponible] DOUBLE(8),
  [PrecioImpuestos] DOUBLE(8),
  [ImporteAFacturar] DOUBLE(8),
  [NombreFacturaAnexa] TEXT(255)
);
-- Rows: 53

CREATE TABLE [TbFacturaPrincipalPerfiles] (
  [IDFacturaPerfil] LONG(4),
  [IDFactura] LONG(4),
  [Perfil] TEXT(255),
  [PrecioHoraLab] DOUBLE(8),
  [PrecioHoraExt] DOUBLE(8),
  [NHorasLab] DOUBLE(8),
  [NHorasExt] DOUBLE(8)
);
-- Rows: 304

CREATE TABLE [TbFacturaSubcontratacionesInvolucradas] (
  [IDFactura] LONG(4),
  [IDSubcontratacion] LONG(4),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [Empresa] TEXT(255),
  [Descripcion] MEMO,
  [Importe] DOUBLE(8)
);
-- Rows: 74

CREATE TABLE [TbGastos] (
  [IDGasto] LONG(4),
  [TipoGasto] TEXT(255),
  [ImporteUnitario] DOUBLE(8),
  [NumeroUnidades] DOUBLE(8),
  [FechaImputacionGasto] SHORT_DATE_TIME(8),
  [ALIAS] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 519

CREATE TABLE [TbGastosImportePorTipo] (
  [TipoGasto] TEXT(255),
  [FechaValorInicial] SHORT_DATE_TIME(8),
  [FechaValorFinal] SHORT_DATE_TIME(8),
  [ImporteUnitario] DOUBLE(8),
  [FechaFinalImporte] TEXT(50),
  [Observaciones] MEMO
);
-- Rows: 3

CREATE TABLE [TbGuiaConciliaciones] (
  [IDGuia] LONG(4),
  [Parte] TEXT(255),
  [Nodo] TEXT(255),
  [BUI] TEXT(255),
  [SubSistema] TEXT(255),
  [EquipoID] TEXT(255),
  [IDEvento] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbHerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 4

CREATE TABLE [TbMaterial] (
  [IDMaterial] TEXT(255),
  [IDEquipo] LONG(4),
  [FechaIncio] SHORT_DATE_TIME(8),
  [Garantia] TEXT(50),
  [IDActividad] TEXT(50),
  [TipoAccion] TEXT(255),
  [Material] TEXT(255),
  [PN] TEXT(50),
  [NS] TEXT(50),
  [COSTE] DOUBLE(8),
  [ReparadoPor] TEXT(255),
  [FechaEntrega] SHORT_DATE_TIME(8),
  [Descripcion] MEMO,
  [CodImportacion] TEXT(50),
  [EsReparacion] TEXT(2)
);
-- Rows: 2351

CREATE TABLE [TbMaterialSeguimiento] (
  [IDSeguimiento] TEXT(50),
  [IDMaterial] TEXT(50),
  [FechaInicioIntervencion] SHORT_DATE_TIME(8),
  [Destino] TEXT(255),
  [Motivo] MEMO,
  [FechaFinIntervencion] SHORT_DATE_TIME(8),
  [Resultado] MEMO,
  [UltimaIntervencion] TEXT(2),
  [TituloAnexo] TEXT(255),
  [DescripcionAnexo] MEMO,
  [NombreArchivoAnexo] TEXT(255)
);
-- Rows: 2049

CREATE TABLE [TbNodoBUI] (
  [NODO] TEXT(50),
  [BUI] TEXT(50)
);
-- Rows: 19

CREATE TABLE [TbNodos] (
  [NODO] TEXT(50)
);
-- Rows: 6

CREATE TABLE [TbOriginador] (
  [ORIGINADOR] TEXT(50)
);
-- Rows: 7

CREATE TABLE [TbPartesDetalle] (
  [IDParte] TEXT(50),
  [IDEvento] TEXT(50),
  [FechaAlta] SHORT_DATE_TIME(8),
  [HoraAlta] SHORT_DATE_TIME(8),
  [Descripcion] MEMO,
  [CausaFin] MEMO,
  [FechaFin] SHORT_DATE_TIME(8),
  [HoraFin] SHORT_DATE_TIME(8),
  [AliasCreador] TEXT(50),
  [Originador] TEXT(255),
  [TiempoRespuesta] SHORT_DATE_TIME(8),
  [SubSistema] TEXT(255),
  [Equipo] TEXT(255),
  [PMPR] TEXT(50)
);
-- Rows: 8819

CREATE TABLE [TbPartesPpal] (
  [IDParte] TEXT(50),
  [FechaObtencion] SHORT_DATE_TIME(8),
  [FechaInicial] SHORT_DATE_TIME(8),
  [FechaFinal] SHORT_DATE_TIME(8),
  [Descripcion] MEMO,
  [NombreArchivoAdjunto] TEXT(255),
  [IDFactura] LONG(4),
  [Centro] TEXT(255)
);
-- Rows: 342

CREATE TABLE [TbPlanificacion] (
  [IDPlanificacion] LONG(4),
  [FechaRegistro] SHORT_DATE_TIME(8),
  [IDEquipo] LONG(4),
  [ANIO] LONG(4),
  [Mes] INT(2),
  [Semana] INT(2),
  [MotivoCierre] TEXT(50),
  [IDEvento] TEXT(50),
  [IDNoEvento] LONG(4),
  [FechaPrevistaCierre] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [MotivacionReprogramacion] MEMO,
  [IDNuevaPlanificacion] LONG(4)
);
-- Rows: 624

CREATE TABLE [TbPlanificacionAnexos] (
  [IDAnexoPlanificacion] LONG(4),
  [IDPlanificacion] LONG(4),
  [NombreAnexo] TEXT(255),
  [UsuarioAnexa] TEXT(50),
  [FechaAnexo] SHORT_DATE_TIME(8),
  [URLAnexo] TEXT(255)
);
-- Rows: 5

CREATE TABLE [TbPlanificacionEquipos] (
  [IDEquipo] LONG(4),
  [Alias] TEXT(50),
  [PeriodicidadEnMesesRecomendada] INT(2),
  [FechaAltaParaPlanificacion] SHORT_DATE_TIME(8),
  [FechabajaParaPlanificacion] SHORT_DATE_TIME(8)
);
-- Rows: 82

CREATE TABLE [TbPlanificacionRegistrada] (
  [ID] LONG(4),
  [IDEquipo] LONG(4),
  [FechaMantoRealizado] SHORT_DATE_TIME(8),
  [Observaciones] MEMO,
  [NombreAnexo] TEXT(255)
);
-- Rows: 23

CREATE TABLE [TbRepuestosTipo] (
  [TIPOREPUESTOS] TEXT(255)
);
-- Rows: 3

CREATE TABLE [TbResultadoVerificacion] (
  [RESULTADOVERIFICACION] TEXT(50)
);
-- Rows: 3

CREATE TABLE [TbSistema] (
  [SISTEMA] TEXT(50)
);
-- Rows: 2

CREATE TABLE [TbSubcontrataciones] (
  [IDSubcontratacion] LONG(4),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [Empresa] TEXT(255),
  [Descripcion] MEMO,
  [Importe] DOUBLE(8),
  [Observaciones] MEMO
);
-- Rows: 75

CREATE TABLE [TbSubsistemaBui] (
  [SUBSISTEMA] TEXT(255),
  [BUI] TEXT(50)
);
-- Rows: 162

CREATE TABLE [TbTecnicos] (
  [ALIAS] TEXT(50),
  [TIPO] TEXT(50),
  [NOMBRE] TEXT(255),
  [TELEFONO] TEXT(50),
  [EMAIL] TEXT(255),
  [FECHAALTA] SHORT_DATE_TIME(8),
  [FECHABAJA] SHORT_DATE_TIME(8),
  [DNI] TEXT(255)
);
-- Rows: 115

CREATE TABLE [TbTecnicosAusencias] (
  [IDAusencia] LONG(4),
  [Alias] TEXT(50),
  [FechaLibranza] SHORT_DATE_TIME(8),
  [Observaciones] MEMO,
  [Motivo] TEXT(255),
  [horas] DOUBLE(8)
);
-- Rows: 1236

CREATE TABLE [TbTecnicosFiestas] (
  [FechaFiesta] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 104

CREATE TABLE [TbTipoAccion] (
  [ESTADOREPARABLE] TEXT(255)
);
-- Rows: 8

CREATE TABLE [TbTipoAsistencia] (
  [TIPOASISTENCIA] TEXT(255)
);
-- Rows: 4

CREATE TABLE [TbTipoEvento] (
  [TIPOEVENTO] TEXT(50)
);
-- Rows: 6

CREATE TABLE [TbTipoTecnico] (
  [TIPO] TEXT(255)
);
-- Rows: 10

CREATE TABLE [TbTipoTecnicoPrecios] (
  [IDTipoTecnicoFecha] LONG(4),
  [TIPOTECNICO] TEXT(50),
  [FechaInicial] SHORT_DATE_TIME(8),
  [FechaFinal] SHORT_DATE_TIME(8),
  [PrecioHoraLab] DOUBLE(8),
  [PrecioHoraExtra] DOUBLE(8)
);
-- Rows: 19

CREATE TABLE [TbUbicacion] (
  [UBICACION] TEXT(50)
);
-- Rows: 3

-- Summary: 57 local tables, 0 linked/skipped
