-- Schema dump: Solicitudes_HPS_datos.accdb
-- Tables: 20

CREATE TABLE [Copia de TbExpedientes] (
  [IDExpediente] LONG(4),
  [Cod_Exp] TEXT(255),
  [OC] TEXT(255),
  [Grado] TEXT(255),
  [Expediente] TEXT(255),
  [F_Inicio] SHORT_DATE_TIME(8),
  [F_Fin] SHORT_DATE_TIME(8),
  [IDJuridicaContrato] LONG(4),
  [IDExpedienteNuevo] LONG(4)
);
-- Rows: 70

CREATE TABLE [TbConfiguracion] (
  [ID] LONG(4),
  [DiasParaRecordatorioExcel1] INT(2),
  [DiasParaRecordatorioExcel2] INT(2),
  [DiasCancelacionPreMARGA] INT(2),
  [DiasParaRecordatorioRellenoMarga1] INT(2),
  [DiasParaRecordatorioRellenoMarga2] INT(2),
  [DiasCancelacionMARGA] INT(2),
  [BuzonSeguridad] TEXT(255),
  [EmailDirectorSeguridad] TEXT(255),
  [CorreodeEnvio] TEXT(255),
  [AutocancelacionPreMARGA] TEXT(2),
  [AutocancelacionMARGA] TEXT(2),
  [CorreosAutomaticos] TEXT(2),
  [VersionPlantillasHTML] TEXT(255),
  [VersionPlantillasExcel] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbCorreosEnviados] (
  [IDCorreo] LONG(4),
  [URLAdjunto] MEMO,
  [Aplicacion] TEXT(255),
  [Destinatarios] MEMO,
  [DestinatariosConCopia] MEMO,
  [DestinatariosConCopiaOculta] MEMO,
  [Asunto] TEXT(255),
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaOrdenEnvio] SHORT_DATE_TIME(8),
  [FechaGrabacion] SHORT_DATE_TIME(8),
  [NombrePlantilla] TEXT(255),
  [VersionPlantilla] TEXT(255),
  [CadenaRecursos] MEMO,
  [IDSolicitud] LONG(4),
  [Accion] TEXT(255),
  [DesencadenadoPor] TEXT(255),
  [Programado] TEXT(2),
  [TipoCorreo] LONG(4),
  [Observaciones] MEMO,
  [Intentos] LONG(4),
  [FechaProceso] SHORT_DATE_TIME(8)
);
-- Rows: 9

-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesLugaresEjecucion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbHPS] (given file does not exist: C:\00repos\datos\HPST.accdb)
-- LINKED/SKIPPED: [TbHPSEquivalencia] (given file does not exist: C:\00repos\datos\HPST.accdb)
CREATE TABLE [TbHPSGrado] (
  [TipoHPS] TEXT(255),
  [Grado] TEXT(255)
);
-- Rows: 13

-- LINKED/SKIPPED: [TbHPSGrado1] (given file does not exist: C:\00repos\datos\HPST.accdb)
CREATE TABLE [TbJustificaciones] (
  [idjustificacion] LONG(4),
  [titulo] TEXT(255),
  [descripcion] MEMO,
  [activa] BOOLEAN(1)
);
-- Rows: 7

CREATE TABLE [TbLogs] (
  [IDLog] LONG(4),
  [IDSolicitud] LONG(4),
  [TipoOperacion] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Usuario] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 0

CREATE TABLE [TbLogsGeneral] (
  [IDLog] LONG(4),
  [Fecha] SHORT_DATE_TIME(8),
  [Usuario] TEXT(255),
  [Accion] TEXT(255),
  [Descripcion] MEMO,
  [DNI] TEXT(255),
  [CorreosAutomaticos] TEXT(2),
  [IDSolicitud] LONG(4)
);
-- Rows: 2058

-- LINKED/SKIPPED: [TbMotivoHPS] (given file does not exist: C:\00repos\datos\HPST.accdb)
CREATE TABLE [TbResponsables] (
  [IDResponsable] LONG(4),
  [Nombre] TEXT(255),
  [Correo] TEXT(255),
  [Descripcion] MEMO
);
-- Rows: 27

CREATE TABLE [TbSolicitudes] (
  [IDSolicitud] LONG(4),
  [Gestor] TEXT(255),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido1] TEXT(255),
  [Apellido2] TEXT(255),
  [FNacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [email] TEXT(255),
  [Telefono] TEXT(255),
  [IDExpediente] LONG(4),
  [IDEmpresaUsuario] LONG(4),
  [IDEmpresaTramitadora] LONG(4),
  [URLAdjunto] MEMO,
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaUltimoCambio] SHORT_DATE_TIME(8),
  [UsuarioCreacion] TEXT(255),
  [UsuarioUltimoCambio] TEXT(255),
  [emailResponsable] TEXT(255),
  [Estado] TEXT(255),
  [Observaciones] MEMO,
  [Tipo] TEXT(255),
  [CorreosAutomaticos] TEXT(2),
  [Motivo_HPS] MEMO,
  [IDUsuarioHPS] LONG(4),
  [URLAdjuntoEnvioONS] MEMO,
  [Nemotecnico] TEXT(255),
  [idjustificacion] LONG(4)
);
-- Rows: 245

CREATE TABLE [TbSolicitudesFechas] (
  [IDSolicitud] LONG(4),
  [FechaEnvioExcel] SHORT_DATE_TIME(8),
  [FechaRecepcionExcel] SHORT_DATE_TIME(8),
  [FechaTramitacionAltaMarga] SHORT_DATE_TIME(8),
  [FechaEnvioDPS] SHORT_DATE_TIME(8),
  [FechaEnvioONS] SHORT_DATE_TIME(8),
  [FechaRegistroEnHPS] SHORT_DATE_TIME(8),
  [FechaCorreoRecordatorioExcel1] SHORT_DATE_TIME(8),
  [FechaCorreoRecordatorioExcel2] SHORT_DATE_TIME(8),
  [FechaCorreoRecordatorioRellenoMarga1] SHORT_DATE_TIME(8),
  [FechaCorreoRecordatorioRellenoMarga2] SHORT_DATE_TIME(8),
  [FechaDesestimado] SHORT_DATE_TIME(8),
  [FechaAutocancelacion] SHORT_DATE_TIME(8),
  [FechaPrevistaCorreoRecordatorioExcel1] SHORT_DATE_TIME(8),
  [FechaPrevistaCorreoRecordatorioExcel2] SHORT_DATE_TIME(8),
  [FechaPrevistaCorreoRecordatorioRellenoMarga1] SHORT_DATE_TIME(8),
  [FechaPrevistaCorreoRecordatorioRellenoMarga2] SHORT_DATE_TIME(8),
  [FechaPrevistaAutocancelacionPreMarga] SHORT_DATE_TIME(8),
  [FechaPrevistaAutocancelacionMarga] SHORT_DATE_TIME(8),
  [FechaCancelado] SHORT_DATE_TIME(8),
  [FechaEnvioTraspasoONS] SHORT_DATE_TIME(8)
);
-- Rows: 245

CREATE TABLE [TbUltimoCambio] (
  [ID] LONG(4),
  [IDSolicitud] LONG(4),
  [FechaCambio] SHORT_DATE_TIME(8),
  [IDUsuarioCambio] LONG(4)
);
-- Rows: 5

-- LINKED/SKIPPED: [TbUsuarios] (given file does not exist: C:\00repos\datos\HPST.accdb)
-- LINKED/SKIPPED: [TbUsuariosEntidades] (given file does not exist: C:\00repos\datos\HPST.accdb)
-- LINKED/SKIPPED: [TbUsuariosHistoricos] (given file does not exist: C:\00repos\datos\HPST.accdb)
-- Summary: 11 local tables, 9 linked/skipped
