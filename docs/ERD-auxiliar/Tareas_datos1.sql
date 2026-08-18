-- Schema dump: Tareas_datos1.accdb
-- Tables: 17

CREATE TABLE [TbCorreosEnviados] (
  [IDCorreo] LONG(4),
  [Aplicacion] TEXT(255),
  [Destinatarios] MEMO,
  [DestinatariosConCopiaOculta] MEMO,
  [Asunto] TEXT(255),
  [Cuerpo] MEMO,
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaGrabacion] SHORT_DATE_TIME(8),
  [CuerpoHTML] BOOLEAN(1),
  [DestinatariosConCopia] MEMO
);
-- Rows: 5202

-- LINKED/SKIPPED: [TbDPDOfertaSuministrador] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesResponsables] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbFacturasDetalle] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbHTMLPorUsuario] (
  [Usuario] TEXT(255),
  [HTMLTareas] MEMO,
  [NumeroTareas] INT(2),
  [DirCorreo] TEXT(255)
);
-- Rows: 0

-- LINKED/SKIPPED: [TbNPedido] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbProyectos] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbSolicitudesOfertasPrevias] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbSuministradoresSAP] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
CREATE TABLE [TbTareas] (
  [Tarea] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Observaciones] MEMO,
  [Realizado] TEXT(2)
);
-- Rows: 11

CREATE TABLE [TbTipoTareas] (
  [Tarea] TEXT(255),
  [PeriodicidadDias] LONG(4)
);
-- Rows: 9

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesTareas] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbVisadoFacturas_Nueva] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- LINKED/SKIPPED: [TbVisadosGenerales] (given file does not exist: C:\00repos\datos\AGEDYS_DATOS.accdb)
-- Summary: 4 local tables, 13 linked/skipped
