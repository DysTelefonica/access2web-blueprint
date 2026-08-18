-- Schema dump: Correos_datos.accdb
-- Tables: 1

CREATE TABLE [TbCorreosEnviados] (
  [IDCorreo] LONG(4),
  [URLAdjunto] MEMO,
  [Aplicacion] TEXT(255),
  [Originador] TEXT(50),
  [Destinatarios] MEMO,
  [DestinatariosConCopia] MEMO,
  [DestinatariosConCopiaOculta] MEMO,
  [Asunto] TEXT(255),
  [Cuerpo] MEMO,
  [FechaEnvio] SHORT_DATE_TIME(8),
  [Observaciones] MEMO,
  [NDPD] TEXT(50),
  [NPEDIDO] TEXT(50),
  [NFACTURA] TEXT(50),
  [FechaGrabacion] SHORT_DATE_TIME(8),
  [CuerpoHTML] BOOLEAN(1),
  [IDEdicion] INT(2)
);
-- Rows: 35374

-- Summary: 1 local tables, 0 linked/skipped
