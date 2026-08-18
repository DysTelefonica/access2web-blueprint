-- Schema dump: Registro_Ent_Salida_Datos.accdb
-- Tables: 12

-- LINKED/SKIPPED: [Salidas] (given file does not exist: C:\00repos\datos\Registro_Datos.accdb)
CREATE TABLE [TbConexiones] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 12

CREATE TABLE [TbEntradaDestinatarios_RES] (
  [IDEntrada] LONG(4),
  [Destinatario] TEXT(255)
);
-- Rows: 37

CREATE TABLE [TbEntradaPersonasContacto_RES] (
  [IDEntrada] LONG(4),
  [PersonaContacto] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbEntradas_indexacion] (
  [IDEntrada] LONG(4),
  [Hash] TEXT(255),
  [FechaIndexacion] SHORT_DATE_TIME(8),
  [TextoArchivo] MEMO
);
-- Rows: 0

CREATE TABLE [TbEntradas_RES] (
  [IDEntrada] LONG(4),
  [NEntrada] TEXT(255),
  [Juridica] TEXT(25),
  [FEntrada] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [NRef] TEXT(255),
  [Remitente] TEXT(255),
  [Destinatario] TEXT(255),
  [ExpActividad] TEXT(255),
  [Soporte] TEXT(255),
  [Asunto] TEXT(255),
  [NombreAnexo] TEXT(255),
  [UbicacionFisica] TEXT(255),
  [emailInteresado] TEXT(255),
  [Notas] MEMO,
  [Verificado] TEXT(2),
  [UsuarioRegistro] TEXT(255),
  [NDPD] TEXT(255),
  [Encriptado] TEXT(2)
);
-- Rows: 635

CREATE TABLE [TbEntradas_RES_Indexacion] (
  [IDEntrada] LONG(4),
  [Hash] TEXT(255),
  [FechaIndexacion] SHORT_DATE_TIME(8),
  [TextoArchivo] MEMO
);
-- Rows: 2

CREATE TABLE [TbLog] (
  [IDLog] LONG(4),
  [IDEntrada] LONG(4),
  [IDSalida] LONG(4),
  [IDCorreo] LONG(4),
  [Usuario] TEXT(255),
  [Fecha] SHORT_DATE_TIME(8),
  [Titulo] TEXT(255),
  [Linea] MEMO
);
-- Rows: 3341

CREATE TABLE [TbSalidaDestinatarios_RES] (
  [IDSalida] LONG(4),
  [Destinatario] TEXT(255)
);
-- Rows: 5

CREATE TABLE [TbSalidas_Indexacion] (
  [IDSalida] LONG(4),
  [Hash] TEXT(255),
  [FechaIndexacion] SHORT_DATE_TIME(8),
  [TextoArchivo] MEMO
);
-- Rows: 0

CREATE TABLE [TbSalidas_RES] (
  [IDSalida] LONG(4),
  [NSalida] TEXT(255),
  [Juridica] TEXT(25),
  [FSalida] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [NRef] TEXT(255),
  [Remitente] TEXT(255),
  [Destinatario] TEXT(255),
  [ExpActividad] TEXT(255),
  [Soporte] TEXT(255),
  [Asunto] TEXT(255),
  [NombreAnexo] TEXT(255),
  [UbicacionFisica] TEXT(255),
  [emailRemitente] TEXT(255),
  [CodAGEDO] TEXT(255),
  [Notas] MEMO,
  [Verificado] TEXT(2),
  [UsuarioRegistro] TEXT(255),
  [NDPD] TEXT(255),
  [MedioEnvio] TEXT(255),
  [Encriptado] TEXT(2)
);
-- Rows: 4055

CREATE TABLE [TbSalidas_RES_Indexacion] (
  [IDSalida] LONG(4),
  [Hash] TEXT(255),
  [FechaIndexacion] SHORT_DATE_TIME(8),
  [TextoArchivo] MEMO
);
-- Rows: 1063

-- Summary: 11 local tables, 1 linked/skipped
