-- Schema dump: Registro_Datos.accdb
-- Tables: 5

CREATE TABLE [DClases] (
  [Cod] TEXT(1),
  [DClase] TEXT(10)
);
-- Rows: 9

CREATE TABLE [Entradas] (
  [IDEntrada] LONG(4),
  [NEntrada] TEXT(255),
  [FEntrada] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [Procedencia] TEXT(255),
  [Clase] TEXT(1),
  [Expediente] TEXT(255),
  [Entrega] TEXT(255),
  [Extracto] TEXT(255),
  [Archivo] TEXT(3),
  [Anexo] TEXT(255),
  [DPD] TEXT(50),
  [Remitente] TEXT(255),
  [ArchivadoPor] TEXT(255),
  [Observaciones] MEMO,
  [Encriptado] TEXT(2)
);
-- Rows: 14905

CREATE TABLE [Nombres] (
  [Cod] TEXT(3),
  [Nombre] TEXT(50)
);
-- Rows: 57

CREATE TABLE [Salidas] (
  [IDSalida] LONG(4),
  [NSalida] TEXT(255),
  [FSalida] SHORT_DATE_TIME(8),
  [FDocumento] SHORT_DATE_TIME(8),
  [Destino] TEXT(255),
  [Clase] TEXT(1),
  [Extracto] TEXT(255),
  [Anexo] TEXT(255),
  [DPD] TEXT(50),
  [ArchivadoPor] TEXT(255),
  [Expediente] TEXT(255),
  [Observaciones] MEMO,
  [Encriptado] TEXT(2)
);
-- Rows: 13630

CREATE TABLE [TbConfiguracion] (
  [DirectorioPorDefectoAnexosEntrada] TEXT(50),
  [DirectorioPorDefectoAnexosSalida] TEXT(50)
);
-- Rows: 1

-- Summary: 5 local tables, 0 linked/skipped
