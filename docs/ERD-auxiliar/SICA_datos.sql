-- Schema dump: SICA_datos.accdb
-- Tables: 4

CREATE TABLE [TbAperturas] (
  [FechaApertura] TEXT(255),
  [Usuario] TEXT(255),
  [URL] TEXT(255)
);
-- Rows: 7800

CREATE TABLE [TbArchivos] (
  [URL] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCierres] (
  [FechaCierre] TEXT(255),
  [Usuario] TEXT(255)
);
-- Rows: 54

CREATE TABLE [TbUltimaImportacion] (
  [FechaImportacion] SHORT_DATE_TIME(8),
  [FechaArchivo] SHORT_DATE_TIME(8),
  [Terminal] TEXT(255),
  [UsuarioImporta] TEXT(255),
  [URLCarpetaLocal] TEXT(255)
);
-- Rows: 1

-- Summary: 4 local tables, 0 linked/skipped
