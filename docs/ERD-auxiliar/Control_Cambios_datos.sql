-- Schema dump: Control_Cambios_datos.accdb
-- Tables: 3

CREATE TABLE [TbAplicacionesVersiones] (
  [IDVersion] LONG(4),
  [IDAplicacion] LONG(4),
  [Version] TEXT(255),
  [Descripcion] MEMO,
  [FechaInicio] SHORT_DATE_TIME(8),
  [FechaFin] SHORT_DATE_TIME(8),
  [VersionPrincipal] TEXT(2),
  [URLRequisitosPrevios] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbCambioDocumentos] (
  [IDDoc] LONG(4),
  [IDCambio] LONG(4),
  [URLDocumento] TEXT(255)
);
-- Rows: 19

CREATE TABLE [TbVersionCambios] (
  [IDCambio] LONG(4),
  [IDVersion] LONG(4),
  [NombreCambio] TEXT(255),
  [DescripcionCambio] MEMO,
  [EsBug] TEXT(2),
  [EsMejora] TEXT(2),
  [MostrarEninformeGeneral] TEXT(2),
  [CodCambio] TEXT(255)
);
-- Rows: 34

-- Summary: 3 local tables, 0 linked/skipped
