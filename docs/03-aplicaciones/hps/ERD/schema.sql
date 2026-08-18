-- Schema dump: HPST.accdb
-- Tables: 31

CREATE TABLE [Copia de TbExpedienteLugares] (
  [IDExpLugar] LONG(4),
  [IDExpediente] LONG(4),
  [Lugar] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 87

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

CREATE TABLE [Copia de TbUsuarios] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [Correo_e] TEXT(255),
  [IDExpediente] LONG(4),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(2),
  [Requiere_Curso] TEXT(2),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [F_Baja] SHORT_DATE_TIME(8),
  [LugarPrestacionServicio] MEMO,
  [IDEmpresaUsuario] LONG(4),
  [IDEmpresaHPS] LONG(4),
  [IDJuridicaContrato] LONG(4),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(2),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [IDSolicitud] LONG(4)
);
-- Rows: 335

CREATE TABLE [Copia de TbUsuariosEntidades] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [EmpresaUsuario] TEXT(255),
  [EmpresaTramitadora] TEXT(255),
  [CadenaContratistas] TEXT(255),
  [Correo_e] TEXT(255),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [IDExpediente] LONG(4),
  [CodExp] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(2),
  [F_Baja] SHORT_DATE_TIME(8),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(2),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [Curso_Realizado] TEXT(2),
  [Requiere_Curso] TEXT(2),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [Requiere_PrimeraConvocatoriaCurso] TEXT(2),
  [Requiere_SegundaConvocatoriaCurso] TEXT(2),
  [Requiere_CorreoJefeSeguridadCurso] TEXT(2),
  [Observaciones] MEMO,
  [HPS_NAC_SIN_DATOS] TEXT(2),
  [HPS_NAC_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_NAC_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_NAC_Grado] TEXT(255),
  [HPS_NAC_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_NAC_Especialidad] TEXT(255),
  [HPS_NAC_Renovacion] TEXT(2),
  [HPS_NAC_F_Baja] SHORT_DATE_TIME(8),
  [HPS_NAC_Activo] TEXT(2),
  [HPS_NAC_ApuntoDeCaducar] TEXT(2),
  [HPS_NAC_Baja] TEXT(2),
  [HPS_NAC_Caducado] TEXT(2),
  [HPS_NAC_PendienteRenovacion] TEXT(2),
  [HPS_NAC_Solicitado] TEXT(2),
  [HPS_NAC_ESTADO] TEXT(255),
  [HPS_NAC_Irregular] TEXT(2),
  [HPS_NAC_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_NAC_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_OTAN_SIN_DATOS] TEXT(2),
  [HPS_OTAN_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_OTAN_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_OTAN_Grado] TEXT(255),
  [HPS_OTAN_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_OTAN_Especialidad] TEXT(255),
  [HPS_OTAN_Renovacion] TEXT(2),
  [HPS_OTAN_F_Baja] SHORT_DATE_TIME(8),
  [HPS_OTAN_Activo] TEXT(2),
  [HPS_OTAN_ApuntoDeCaducar] TEXT(2),
  [HPS_OTAN_Baja] TEXT(2),
  [HPS_OTAN_Caducado] TEXT(2),
  [HPS_OTAN_PendienteRenovacion] TEXT(2),
  [HPS_OTAN_Solicitado] TEXT(2),
  [HPS_OTAN_ESTADO] TEXT(255),
  [HPS_OTAN_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_OTAN_Irregular] TEXT(2),
  [HPS_OTAN_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_ESA_SIN_DATOS] TEXT(2),
  [HPS_ESA_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_ESA_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_ESA_Grado] TEXT(255),
  [HPS_ESA_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_ESA_Especialidad] TEXT(255),
  [HPS_ESA_Renovacion] TEXT(2),
  [HPS_ESA_F_Baja] SHORT_DATE_TIME(8),
  [HPS_ESA_Activo] TEXT(2),
  [HPS_ESA_ApuntoDeCaducar] TEXT(2),
  [HPS_ESA_Baja] TEXT(2),
  [HPS_ESA_Caducado] TEXT(2),
  [HPS_ESA_PendienteRenovacion] TEXT(2),
  [HPS_ESA_Solicitado] TEXT(2),
  [HPS_ESA_ESTADO] TEXT(255),
  [HPS_ESA_Irregular] TEXT(2),
  [HPS_ESA_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_ESA_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_UE_SIN_DATOS] TEXT(2),
  [HPS_UE_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_UE_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_UE_Grado] TEXT(255),
  [HPS_UE_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_UE_Especialidad] TEXT(255),
  [HPS_UE_Renovacion] TEXT(2),
  [HPS_UE_F_Baja] SHORT_DATE_TIME(8),
  [HPS_UE_Activo] TEXT(2),
  [HPS_UE_ApuntoDeCaducar] TEXT(2),
  [HPS_UE_Baja] TEXT(2),
  [HPS_UE_Caducado] TEXT(2),
  [HPS_UE_PendienteRenovacion] TEXT(2),
  [HPS_UE_Solicitado] TEXT(2),
  [HPS_UE_ESTADO] TEXT(255),
  [HPS_UE_Irregular] TEXT(2),
  [HPS_UE_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_UE_MESES_PARA_RENOVAR] TEXT(255)
);
-- Rows: 331

CREATE TABLE [Errores de pegado] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [Correo_e] TEXT(255),
  [IDExpediente] LONG(4),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(255),
  [Requiere_Curso] TEXT(255),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [F_Baja] SHORT_DATE_TIME(8),
  [LugarPrestacionServicio] MEMO,
  [IDEmpresaUsuario] LONG(4),
  [IDEmpresaHPS] LONG(4),
  [IDJuridicaContrato] LONG(4),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(255),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [CadenaContratistas] TEXT(255)
);
-- Rows: 1

CREATE TABLE [TbAnexosUsuariosHistoricos] (
  [IDAnexo] LONG(4),
  [NombreAnexo] TEXT(255),
  [IDUsuario] LONG(4),
  [Hash] TEXT(64)
);
-- Rows: 1395

CREATE TABLE [TbAnexosUsuariosHPS] (
  [IDAnexo] LONG(4),
  [NombreAnexo] TEXT(255),
  [IDUsuario] LONG(4),
  [Hash] TEXT(64),
  [EsHistorico] TEXT(2)
);
-- Rows: 2183

CREATE TABLE [TbAnexosUsuariosSICA] (
  [IDAnexo] LONG(4),
  [NombreAnexo] TEXT(255),
  [IDUsuarioSICA] TEXT(255),
  [Hash] TEXT(64)
);
-- Rows: 9

CREATE TABLE [TbAuxCursos] (
  [ID] LONG(4)
);
-- Rows: 90

CREATE TABLE [TbConfiguracionUsuariosAccesoLocal] (
  [ID] LONG(4),
  [CorreoUsuario] TEXT(255),
  [Activo] BOOLEAN(1),
  [FechaAlta] SHORT_DATE_TIME(8),
  [Observaciones] MEMO
);
-- Rows: 4

CREATE TABLE [TbConsultas] (
  [IDConsulta] LONG(4),
  [Nombre] TEXT(255)
);
-- Rows: 8

-- LINKED/SKIPPED: [TbExpedientes] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesConEntidades] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
-- LINKED/SKIPPED: [TbExpedientesLugaresEjecucion] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbHPS] (
  [IDUsuario] LONG(4),
  [TipoHPS] TEXT(255),
  [F_Concesion] SHORT_DATE_TIME(8),
  [F_Caducidad] SHORT_DATE_TIME(8),
  [Grado] TEXT(255),
  [F_Baja] SHORT_DATE_TIME(8),
  [Renovacion] TEXT(2),
  [F_Solicitud] SHORT_DATE_TIME(8),
  [Especialidad] TEXT(255)
);
-- Rows: 1360

CREATE TABLE [TbHPSEquivalencia] (
  [GradoNacional] TEXT(255),
  [GradoOTAN] TEXT(255),
  [GradoESA] TEXT(255),
  [GradoUE] TEXT(255)
);
-- Rows: 3

CREATE TABLE [TbHPSGrado] (
  [TipoHPS] TEXT(255),
  [Grado] TEXT(255)
);
-- Rows: 13

CREATE TABLE [TbJuridicasContratacion] (
  [Juridica] TEXT(255)
);
-- Rows: 3

CREATE TABLE [TbMotivoHPS] (
  [MotivoHPS] TEXT(255)
);
-- Rows: 58

CREATE TABLE [TbObservaciones] (
  [IDObservacion] LONG(4),
  [ID] LONG(4),
  [Fecha] SHORT_DATE_TIME(8),
  [Observacion] MEMO,
  [Tipo] TEXT(255)
);
-- Rows: 340

CREATE TABLE [TbObservacionesHistoricas] (
  [IDObservacion] LONG(4),
  [ID] LONG(4),
  [Fecha] SHORT_DATE_TIME(8),
  [Observacion] MEMO,
  [Tipo] TEXT(255)
);
-- Rows: 365

-- LINKED/SKIPPED: [TbSolicitudes] (given file does not exist: C:\00repos\datos\Solicitudes_HPS_datos.accdb)
-- LINKED/SKIPPED: [TbSolicitudesFechas] (given file does not exist: C:\00repos\datos\Solicitudes_HPS_datos.accdb)
-- LINKED/SKIPPED: [TbSuministradores] (given file does not exist: C:\00repos\datos\Expedientes_datos.accdb)
CREATE TABLE [TbUsuarioAnexos] (
  [IDAnexo] LONG(4),
  [NombreAnexo] TEXT(255),
  [IDUsuario] LONG(4),
  [Hash] TEXT(64)
);
-- Rows: 385

CREATE TABLE [TbUsuarios] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [Correo_e] TEXT(255),
  [IDExpediente] LONG(4),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(2),
  [Requiere_Curso] TEXT(2),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [F_Baja] SHORT_DATE_TIME(8),
  [LugarPrestacionServicio] MEMO,
  [IDEmpresaUsuario] LONG(4),
  [IDEmpresaHPS] LONG(4),
  [IDJuridicaContrato] LONG(4),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(2),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [IDSolicitud] LONG(4),
  [CadenaContratistas] TEXT(255)
);
-- Rows: 368

-- LINKED/SKIPPED: [TbUsuariosAplicaciones] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
-- LINKED/SKIPPED: [TbUsuariosAplicacionesPermisos] (given file does not exist: C:\00repos\datos\Lanzadera_Datos.accdb)
CREATE TABLE [TbUsuariosEntidades] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [EmpresaUsuario] TEXT(255),
  [EmpresaTramitadora] TEXT(255),
  [CadenaContratistas] TEXT(255),
  [Correo_e] TEXT(255),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [IDExpediente] LONG(4),
  [CodExp] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(2),
  [F_Baja] SHORT_DATE_TIME(8),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(2),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [Curso_Realizado] TEXT(2),
  [Requiere_Curso] TEXT(2),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [Requiere_PrimeraConvocatoriaCurso] TEXT(2),
  [Requiere_SegundaConvocatoriaCurso] TEXT(2),
  [Requiere_CorreoJefeSeguridadCurso] TEXT(2),
  [Observaciones] MEMO,
  [HPS_NAC_SIN_DATOS] TEXT(2),
  [HPS_NAC_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_NAC_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_NAC_Grado] TEXT(255),
  [HPS_NAC_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_NAC_Especialidad] TEXT(255),
  [HPS_NAC_Renovacion] TEXT(2),
  [HPS_NAC_F_Baja] SHORT_DATE_TIME(8),
  [HPS_NAC_Activo] TEXT(2),
  [HPS_NAC_ApuntoDeCaducar] TEXT(2),
  [HPS_NAC_Baja] TEXT(2),
  [HPS_NAC_Caducado] TEXT(2),
  [HPS_NAC_PendienteRenovacion] TEXT(2),
  [HPS_NAC_Solicitado] TEXT(2),
  [HPS_NAC_ESTADO] TEXT(255),
  [HPS_NAC_Irregular] TEXT(2),
  [HPS_NAC_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_NAC_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_OTAN_SIN_DATOS] TEXT(2),
  [HPS_OTAN_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_OTAN_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_OTAN_Grado] TEXT(255),
  [HPS_OTAN_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_OTAN_Especialidad] TEXT(255),
  [HPS_OTAN_Renovacion] TEXT(2),
  [HPS_OTAN_F_Baja] SHORT_DATE_TIME(8),
  [HPS_OTAN_Activo] TEXT(2),
  [HPS_OTAN_ApuntoDeCaducar] TEXT(2),
  [HPS_OTAN_Baja] TEXT(2),
  [HPS_OTAN_Caducado] TEXT(2),
  [HPS_OTAN_PendienteRenovacion] TEXT(2),
  [HPS_OTAN_Solicitado] TEXT(2),
  [HPS_OTAN_ESTADO] TEXT(255),
  [HPS_OTAN_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_OTAN_Irregular] TEXT(2),
  [HPS_OTAN_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_ESA_SIN_DATOS] TEXT(2),
  [HPS_ESA_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_ESA_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_ESA_Grado] TEXT(255),
  [HPS_ESA_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_ESA_Especialidad] TEXT(255),
  [HPS_ESA_Renovacion] TEXT(2),
  [HPS_ESA_F_Baja] SHORT_DATE_TIME(8),
  [HPS_ESA_Activo] TEXT(2),
  [HPS_ESA_ApuntoDeCaducar] TEXT(2),
  [HPS_ESA_Baja] TEXT(2),
  [HPS_ESA_Caducado] TEXT(2),
  [HPS_ESA_PendienteRenovacion] TEXT(2),
  [HPS_ESA_Solicitado] TEXT(2),
  [HPS_ESA_ESTADO] TEXT(255),
  [HPS_ESA_Irregular] TEXT(2),
  [HPS_ESA_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_ESA_MESES_PARA_RENOVAR] TEXT(255),
  [HPS_UE_SIN_DATOS] TEXT(2),
  [HPS_UE_F_Concesion] SHORT_DATE_TIME(8),
  [HPS_UE_F_Caducidad] SHORT_DATE_TIME(8),
  [HPS_UE_Grado] TEXT(255),
  [HPS_UE_F_Solicitud] SHORT_DATE_TIME(8),
  [HPS_UE_Especialidad] TEXT(255),
  [HPS_UE_Renovacion] TEXT(2),
  [HPS_UE_F_Baja] SHORT_DATE_TIME(8),
  [HPS_UE_Activo] TEXT(2),
  [HPS_UE_ApuntoDeCaducar] TEXT(2),
  [HPS_UE_Baja] TEXT(2),
  [HPS_UE_Caducado] TEXT(2),
  [HPS_UE_PendienteRenovacion] TEXT(2),
  [HPS_UE_Solicitado] TEXT(2),
  [HPS_UE_ESTADO] TEXT(255),
  [HPS_UE_Irregular] TEXT(2),
  [HPS_UE_MOTIVO_IRREGULAR] TEXT(255),
  [HPS_UE_MESES_PARA_RENOVAR] TEXT(255)
);
-- Rows: 363

CREATE TABLE [TbUsuariosHistoricos] (
  [ID] LONG(4),
  [DNI] TEXT(255),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [Telefono] TEXT(255),
  [Correo_e] TEXT(255),
  [IDExpediente] LONG(4),
  [F_Nacimiento] SHORT_DATE_TIME(8),
  [LugarNacimiento] TEXT(255),
  [Motivo_HPS] TEXT(255),
  [F_Curso] SHORT_DATE_TIME(8),
  [CursoEnVigor] TEXT(2),
  [Requiere_Curso] TEXT(2),
  [FechaPrimeraConvocatoria] SHORT_DATE_TIME(8),
  [FechaSegundaConvocatoria] SHORT_DATE_TIME(8),
  [FechaCorreoNoCurso] SHORT_DATE_TIME(8),
  [F_Baja] SHORT_DATE_TIME(8),
  [LugarPrestacionServicio] MEMO,
  [IDEmpresaUsuario] LONG(4),
  [IDEmpresaHPS] LONG(4),
  [IDJuridicaContrato] LONG(4),
  [FAvisoConcesion] SHORT_DATE_TIME(8),
  [RequiereComunicacionConcesion] TEXT(2),
  [FechaHPSConcesionMinima] SHORT_DATE_TIME(8),
  [CadenaContratistas] TEXT(255),
  [IDSolicitud] LONG(4)
);
-- Rows: 247

CREATE TABLE [TbUsuariosSICA] (
  [ID] TEXT(255),
  [IDHPS] LONG(4),
  [IDHPSHistorico] LONG(4),
  [Nombre] TEXT(255),
  [Apellido_1] TEXT(255),
  [Apellido_2] TEXT(255),
  [TramitacionExterna] TEXT(2),
  [UsuarioSICATSOL] TEXT(255),
  [UsuarioSICATDE] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [FechaModificacion] SHORT_DATE_TIME(8),
  [Correo_e] TEXT(255),
  [Observaciones] MEMO,
  [IDEmpresaTramitadora] LONG(4),
  [FAltaTSOL] SHORT_DATE_TIME(8),
  [FAltaTdE] SHORT_DATE_TIME(8),
  [FbajaTSOL] SHORT_DATE_TIME(8),
  [FBajaTdE] SHORT_DATE_TIME(8)
);
-- Rows: 50

-- Summary: 23 local tables, 8 linked/skipped
