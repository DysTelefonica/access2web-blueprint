-- Schema dump: Lanzadera_Datos.accdb
-- Tables: 35

CREATE TABLE [Tb0HerramientaDocAyuda] (
  [NombreFormulario] TEXT(255),
  [NombreArchivoAyuda] TEXT(255)
);
-- Rows: 4

CREATE TABLE [TbAplicaciones] (
  [IDAplicacion] LONG(4),
  [NombreAplicacion] TEXT(255),
  [NombreCorto] TEXT(255),
  [NombreEjecutable] TEXT(255),
  [NombreArchivoDatos] TEXT(255),
  [Pass] TEXT(255),
  [NombreCarpeta] TEXT(255),
  [NombreFuncionPublicacion] TEXT(255),
  [NombreCarpetaTemporal] TEXT(255),
  [TituloAplicacion] TEXT(255),
  [NombreIconoParaArbol] TEXT(255),
  [NombreIcono] TEXT(255),
  [NombreIconoLanzadera] TEXT(255),
  [EjecucionEnOficina] TEXT(2),
  [NombreCarpetaDocumentacion] TEXT(255),
  [NombreDirectorioIconos] TEXT(255),
  [NombreDirectorioAyuda] TEXT(255),
  [NombreDirectorioRecursos] TEXT(255),
  [URLDIrectorioIconoAplicacion] TEXT(255),
  [EnPruebas] TEXT(2),
  [ConIconoEnLanzadera] TEXT(2),
  [Comando] MEMO
);
-- Rows: 20

CREATE TABLE [TbAplicacionesAperturas] (
  [IDApertura] LONG(4),
  [IDAplicacion] LONG(4),
  [NombreUsuario] TEXT(255),
  [FechaApertura] SHORT_DATE_TIME(8),
  [HoraApertura] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [HoraCierre] SHORT_DATE_TIME(8),
  [NombreAplicacion] TEXT(255),
  [FechaEnvioCorreoAdministrador] SHORT_DATE_TIME(8),
  [EnOficina] TEXT(2),
  [UsuarioConectadoMaquina] TEXT(255),
  [VersionAplicacion] TEXT(255),
  [NombreMaquina] TEXT(255),
  [UsuarioMaquina] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 1697

CREATE TABLE [TbAplicacionesEdiciones] (
  [IDAplicacion] LONG(4),
  [IDVersion] LONG(4),
  [Version] TEXT(255),
  [FechaPublicacion] SHORT_DATE_TIME(8),
  [ParaInforme] TEXT(255)
);
-- Rows: 1262

CREATE TABLE [TbAplicacionesEdicionesCambios] (
  [IDCambio] LONG(4),
  [IDVersion] LONG(4),
  [Cambio] TEXT(255),
  [FechaCambio] SHORT_DATE_TIME(8),
  [DescripcionCambio] MEMO
);
-- Rows: 1364

CREATE TABLE [TbAplicacionesEstados] (
  [PerfilAplicacion] TEXT(255),
  [PerfilAplicacionEncriptado] TEXT(255)
);
-- Rows: 6

CREATE TABLE [TbAplicacionesParametros] (
  [IDAplicacion] LONG(4),
  [IDParametro] LONG(4),
  [Valor] MEMO
);
-- Rows: 60

CREATE TABLE [TbAplicacionesPerfiles] (
  [IDAplicacion] LONG(4),
  [Perfil] TEXT(255)
);
-- Rows: 34

CREATE TABLE [TbAplicacionesVideos] (
  [IDAplicacionVideo] LONG(4),
  [IDVideo] LONG(4),
  [IDAplicacion] LONG(4),
  [Descripcion] MEMO,
  [NombreArchivo] TEXT(255),
  [FechaCreacion] SHORT_DATE_TIME(8),
  [UsuarioCrea] TEXT(255),
  [FechaModificacion] SHORT_DATE_TIME(8),
  [UsuarioModifica] TEXT(255)
);
-- Rows: 0

CREATE TABLE [TbCategorias] (
  [IDCategoria] LONG(4),
  [NombreCategoria] TEXT(255)
);
-- Rows: 2

CREATE TABLE [TbConexiones] (
  [Usuario] TEXT(255),
  [UltimaConexion] SHORT_DATE_TIME(8),
  [UltimaDesconexion] SHORT_DATE_TIME(8),
  [InstaladoFW3] TEXT(2),
  [InstaladoFW4] TEXT(2),
  [Exitoso] TEXT(2)
);
-- Rows: 24

CREATE TABLE [TbConexionesRegistro] (
  [IDConexion] LONG(4),
  [Usuario] TEXT(255),
  [FechaConexion] SHORT_DATE_TIME(8),
  [FechaCierre] SHORT_DATE_TIME(8),
  [ConContraseña] BOOLEAN(1),
  [UsuarioSSID] TEXT(255),
  [EnOficina] BOOLEAN(1),
  [Vertical] LONG(4),
  [Horizontal] LONG(4)
);
-- Rows: 10360

CREATE TABLE [TbConexionUltimaAppAbierta] (
  [IDConexion] LONG(4),
  [IDUltimaAplicacionAbierta] LONG(4)
);
-- Rows: 1469

CREATE TABLE [TbConfiguracion] (
  [Clave] TEXT(50),
  [Valor] TEXT(255),
  [TipoDato] TEXT(20),
  [Descripcion] MEMO,
  [FechaModificacion] SHORT_DATE_TIME(8)
);
-- Rows: 1

CREATE TABLE [TbCuestionarioPreguntas] (
  [IDPregunta] LONG(4),
  [IDCuestionario] LONG(4),
  [Texto] MEMO
);
-- Rows: 0

CREATE TABLE [TbCuestionarios] (
  [IDCuestionario] LONG(4),
  [FechaRealizado] SHORT_DATE_TIME(8),
  [IDUsuarioRealiza] LONG(4),
  [IDAplicacion] LONG(4),
  [IDRespuestaCorrecta] LONG(4),
  [Observaciones] MEMO
);
-- Rows: 0

CREATE TABLE [TbCuestionaroRespuestas] (
  [IDRespuesta] LONG(4),
  [IDPregunta] LONG(4),
  [Letra] TEXT(255),
  [Texto] MEMO
);
-- Rows: 0

CREATE TABLE [TbDetalleVersiones] (
  [IDAplicacion] LONG(4),
  [IDVersion] LONG(4),
  [IDDetalle] LONG(4),
  [Detalle] MEMO
);
-- Rows: 28

CREATE TABLE [TbParametros] (
  [IDParametro] LONG(4),
  [Parametro] TEXT(255)
);
-- Rows: 53

CREATE TABLE [TbPermisos] (
  [IDAplicacion] LONG(4),
  [Usuario] TEXT(255),
  [F3] TEXT(255),
  [F4] TEXT(255),
  [F5] TEXT(255),
  [F6] TEXT(255),
  [F7] TEXT(255),
  [F8] TEXT(255),
  [F9] TEXT(255)
);
-- Rows: 205

CREATE TABLE [TbTablasAVincular] (
  [IDBBDD] LONG(4),
  [IDAplicacion] LONG(4),
  [NombreTabla] TEXT(255),
  [NombreTablaEnLocal] TEXT(255)
);
-- Rows: 411

CREATE TABLE [TbUbicaciones] (
  [NombreUbicacion] TEXT(255),
  [Sirdee] TEXT(2),
  [Ubicacion] TEXT(255)
);
-- Rows: 8

CREATE TABLE [TbUsuarioAplicacionesSolicitud] (
  [CorreoUsuario] TEXT(255),
  [Password] TEXT(255),
  [Nombre] TEXT(255),
  [Matricula] TEXT(255),
  [Telefono] TEXT(255),
  [Movil] TEXT(255),
  [FechaSolicitud] SHORT_DATE_TIME(8)
);
-- Rows: 0

CREATE TABLE [TbUsuarioConfiguracion] (
  [UsuarioDeRed] TEXT(255),
  [MantenerLanzaderaAbierta] TEXT(2)
);
-- Rows: 14

CREATE TABLE [tbUsuarios] (
  [Id] LONG(4),
  [Nombre] TEXT(50),
  [UsuarioRed] TEXT(50),
  [DirCorreo] TEXT(255),
  [Matricula_DNI] TEXT(50),
  [Cargo] TEXT(50),
  [telfijo] LONG(4),
  [telmovil] LONG(4),
  [JefeDelUsuario] TEXT(50),
  [FechaAlta] SHORT_DATE_TIME(8),
  [FechaBaja] SHORT_DATE_TIME(8),
  [EmplazamientoExterno] TEXT(2),
  [SeLogean] BOOLEAN(1),
  [ParaTareasProgramadas] BOOLEAN(1),
  [Autorizador] BOOLEAN(1),
  [DiaEnvioTareas] BYTE(1),
  [UsuarioDeGestionRiesgos] TEXT(2),
  [UsuariosI3D] TEXT(2)
);
-- Rows: 156

CREATE TABLE [TbUsuariosAplicaciones] (
  [CorreoUsuario] TEXT(255),
  [Password] TEXT(255),
  [UsuarioRed] TEXT(255),
  [Nombre] TEXT(255),
  [Matricula] TEXT(255),
  [FechaAlta] SHORT_DATE_TIME(8),
  [Activado] BOOLEAN(1),
  [FechaProximoCambioContrasenia] SHORT_DATE_TIME(8),
  [FechaUltimaConexion] SHORT_DATE_TIME(8),
  [TieneQueCambiarLaContrasenia] BOOLEAN(1),
  [Telefono] TEXT(255),
  [Movil] TEXT(255),
  [Observaciones] MEMO,
  [UsuarioImborrable] BOOLEAN(1),
  [EsAdministrador] TEXT(2),
  [PermisosAsignados] BOOLEAN(1),
  [FechaBaja] SHORT_DATE_TIME(8),
  [PasswordNuncaCaduca] BOOLEAN(1),
  [MantenerLanzaderaAbierta] BOOLEAN(1),
  [PassIncialPlana] TEXT(255),
  [UsuarioSSID] TEXT(255),
  [Id] INT(2),
  [JefeDelUsuario] TEXT(50),
  [PermisoPruebas] TEXT(2),
  [ParaTareasProgramadas] BOOLEAN(1),
  [FechaBloqueo] SHORT_DATE_TIME(8)
);
-- Rows: 186

CREATE TABLE [TbUsuariosAplicacionesPermisos] (
  [CorreoUsuario] TEXT(255),
  [IDAplicacion] LONG(4),
  [EsUsuarioAdministrador] TEXT(2),
  [EsUsuarioCalidad] TEXT(2),
  [EsUsuarioEconomia] TEXT(2),
  [EsUsuarioSecretaria] TEXT(2),
  [EsUsuarioTecnico] TEXT(2),
  [EsUsuarioSinAcceso] TEXT(2),
  [EsUsuarioCalidadAvisos] TEXT(2)
);
-- Rows: 622

CREATE TABLE [TbUsuariosAplicacionesTareas] (
  [CorreoUsuario] TEXT(255),
  [EsAdministrador] TEXT(2),
  [EsTecnico] TEXT(2),
  [EsCalidad] TEXT(2),
  [EsEconomia] TEXT(2)
);
-- Rows: 7

CREATE TABLE [TbUsuariosCorreosEnvio] (
  [IDCorreo] LONG(4),
  [Destinatarios] MEMO,
  [DestinatariosConCopia] MEMO,
  [DestinatariosConCopiaOculta] MEMO,
  [Asunto] TEXT(255),
  [Cuerpo] MEMO,
  [FechaEnvio] SHORT_DATE_TIME(8),
  [FechaCreado] SHORT_DATE_TIME(8),
  [URLAdjunto] TEXT(255)
);
-- Rows: 1042

CREATE TABLE [TbUsuariosHistoricoContrasenias] (
  [Usuario] TEXT(255),
  [PassAntigua] TEXT(255),
  [FechaPass] SHORT_DATE_TIME(8)
);
-- Rows: 455

CREATE TABLE [TbUsuariosTareasDiarias] (
  [UsuarioRed] TEXT(255)
);
-- Rows: 8

CREATE TABLE [TbVideos] (
  [IDVideo] LONG(4),
  [Titulo] TEXT(255),
  [NombreArchivo] TEXT(255),
  [IDAplicacion] LONG(4),
  [Observaciones] MEMO,
  [Descripcion] MEMO,
  [SubidoPor] TEXT(255),
  [FechaSubido] SHORT_DATE_TIME(8),
  [ParaCalidad] TEXT(2),
  [ParaAdministrador] TEXT(2),
  [ParaTecnicos] TEXT(2)
);
-- Rows: 16

CREATE TABLE [TbVideosCategorias] (
  [IDCategoriaVideo] LONG(4),
  [IDCategoria] LONG(4),
  [IDVideo] LONG(4)
);
-- Rows: 15

CREATE TABLE [TbVideosCuestionario] (
  [IDVideoCuestionario] LONG(4),
  [IDCuestionario] LONG(4),
  [IDVideo] LONG(4),
  [FechaRealizado] SHORT_DATE_TIME(8),
  [UsuarioRealiza] TEXT(255),
  [Observaciones] MEMO
);
-- Rows: 0

CREATE TABLE [TbVideosVisionados] (
  [IDVisionado] LONG(4),
  [IDVideo] LONG(4),
  [TiempoVisionado] LONG(4),
  [TiempoVideo] LONG(4),
  [IDUsuario] LONG(4),
  [FechaVisionado] SHORT_DATE_TIME(8)
);
-- Rows: 138

-- Summary: 35 local tables, 0 linked/skipped
