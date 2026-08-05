# Lanzadera · Modelo de datos y evidencia ERD

## Identidad y configuración

El frontend contiene únicamente `TbConfiguracionBackends`. Su esquema incluye `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `PasswordBackend`, `IDAplicacion`, rutas PROD/LOCAL y `EnPruebas`. El backend contiene 35 tablas. La configuración actual consultada por Dysflow declara backend activo `PROD`, ID de aplicación `12` y rutas de producción/local; los valores sensibles no se reproducen.

## Entidades principales

| Entidad | Tabla(s) | Claves/campos observados |
|---|---|---|
| Usuario | `tbUsuarios`, `TbUsuariosAplicaciones` | `Id`; `CorreoUsuario`, `UsuarioRed`, estado, fechas, bloqueo |
| Aplicación | `TbAplicaciones` | `IDAplicacion`, nombre corto, ejecutable, backend, carpeta, ubicación, pruebas |
| Asignación y permisos | `TbUsuariosAplicacionesPermisos` | `CorreoUsuario`, `IDAplicacion`, siete flags de rol |
| Perfiles disponibles | `TbAplicacionesPerfiles` | `IDAplicacion`, `Perfil` |
| Apertura | `TbAplicacionesAperturas` | `IDApertura`, `IDAplicacion`, usuario, fechas, máquina, oficina, versión |
| Conexión | `TbConexiones`, `TbConexionesRegistro` | usuario, fechas, SSID, oficina, coordenadas y resultado |
| Credencial histórica | `TbUsuariosHistoricoContrasenias` | usuario, contraseña histórica, fecha |
| Correo | `TbUsuariosCorreosEnvio`; código también usa `TbCorreosEnviados` | destinatarios, asunto, cuerpo, fechas, adjunto |
| Formación | `TbVideos`, `TbCategorias`, `TbVideosCategorias`, `TbVideosVisionados`, `TbVideosCuestionario` | vídeo, aplicación, categoría, usuario y tiempo visionado |

## Relaciones físicas devueltas por Dysflow

El `get_relationships(target=backend)` del 2026-08-04 devolvió cinco relaciones de negocio:

| Tabla | Tabla relacionada | Campo ↔ campo |
|---|---|---|
| `TbParametros` | `TbAplicacionesParametros` | `IDParametro` ↔ `IDParametro` |
| `TbAplicaciones` | `TbAplicacionesParametros` | `IDAplicacion` ↔ `IDAplicacion` |
| `TbUsuariosAplicaciones` | `TbUsuariosHistoricoContrasenias` | `CorreoUsuario` ↔ `Usuario` |
| `TbAplicaciones` | `TbUsuariosAplicacionesPermisos` | `IDAplicacion` ↔ `IDAplicacion` |
| `TbUsuariosAplicaciones` | `TbUsuariosAplicacionesPermisos` | `CorreoUsuario` ↔ `CorreoUsuario` |

Además apareció una relación interna de tablas `MSysNavPane`, no de negocio. Estas son relaciones físicas observadas, no una afirmación de cardinalidad completa.

## Joins inferidos, no FKs demostradas

El código y los esquemas sugieren joins entre `TbVideos.IDAplicacion` y `TbAplicaciones.IDAplicacion`, `TbVideosVisionados.IDVideo` y `TbVideos.IDVideo`, y `TbVideosVisionados.IDUsuario` con alguna identidad de usuario. Dysflow no devolvió esas relaciones físicas en esta ejecución: se etiquetan **Likely** y requieren índices/PK y cardinalidades confirmadas.
