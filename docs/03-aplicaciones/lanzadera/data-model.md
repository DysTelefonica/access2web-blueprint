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

## Schemas detallados (segunda pasada, 30+ tablas obtenidas)

### `tbUsuarios` (27 columnas, 156 filas) — **cabecera de usuario con datos personales**

Documentado en Lote 1 + Lote 8 (inventario backend). Hallazgos:
- `Matricula_DNI` (Text 50) — **DNI explícito** en columna.
- `DirCorreo`, `telfijo`, `telmovil` — datos personales.
- `SeLogean`, `ParaTareasProgramadas`, `Autorizador` son `YesNo` (type 1) en `tbUsuarios` (consistente con booleanos reales).
- `EmplazamientoExterno`, `UsuarioDeGestionRiesgos`, `UsuariosI3D` son `Text 2` (inconsistencia — booleanos como texto).
- `DiaEnvioTareas` (Integer 1) — sistema de tareas programadas.

### `TbAplicaciones` (22 columnas, 20 filas presumidas) — **catálogo de aplicaciones con `Pass` ⚠️**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `NombreAplicacion` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `NombreCorto` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `NombreEjecutable` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreArchivoDatos` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| **`Pass`** | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **contraseña de la app en texto plano ⚠️** |
| `NombreCarpeta` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreFuncionPublicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreCarpetaTemporal` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `TituloAplicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreIconoParaArbol` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreIcono` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreIconoLanzadera` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `EjecucionEnOficina` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `NombreCarpetaDocumentacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreDirectorioIconos` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreDirectorioAyuda` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreDirectorioRecursos` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `URLDIrectorioIconoAplicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `EnPruebas` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `ConIconoEnLanzadera` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `Comando` | 12 (Memo) | 0 | false | `TEXT NULL` — comando de Shell |

⚠️ **`Pass` es la contraseña de la app en texto plano**. Migración: secret manager (HR-3).

### `TbUsuariosAplicaciones` (28 columnas, 622 filas) — **usuarios con `Password` en texto plano ⚠️⚠️⚠️**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `CorreoUsuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK |
| **`Password`** | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **contraseña en texto plano ⚠️⚠️⚠️** |
| `UsuarioRed` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Matricula` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaAlta` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `Activado` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT TRUE` |
| `FechaProximoCambioContrasenia` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaUltimaConexion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `TieneQueCambiarLaContrasenia` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `Telefono` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — dato personal |
| `Movil` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — dato personal |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `UsuarioImborrable` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `EsAdministrador` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `PermisosAsignados` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `FechaBaja` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `PasswordNuncaCaduca` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `MantenerLanzaderaAbierta` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `PassIncialPlana` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **contraseña inicial plana ⚠️** |
| `UsuarioSSID` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Id` | 3 (Integer) | 2 | true | `SMALLINT NOT NULL` — PK interna |
| `JefeDelUsuario` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `PermisoPruebas` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `ParaTareasProgramadas` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `FechaBloqueo` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

⚠️⚠️⚠️ **`Password` en texto plano**: **622 usuarios con contraseñas en texto plano**. Sistema de caducidad de contraseñas (`FechaProximoCambioContrasenia`, `TieneQueCambiarLaContrasenia`, `PasswordNuncaCaduca`).

⚠️ **D109 propuesta**: **migración de contraseñas con hash + salt** (bcrypt o argon2). Eliminar `Password` legacy después de la migración. Conservar `TbUsuariosHistoricoContrasenias` como log de eventos de cambio (sin contraseñas).

### `TbUsuariosAplicacionesPermisos` (9 columnas, 622 filas) — **permisos por aplicación**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `CorreoUsuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — PK parte 1 |
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 2 |
| `EsUsuarioAdministrador` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioCalidad` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioEconomia` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioSecretaria` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioTecnico` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioSinAcceso` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `EsUsuarioCalidadAvisos` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

⚠️ **7 booleanos como Text(2)** (D102 cross-cutting). Migración: `BOOLEAN` en PostgreSQL.

### `TbConfiguracion` (5 columnas, 1 fila presumida)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `Clave` | 10 (Text) | 50 | true | `VARCHAR(50) NOT NULL` PK |
| `Valor` | 10 (Text) | 255 | false | `TEXT NULL` |
| `TipoDato` | 10 (Text) | 20 | false | `VARCHAR(20) NULL` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `FechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

### `TbConexiones` (6 columnas, volumen TBD)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `Usuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK |
| `UltimaConexion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UltimaDesconexion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `InstaladoFW3` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `InstaladoFW4` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `Exitoso` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |

### `TbAplicacionesAperturas` (15 columnas, volumen TBD)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDApertura` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK |
| `NombreUsuario` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaApertura` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `HoraApertura` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCierre` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `HoraCierre` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `NombreAplicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaEnvioCorreoAdministrador` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `EnOficina` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ Sí/No como texto (D102) |
| `UsuarioConectadoMaquina` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `VersionAplicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NombreMaquina` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `UsuarioMaquina` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |

⚠️ Datos de **telemetría de usuario** (nombre de máquina, usuario de máquina, ubicación). Consideraciones de privacidad.

### `TbAplicacionesEstados` (2 columnas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `PerfilAplicacion` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK ⚠️ **PK como Text, no Long** |
| `PerfilAplicacionEncriptado` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

⚠️ **PK como Text(255)** — no normalizada. Migración: agregar ID numérica como PK; mantener `PerfilAplicacion` como campo único.

### `TbAplicacionesParametros` (3 columnas, volumen TBD)

PK compuesta (`IDAplicacion`, `IDParametro`).

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 1 |
| `IDParametro` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 2 |
| `Valor` | 12 (Memo) | 0 | true | `TEXT NOT NULL` |

### `TbAplicacionesEdiciones` (6 columnas, volumen TBD) — **versionado de aplicaciones**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 1 |
| `IDVersion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 2 |
| `Version` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `FechaPublicacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `ParaInforme` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbAplicacionesEdicionesCambios` (5 columnas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDCambio` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDVersion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK a `TbAplicacionesEdiciones` |
| `Cambio` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `FechaCambio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `DescripcionCambio` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbAplicacionesPerfiles` (2 columnas)

PK compuesta (`IDAplicacion`, `Perfil`).

### `TbAplicacionesVideos` (9 columnas) — **asociación app-vídeo**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAplicacionVideo` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `IDVideo` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — FK a `TbVideos` |
| `IDAplicacion` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK a `TbAplicaciones` |
| `Descripcion` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `NombreArchivo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaCreacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UsuarioCrea` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `FechaModificacion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `UsuarioModifica` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbCategorias` (2 columnas)

Catálogo simple de categorías. PK + Nombre.

### `TbUbicaciones` (3 columnas) — **catálogo con flag SIRDEE**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `NombreUbicacion` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK |
| `Sirdee` | 10 (Text) | 2 | true | `VARCHAR(2) NOT NULL` — ⚠️ Sí/No como texto (D102) |
| `Ubicacion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbParametros` (2 columnas)

Catálogo de parámetros. PK + Parametro.

### `TbPermisos` (9 columnas) — **permisos por aplicación con F3-F9 (campos dinámicos)**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDAplicacion` | 4 (LongInteger) | 4 | true | `INTEGER NOT NULL` — PK parte 1 |
| `Usuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — PK parte 2 |
| `F3` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F4` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F5` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F6` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F7` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F8` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `F9` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

⚠️ **Patrón F3-F9 (campos dinámicos)**: ⚠️ **anti-patrón de modelado**. Migración: en PostgreSQL, modelar como JSONB o tabla hija con `nombre_campo` y `valor_campo`. Decidir cómo se usan estos campos en la nueva plataforma.

### `TbTablasAVincular` (4 columnas) — **config de tablas vinculadas por app**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDBBDD` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `IDAplicacion` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `NombreTabla` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `NombreTablaEnLocal` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbDetalleVersiones` (4 columnas)

PK compuesta (`IDAplicacion`, `IDVersion`, `IDDetalle`).

### `TbConexionesRegistro` (9 columnas)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDConexion` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Usuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `FechaConexion` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCierre` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `ConContraseña` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `UsuarioSSID` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `EnOficina` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `Vertical` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `Horizontal` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |

### `TbConexionUltimaAppAbierta` (2 columnas)

PK + IDUltimaAplicacionAbierta (FK a TbAplicaciones).

### `TbUsuarioAplicacionesSolicitud` (7 columnas) — **solicitudes de aplicación de usuario**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `CorreoUsuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK |
| **`Password`** | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — **contraseña en texto plano ⚠️⚠️⚠️** |
| `Nombre` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Matricula` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Telefono` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — dato personal |
| `Movil` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — dato personal |
| `FechaSolicitud` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

⚠️ **Otra `Password` en texto plano** + datos personales. D109.

### `TbUsuarioConfiguracion` (2 columnas)

PK `UsuarioDeRed` + `MantenerLanzaderaAbierta` (Text 2). ⚠️ D102.

### `TbUsuariosCorreosEnvio` (9 columnas) — **configuración de correos por usuario**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDCorreo` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Destinatarios` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `DestinatariosConCopia` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `DestinatariosConCopiaOculta` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Asunto` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Cuerpo` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `FechaEnvio` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `FechaCreado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `URLAdjunto` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |

### `TbUsuariosHistoricoContrasenias` (3 columnas, volumen TBD) — **histórico de contraseñas**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `Usuario` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` PK |
| **`PassAntigua`** | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` — **contraseña antigua en texto plano ⚠️** |
| `FechaPass` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |

⚠️ **Otra `PassAntigua` en texto plano** (D37 histórico). Migración: **eliminar contraseñas del histórico** (solo conservar el evento de cambio, no la contraseña). En PostgreSQL, dejar `PassAntigua = NULL` o eliminar la columna.

### `TbUsuariosTareasDiarias` (2 columnas) — **tareas diarias de usuario**

PK `UsuarioDeRed` + `MantenerLanzaderaAbierta` (Text 2). ⚠️ D102.

### `TbCuestionarios` (6 columnas) — **cuestionarios de formación**

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDCuestionario` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `FechaRealizado` | 8 (DateTime) | 8 | false | `TIMESTAMP NULL` |
| `IDUsuarioRealiza` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `IDAplicacion` | 4 (LongInteger) | 4 | false | `INTEGER NULL` — FK |
| `IDRespuestaCorrecta` | 4 (LongInteger) | 4 | false | `INTEGER NULL` |
| `Observaciones` | 12 (Memo) | 0 | false | `TEXT NULL` |

### `TbCuestionarioPreguntas` (volumen TBD) — **preguntas de cuestionarios**

### `TbCuestionaroRespuestas` (volumen TBD) — **respuestas de cuestionarios (con TYPO en nombre)**

### `TbVideos` (volumen TBD) — **vídeos de formación**

⚠️ **Pendiente**: 4 tablas no se pudieron obtener por typo en el path o tablas vacías. Se reintentarán en una iteración posterior.

## Hallazgos críticos del esquema de Lanzadera

1. **D109 (CRÍTICO)**: `Password` en `tbUsuariosAplicaciones`, `TbUsuarioAplicacionesSolicitud` y `PassAntigua` en `TbUsuariosHistoricoContrasenias` son **contraseñas en texto plano**. **622 usuarios con contraseñas expuestas**. La nueva plataforma debe migrar a **bcrypt/argon2 con salt** y eliminar las contraseñas del histórico.

2. **D109 también**: `Pass` en `TbAplicaciones` (contraseña de la app) — migrar a secret manager.

3. **D102 (cross-cutting)**: **20+ columnas como `Text 2)` (Sí/No como texto)**. La nueva plataforma debe estandarizar a `BOOLEAN`.

4. **D110**: `TbPermisos` con campos dinámicos `F3-F9` (anti-patrón). Migración: JSONB o tabla hija con `nombre_campo` y `valor_campo`.

5. **D111**: `TbAplicacionesEstados` con PK como `Text(255)`. Migración: agregar ID numérica como PK.

6. **D112**: `TbConexiones` y `TbConexionesRegistro` con datos de telemetría (máquina, SSID). Consideraciones de privacidad.

## Joins inferidos, no FKs demostradas

El código y los esquemas sugieren joins entre `TbVideos.IDAplicacion` y `TbAplicaciones.IDAplicacion`, `TbVideosVisionados.IDVideo` y `TbVideos.IDVideo`, y `TbVideosVisionados.IDUsuario` con alguna identidad de usuario. Dysflow no devolvió esas relaciones físicas en esta ejecución: se etiquetan **Likely** y requieren índices/PK y cardinalidades confirmadas.
