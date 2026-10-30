# Clases de usuario, permisos, suministradores y notificaciones

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/NotificacionServicio.cls` (506 líneas), `UsuarioServicio.cls` (149), `Usuario.cls` (92), `Suministrador.cls` (66), `Correo.cls` (60), `SuministradorServicio.cls` (59), `UsuarioAplicacionPermisos.cls` (55) · **PRD:** [README](../README.md)

Identidad, catálogos de suministro y notificación. Dos clases no aportan comportamiento (usuario, permisos, correo, suministrador); las otras dos hacen el trabajo: `UsuarioServicio` resuelve el rol a partir de los permisos y `NotificacionServicio` compone y registra la notificación que acompaña a cada cambio de estado. La notificación no se envía desde aquí: se **registra** en la base de correos, y su entrega es responsabilidad de otra pieza.

Abreviaturas: `NS`, `classes/NotificacionServicio.cls`; `US`, `classes/UsuarioServicio.cls`; `U`, `classes/Usuario.cls`; `UAP`, `classes/UsuarioAplicacionPermisos.cls`; `COR`, `classes/Correo.cls`; `SUM`, `classes/Suministrador.cls`; `SUMS`, `classes/SuministradorServicio.cls`; `CR`, `modules/CorreoRepositorio.bas`; `VG`, `modules/Variables Globales.bas`; `ENUM`, `modules/modEnumeradores.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `Usuario` | clase de entidad | `CorreoUsuario`, `Password`, `UsuarioRed`, `nombre`, `Matricula`, `FechaAlta`, `Activado`, fechas de contraseña y conexión, `Telefono`, `Movil` (`U:18-…`), `getPropiedad` y `SetPropiedad` (`U:84-92`) | Fila de usuario del catálogo de Lanzadera |
| `UsuarioAplicacionPermisos` | clase de entidad | `CorreoUsuario`, `IDAplicacion` y un indicador por perfil (`UAP:18-…`) | Permisos de un usuario en una aplicación |
| `Correo` | clase de entidad | originator, destinatarios, asunto, cuerpo, adjunto, marca de cuerpo HTML y fechas (`COR:16-27`) | Fila de `tbCorreosEnviados` |
| `Suministrador` | clase de entidad | campos del suministrador (`SUM:…`), `getPropiedad` y `SetPropiedad` (`SUM:58-66`) | Fila del catálogo de suministros |
| `UsuarioServicio` | clase de servicio | cinco procedimientos (`US:17-149`) | Resolver usuario, rol y listas de responsables |
| `SuministradorServicio` | clase de servicio | `getSuministradores`, `getSuministradorPorID` (`SUMS:18-59`) | Buscar y leer suministros |
| `NotificacionServicio` | clase de servicio | siete procedimientos públicos y seis privados (`NS:14-506`) | Componer y registrar la notificación de un cambio de estado |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `US` | `getUsuarioConPermisos` | Devuelve el usuario con sus permisos cargados | `UsuarioRepositorio`, `AplicacionRepositorio` | Lectura de Lanzadera | `US:17-51` |
| `US` | `getUsuarioConectadoConPermisos` | Devuelve el usuario de la sesión con sus permisos | `UsuarioRepositorio`, `getUsuarioConectadoPorMaquina` | Lectura | `US:52-77` |
| `US` | `DeterminarRol` | Traduce los indicadores de permiso a un rol del enumerador, con jerarquía administrador, calidad y técnico | — | Sin escritura | `US:78-111`; objeto nulo en `US:81-83`, administrador en `US:86-89` |
| `US` | `getResponsablesTecnicos`, `getResponsablesCalidad` | Devuelven los usuarios que pueden asumir cada responsabilidad | repositorios de usuario | Lectura | `US:112-149` |
| `SUMS` | `getSuministradores`, `getSuministradorPorID` | Buscan y leen suministros | `SuministradorRepositorio` | Lectura | `SUMS:18-59` |
| `NS` | `EnviarNotificacion` | Compone la notificación de un cambio de estado y la registra | `GenerarCuerpoHTML`, `CR.GuardarCorreo` | `INSERT` en `tbCorreosEnviados` | `NS:14-77`; validación previa en `NS:23-24`, clave en `NS:44`, originator en `NS:55-62`, registro en `NS:66` |
| `NS` | `EnviarNotificacionConDetalle` | Igual, con un bloque HTML adicional | `GenerarCuerpoHTML`, `CR.GuardarCorreo` | `INSERT` en `tbCorreosEnviados` | `NS:78-135` |
| `NS` | `EnviarNotificacionEliminacion` | Compone la notificación de una eliminación | `GenerarCuerpoHTML_Eliminacion`, `CR.GuardarCorreo` | `INSERT` en `tbCorreosEnviados` | `NS:136-194` |
| `NS` | `GenerarCuerpoHTML`, `GenerarCuerpoHTML_Eliminacion`, `GenerarTarjeta`, `GenerarTarjetaAlerta`, `GenerarChecklistSiguienteFase` | Construyen el cuerpo del correo con tarjetas y la lista de comprobación de la fase siguiente | `ChecklistHelper` | Sin escritura | `NS:195-392` |
| `NS` | `ObtenerDestinatariosPorEstado` | Resuelve a quién se avisa según el estado | — | Sin escritura | `NS:393-443` |
| `NS` | `ObtenerProximosPasosPorEstado` | Devuelve los próximos pasos según estado y tipo | `ChecklistHelper` | Sin escritura | `NS:444-481`; cierre en `NS:470`, aprobación en `NS:472-473`, resto en `NS:475-477` |
| `NS` | `ValidarCorreo` | Aplica las reglas mínimas del correo | — | Sin escritura | `NS:482-506` |
| `U`, `UAP`, `COR`, `SUM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre | — | Sin efectos | `U:84-92`, `UAP:47-55`, `COR:53-60`, `SUM:58-66` |

## Flujos de trabajo

- **Normal — cambio de estado:** el flujo compone el modelo de vista, `EnviarNotificacion` valida (`NS:23-24`), genera el cuerpo (`NS:242-321`), decide el originator según haya suplantación (`NS:55-57`), lo acota a 50 caracteres (`NS:59-60`), obtiene la clave del correo (`NS:44`) y lo registra en la base de correos (`NS:66`).
- **Normal — rol:** `DeterminarRol` aplica una jerarquía explícita, con administrador por encima de calidad y de técnico (`US:85-89`).
- **Alternativos:** hay tres caminos de notificación —cambio de estado, con detalle y eliminación (`NS:14`, `NS:78`, `NS:136`)— que comparten la construcción del cuerpo pero difieren en la tarjeta que lo precede.
- **Errores:** un `ViewModel` nulo se rechaza con el error 513 (`NS:23-24`); `DeterminarRol` hace lo propio con un usuario nulo (`US:81-83`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Los permisos se traducen a un único rol con jerarquía fija | `DeterminarRol` | `US:85-89` |
| Bajo suplantación, el originator nombra a ambos usuarios | `EnviarNotificacion` | `NS:55-57` |
| El originator se acota a 50 caracteres | `EnviarNotificacion` | `NS:59-60` |
| La notificación se registra; el envío es de otra pieza | `CR.GuardarCorreo` | `NS:66` |
| Los próximos pasos dependen del estado y del tipo de solicitud | `ObtenerProximosPasosPorEstado` | `NS:444-481` |

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Usuarios y permisos | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Lectura en Lanzadera | Ninguna | `US:17-77` |
| Suministros | `TbSuministradores` | Lectura | Ninguna | `SUMS:18-59` |
| Notificaciones | `tbCorreosEnviados` | `Recordset` | Ninguna | `NS:44-66` |
| Cuerpo del correo | texto HTML en memoria | — | — | `NS:195-392` |

## Coexistencia

Tres puntos. El primero es la identidad: `Usuario` incluye un campo de contraseña (`U:19`) y los permisos son un conjunto de indicadores de tipo texto (`UAP:20-26`) resueltos contra el catálogo de Lanzadera. En el destino, la autenticación y la autorización tienen que salir de este modelo; la clase `Usuario` no debería cruzar el límite con la contraseña dentro. El segundo es la notificación: hoy `NotificacionServicio` no envía, compone y registra (`NS:66`), de modo que el envío vive en otra aplicación o en un servicio de correo externo; el destino debe decidir si reproduce esa separación o la unifica, y qué ocurre con los correos ya registrados. El tercero es el cuerpo HTML, que está construido con cadenas en el código (`NS:195-392`), incluidas tarjetas de alerta y la lista de comprobación de la fase siguiente: ese HTML es contenido y debería salir del código.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Cuatro entidades sin comportamiento, un servicio que resuelve identidad y roles contra el catálogo de Lanzadera, y un servicio que compone notificaciones en HTML y las registra en la base de correos.

**Posibles defectos, a verificar.**
1. La entidad `Usuario` transporta la contraseña como campo público (`U:19`): viaja con el objeto por toda la aplicación.
2. Los permisos son cadenas de «Sí» o «No» por perfil (`UAP:20-26`), no banderas.
3. `EnviarNotificacion` registra el correo aunque el envío posterior falle, y no guarda estado de entrega (`NS:66`).
4. El cuerpo del correo y los próximos pasos están codificados como literales en el servicio (`NS:195-392`, `NS:444-481`).

**Propuestas web pendientes de decisión.** Que la identidad y los permisos salgan del modelo de entidad y se resuelvan en el servidor. Que la notificación sea un recurso con estado de entrega. Que el cuerpo del correo y la lista de comprobación sean contenido traducible y versionable.

## Criterios de aceptación

1. El rol se resuelve con la jerarquía vigente y sin ambigüedad para varios indicadores a la vez. Trazado a `US:85-89`.
2. La notificación de un cambio de estado registra originator, destinatarios, asunto y cuerpo. Trazado a `NS:14-77`.
3. Bajo suplantación, el originator identifica al usuario real y al activo. Trazado a `NS:55-57`.
4. Los próximos pasos del estado de cierre y del de aprobación son los esperados. Trazado a `NS:470-473`.
5. Un `ViewModel` nulo se rechaza con error de validación. Trazado a `NS:23-24`.

## Pendientes

- Confirmar contra el binario dónde se entrega realmente el correo registrado, porque el destino depende de esa pieza (`NS:66`).
- Confirmar si el campo `Password` de la entidad tiene algún consumidor vivo (`U:19`).
- Confirmar el catálogo de indicadores de `UsuarioAplicacionPermisos` completo, porque las hojas de formularios lo usan para habilitar controles (`UAP:20-26`).
- Confirmar si `ObtenerDestinatariosPorEstado` cubre los tres tipos de solicitud o solo uno (`NS:393-443`).