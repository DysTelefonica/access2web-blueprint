# Matriz campo a campo — Lanzadera

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Issue:** #1074 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de origen de las tablas del backend de Lanzadera, con la evidencia de su uso en `src` y su relación con el esquema web que ya existe (`lanzadera.*`). Ninguna fila está aprobada: la topología PostgreSQL global y el modelo de identidad de la web siguen abiertos.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), extracción del backend `Lanzadera_Datos.accdb` del snapshot (35 tablas locales, 230 columnas). **No está revalidada contra el backend del commit fijado**: el ERD de origen queda pendiente del propietario (D-LAN-3 del README), porque `dysflow` no está disponible en el entorno de trabajo y el backend del clon no es legible.
- **Qué falta en esa fuente:** solo columna y tipo. No hay claves primarias, índices, relaciones, nulabilidad ni valores por defecto. El [estudio previo](../data-model.md) cita cinco relaciones físicas obtenidas con `dysflow` el 2026-08-04 (`TbParametros`-`TbAplicacionesParametros`, `TbAplicaciones`-`TbAplicacionesParametros`, `TbUsuariosAplicaciones`-`TbUsuariosHistoricoContrasenias` por `CorreoUsuario`/`Usuario`, `TbAplicaciones`-`TbUsuariosAplicacionesPermisos` y una más); son contexto, no evidencia de este estudio.
- **Evidencia en `src`:** se busca cada columna en los ficheros de producción (excluidos los módulos `Test*`) que citan la tabla. Es lectura estática.
- **Sin valores:** la matriz no contiene datos de ninguna tabla, solo estructura.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `INTEGER` | Identidad: el módulo web actual usa otros identificadores (UUID o `BIGINT`); la correspondencia de claves entre Access y web no está definida |
| `INT` | `SMALLINT` | — |
| `TEXT(n)` | `VARCHAR(n)` | `Sí`/`No` en texto: conservar hasta decidir `BOOLEAN` |
| `MEMO` | `TEXT` | — |
| `DOUBLE` | `DOUBLE PRECISION` | — |
| `SHORT_DATE_TIME` | `TIMESTAMP` | Fecha o fecha y hora; zona horaria |
| `BOOLEAN` | `BOOLEAN` | — |

## Clase de dato

La columna «Clase de dato» es una **heurística por nombre de columna**, sin revisar contenido: `credencial` (contraseñas e históricos), `personal` (nombre, correo, matrícula, teléfonos, usuario de red, ubicación o red del puesto) o «—». Sirve para priorizar la revisión de privacidad (RGPD) y de cifrado y no sustituye a una clasificación formal. Las credenciales se migran, si se migran, con tratamiento propio y nunca en claro; esa decisión es del propietario.

## Correspondencia con la web actual

El módulo [`app/src/modules/lanzadera`](../../../../app/src/modules/lanzadera) ya tiene un esquema `lanzadera` con `users`, `apps`, `profiles`, `user_app_assignments`, `global_admins`, `sessions`, `audit`, `reset_tokens`, `mail_outbox` (migración `0001`) y `presence_sessions` (migración `0008`). La columna «Correspondencia» solo indica la **tabla candidata**; el mapeo columna a columna llegará con el ERD de origen y con la decisión sobre el modelo de identidad. La web no tiene tabla para la mayor parte del legacy (aperturas, ediciones, parámetros, vídeos, cuestionarios): quedan «sin correspondencia».

## Leyenda de la columna «Evidencia en `src`»

- `F`: la columna aparece como `Fields("columna")` en un fichero que cita la tabla (primer hallazgo).
- `S`: el nombre aparece en un fichero que cita la tabla, sin `Fields(...)`; coincidencia de nombre, poco fiable para nombres comunes.
- `—`: sin cita en los ficheros de producción.

## Resumen por tabla

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` | Familia | Detalle |
|---|---|---|---|---|---|
| `Tb0HerramientaDocAyuda` | 2 | 4 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tb0herramientadocayuda) |
| `TbAplicaciones` | 22 | 20 | 3 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicaciones) |
| `TbAplicacionesAperturas` | 15 | 1697 | 2 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesaperturas) |
| `TbAplicacionesEdiciones` | 5 | 1262 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesediciones) |
| `TbAplicacionesEdicionesCambios` | 5 | 1364 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesedicionescambios) |
| `TbAplicacionesEstados` | 2 | 6 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesestados) |
| `TbAplicacionesParametros` | 3 | 60 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesparametros) |
| `TbAplicacionesPerfiles` | 2 | 34 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesperfiles) |
| `TbAplicacionesVideos` | 9 | 0 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbaplicacionesvideos) |
| `TbCategorias` | 2 | 2 | 2 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbcategorias) |
| `TbConexiones` | 6 | 24 | 0 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbconexiones) |
| `TbConexionesRegistro` | 9 | 10360 | 4 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbconexionesregistro) |
| `TbConexionUltimaAppAbierta` | 2 | 1469 | 1 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbconexionultimaappabierta) |
| `TbConfiguracion` | 5 | 1 | 1 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbconfiguracion) |
| `TbCuestionarioPreguntas` | 3 | 0 | 0 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbcuestionariopreguntas) |
| `TbCuestionarios` | 6 | 0 | 0 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbcuestionarios) |
| `TbCuestionaroRespuestas` | 4 | 0 | 0 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbcuestionarorespuestas) |
| `TbDetalleVersiones` | 4 | 28 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbdetalleversiones) |
| `TbParametros` | 2 | 53 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbparametros) |
| `TbPermisos` | 9 | 205 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbpermisos) |
| `TbTablasAVincular` | 4 | 411 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbtablasavincular) |
| `TbUbicaciones` | 3 | 8 | 0 | catálogo de aplicaciones y configuración | [catalogo](matriz-campos/catalogo.md#tbubicaciones) |
| `TbUsuarioAplicacionesSolicitud` | 7 | 0 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuarioaplicacionessolicitud) |
| `TbUsuarioConfiguracion` | 2 | 14 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuarioconfiguracion) |
| `tbUsuarios` | 18 | 156 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuarios) |
| `TbUsuariosAplicaciones` | 26 | 186 | 3 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuariosaplicaciones) |
| `TbUsuariosAplicacionesPermisos` | 9 | 622 | 3 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuariosaplicacionespermisos) |
| `TbUsuariosAplicacionesTareas` | 5 | 7 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuariosaplicacionestareas) |
| `TbUsuariosCorreosEnvio` | 9 | 1042 | 3 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbusuarioscorreosenvio) |
| `TbUsuariosHistoricoContrasenias` | 3 | 455 | 3 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuarioshistoricocontrasenias) |
| `TbUsuariosTareasDiarias` | 1 | 8 | 0 | identidad, permisos y credenciales | [identidad](matriz-campos/identidad.md#tbusuariostareasdiarias) |
| `TbVideos` | 11 | 16 | 2 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbvideos) |
| `TbVideosCategorias` | 3 | 15 | 2 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbvideoscategorias) |
| `TbVideosCuestionario` | 6 | 0 | 0 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbvideoscuestionario) |
| `TbVideosVisionados` | 6 | 138 | 1 | conexiones, correo y formación | [conexiones](matriz-campos/conexiones.md#tbvideosvisionados) |

## Tablas de Lanzadera que consumen otras aplicaciones

El `ERD/schema.sql` de otras cuatro aplicaciones registra vínculos a `Lanzadera_Datos.accdb`: Expedientes, HPS, Condor y No Conformidades vinculan `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`; Gestión de Riesgos vincula además `TbAplicaciones`. Lanzadera es, por tanto, propietaria de la identidad y de los permisos que usan al menos cinco de las ocho aplicaciones. Cualquier cambio de estructura o de claves de esas tres tablas afecta a las demás y debe coordinarse en el ERD global.

## Tandas de la matriz

Con esta tanda la matriz cubre las 35 tablas del backend: identidad, catálogo de aplicaciones y configuración, y conexiones, correo y formación. Lo que depende del ERD de origen (claves, índices, relaciones y nulabilidad) y la correspondencia columna a columna con el esquema web quedan pendientes; el ERD lo extrae el propietario con `dysflow` en su máquina (D-LAN-3).
