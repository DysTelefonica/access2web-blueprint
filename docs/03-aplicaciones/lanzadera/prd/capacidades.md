# Capacidades de Lanzadera frente al estudio previo

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Issue:** #1074 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Clasifica las 15 capacidades del [estudio previo](../capabilities.md) como **añadida**, **modificada**, **retirada** o **sin cambios** respecto al snapshot del 2026-08-06, con evidencia, y las enlaza con las fichas de formulario y los hallazgos. Las capacidades del estudio no tienen identificador; aquí se numeran C-LAN-1 a C-LAN-15 solo para referirse a ellas.

## Método y límites

- La clasificación compara cada objeto de lógica con el mismo fichero del snapshot (`data/staging/lanzaderas/src`) tras quitar `Attribute VB_`, comentarios, líneas en blanco y diferencias de espacios y mayúsculas. **Resultado: los 59 ficheros de código (28 formularios, 17 módulos, 14 clases) son idénticos.** Por eso todas las capacidades salen «sin cambios en el código». Coincide con lo que ya decía el estudio previo (sin diferencias versionadas entre `staging` y `main`).
- «Sin cambios» es igualdad de texto: no prueba que la capacidad funcione, ni que el estudio previo la describiera bien, ni que el binario no versionado coincida. Las fichas contrastan el comportamiento y sus hallazgos se señalan en la última columna.
- La **disposición aprobada** es la de la sección «Disposiciones finales» del estudio previo (decisiones ya aprobadas por el propietario); este estudio no la modifica.
- Las métricas de uso («Activa», «Rara») son del estudio previo sobre el backend del snapshot y no se han revalidado.

## Clasificación

| Capacidad | Objetos de lógica | Estado frente al snapshot | Disposición aprobada | Fichas | Hallazgos |
|---|---|---|---|---|---|
| C-LAN-1 Autenticación y sesión | `Funciones Genreales.bas` (`Login`), `Variables Globales.bas` (`LoginCorrecto`), `Constructor.bas`, `Usuario.cls` | sin cambios | Preservar y modernizar (auditoría) y migrar hashes con rehash al primer login | [Login](formularios/Form_FormLogin.md) | H-LAN-1, 2 |
| C-LAN-2 Alta, baja y perfil de usuario | `Usuario.cls` | sin cambios | Preservar y modernizar | [UsuariosGestion](formularios/Form_FormUsuariosGestion.md), [UsuarioGeneral](formularios/Form_FormUsuarioGeneral.md), [UsuarioDatosGenerales](formularios/Form_FormUsuarioDatosGenerales.md), [Perfil](formularios/Form_FormPerfil.md) | H-LAN-14, 15 |
| C-LAN-3 Restablecimiento y cambio de contraseña | `Usuario.cls` (`ResetearPass`, histórico) | sin cambios | Preservar y migrar | [ObtenerContraseña](formularios/Form_FormObtenerContraseña.md), [UsuarioCambioPass](formularios/Form_FormUsuarioCambioPass.md) | H-LAN-3, 4, 16 |
| C-LAN-4 Catálogo de aplicaciones | `Aplicacion.cls`, `Constructor.bas` | sin cambios | Preservar y modernizar | [AplicacionesGestion](formularios/Form_FormAplicacionesGestion.md), [AplicacionGeneral](formularios/Form_FormAplicacionGeneral.md), [AplicacionDatosGenerales](formularios/Form_FormAplicacionDatosGenerales.md) | H-LAN-8, 9 |
| C-LAN-5 Visibilidad por ubicación y estado | `Aplicacion.cls`, `Funciones Genreales.bas` | sin cambios | Retirar la segmentación por ubicación; UAT moderno | [Aplicaciones](formularios/Form_FormAplicaciones.md), [AplicacionesOficina](formularios/Form_FormAplicacionesOficina.md), [AplicacionesFueraOficina](formularios/Form_FormAplicacionesFueraOficina.md), [AplicacionesPrueba](formularios/Form_FormAplicacionesPrueba.md) | H-LAN-13 |
| C-LAN-6 Lanzamiento y paso de identidad | `Funciones Genreales.bas` (`Lanzar`), `Aplicacion.cls` | sin cambios | Retirar el lanzador de escritorio; catálogo semántico en el menú web | las cuatro de C-LAN-5 | H-LAN-11, 12 |
| C-LAN-7 Asignación usuario-aplicación | `Usuario.cls`, `UsuarioAplicacionPermisos.cls` | sin cambios | Preservar y modernizar | [UsuarioPerfilAplicaciones](formularios/Form_FormUsuarioPerfilAplicaciones.md) | H-LAN-17 |
| C-LAN-8 Roles y permisos granulares | `UsuarioAplicacionPermisos.cls` | sin cambios | Preservar y modernizar; enforcement central | sin ficha de formulario propia (lógica) | — |
| C-LAN-9 Administrador global | `Usuario.cls` (`EsAdministradorCalculado`) | sin cambios | Preservar y modernizar | sin ficha propia; ver Login y menús | H-LAN-1 |
| C-LAN-10 Menús por perfil | `Menu.cls` | sin cambios | Sustituir por menú global con permisos | [MenuPrincipalAdmin](formularios/Form_FormMenuPrincipalAdmin.md), [MenuPrincipalCalidad](formularios/Form_FormMenuPrincipalCalidad.md), [MenuPrincipalUsuario](formularios/Form_FormMenuPrincipalUsuario.md) | H-LAN-5, 6, 7 |
| C-LAN-11 Contenido formativo | `Video.cls`, `Categoria.cls`, `Visionado.cls`, `Funciones Genreales.bas` | sin cambios | **Retirar** (el histórico queda archivado, no se migra) | [VideosGestion](formularios/Form_FormVideosGestion.md) y las otras seis fichas de vídeos | H-LAN-18 a 22 |
| C-LAN-12 Conexiones y auditoría | `Conexion.cls`, `AplicacionApertura.cls` | sin cambios | Preservar y modernizar; retirar la telemetría de SSID y ubicación | [UsuariosConectados](formularios/Form_FormUsuariosConectados.md) | — |
| C-LAN-13 Cola de correo | `Correo.cls` | sin cambios | Reemplazar el mecanismo (servicio de notificaciones) | sin ficha propia | — |
| C-LAN-14 Configuración de backend y entorno | `Variables Globales.bas` (`getdb`, `LeeConfiguracionLocal`), `Config_BackendHelper.bas` | sin cambios | Retirar como capacidad de administración (la web externaliza la configuración) | sin ficha propia; ver [coexistencia](coexistencia.md) | — |
| C-LAN-15 Tareas, batch y exportaciones | campos `ParaTareasProgramadas`, `DiaEnvioTareas` en `Usuario.cls` | sin cambios | Fusionar con el scheduler compartido | sin ficha propia | — |

**Resumen:** 15 de 15 capacidades «sin cambios en el código»; 0 añadidas, 0 modificadas, 0 retiradas del código. Las disposiciones aprobadas retiran o reemplazan el mecanismo de las capacidades C-LAN-5, 6, 11 y 14 y el correo (C-LAN-13); eso es una decisión de producto, no un cambio del legacy.

## Contraste entre el estudio previo y las fichas

El estudio previo es coherente con las fichas en lo esencial. Lo que las fichas añaden y el estudio no recogía:

- **Defectos de autenticación** ([H-LAN-1 a H-LAN-4](README.md#hallazgos-del-legacy)): administrador de Windows entra sin contraseña, el bloqueo no caduca, la contraseña de un solo uso se guarda en claro. El estudio solo documenta la intención («historial de contraseñas y bloqueo»).
- **`Lanzar` no comprueba permisos** y el botón de una aplicación concreta se habilita para cualquier usuario (H-LAN-12): el estudio ya advertía que la autorización vive en controles de interfaz; las fichas lo concretan.
- **Formularios huérfanos:** `FormEnConstruccion`, `Formulario1`, `FormVideosUsuarios`, `FormObtenerContraseña` y `FormUsuariosConectados` (H-LAN-10, 22): ningún fichero de código de `src` los abre; podrían abrirse desde el formulario de inicio o macros del binario, que no están exportados.
- **Lanzadera abre y modifica los `.accdb` de otras aplicaciones** y es propietaria de las tablas de identidad que vinculan cuatro aplicaciones más ([coexistencia](coexistencia.md#3-lo-específico-de-lanzadera-gestiona-los-frontends-de-las-demás-aplicaciones)).

## Capacidades de código sin entrada propia en el estudio

Candidatas a incorporar al estudio o a confirmar como subsumidas; son propuestas, no decisiones:

| Candidata | Evidencia | Entrada más cercana |
|---|---|---|
| Ayuda contextual por formulario (`AbrirArchivoAyuda`, `TbHerramientaDocAyuda`) | [Login](formularios/Form_FormLogin.md) | C-LAN-10 |
| Consulta de versión y actualización de los ejecutables de las aplicaciones (`getVersion`, `setVersiones`, copia local) | [AplicacionesOficina](formularios/Form_FormAplicacionesOficina.md) | C-LAN-6 |
| Registro diario de aperturas por aplicación (`TbAplicacionesAperturas`, `ActualizarAperturas`) | [Aplicaciones](formularios/Form_FormAplicaciones.md) | C-LAN-12 |

## Pendientes

- Confirmar la asociación capacidad-objeto de la tabla con quien conozca Lanzadera.
- Contrastar semánticamente, y no solo por texto, las capacidades de autenticación, permisos y lanzamiento.
- Decidir si las candidatas entran en el estudio.
