# 02 · Topología del ecosistema — Lanzadera (identidad y permisos)

## Propósito

Documenta el rol transversal de Lanzadera como origen de identidad, permisos y catálogo de aplicaciones. Sirve de referencia para todas las fichas de `03-aplicaciones/` y para `04-integraciones-y-operacion/`.

## Qué va aquí

- Hechos verificados sobre `TbAplicaciones`, `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`.
- Patrón de arranque, comando VBA y aperturas desde otras aplicaciones.
- Contrato de identidad: `UsuarioRed`, `CorreoUsuario`, `VBA.Command`, clases `UsuarioAplicacionPermisos`.
- Tabla de aplicaciones vigentes del catálogo (cuando se extraiga del Lote 1).

## Estado del contenido

**Lote 1 completado (2026-08-04).** El catálogo consultado contiene 19 filas y confirma IDs para Lanzadera, Expedientes, HPS, Solicitudes HPS, Condor, Gestion_Riesgos, No Conformidades y Brass. El detalle y los límites están en [`docs/03-aplicaciones/lanzadera/`](../../03-aplicaciones/lanzadera/README.md).

`TbUsuariosAplicacionesPermisos` usa la pareja `CorreoUsuario` + `IDAplicacion` y siete flags de rol. `Usuario.ColAplicacionesPermisos` carga la asignación; los formularios de aplicaciones convierten la asignación en botones habilitados/deshabilitados. Esto es evidencia estática/runtime de inventario, no una prueba de autorización exhaustiva.

## Disposiciones finales sobre Lanzadera (post-Lote 1)

Las decisiones APROBADAS que cierran el debate abierto por capacidad detectada están resumidas en [`03-aplicaciones/lanzadera/capabilities.md`](../../03-aplicaciones/lanzadera/capabilities.md#disposiciones-finales-post-lote-1). Esta subsección deja constancia del efecto sobre el rol transversal de Lanzadera como **origen de identidad y permisos**, no del detalle técnico.

| Mecanismo Lanzadera hoy | Disposición |
|---|---|
| Identidad, registro de usuarios y permisos por aplicación | **Preservar y modernizar** — núcleo de administración de plataforma. `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` se rediseñan; la web centraliza el enforcement. |
| Catálogo de aplicaciones (`TbAplicaciones`, IDs y comandos) | **Preservar contratos** — los IDs semánticos siguen siendo referencia de migración. |
| Auditoría de conexión y apertura (`TbConexiones*`, `TbAplicacionesAperturas`) | **Preservar y modernizar** como eventos de autenticación y acceso a módulo. |
| Login SHA-256 + tabla `Password` | **Preservar hashes en migración**; verificación legacy versionada y rehash transparente al primer login exitoso. Los campos con credenciales planas no se migran. |
| Lockout y recuperación | **Sustituir** por la política del producto objetivo (umbral configurable, default cinco intentos; duración una hora; desbloqueo por administrador global). |
| Activación y desactivación de usuarios, alta/baja de roles por aplicación | **Reservado al administrador global**. Los administradores de aplicación solo asignan usuarios activos existentes a roles existentes. |
| Cambio de permisos y notificación al usuario | **Sin correo automático**; el envío lo dispara explícitamente un administrador autorizado. |
| Suplantación de identidad para pruebas | **Solo administrador global**, con doble identidad auditada e indicación visible de sesión. |
| UAT y gobernanza del ciclo de pruebas | **Solo administrador global** define participantes, perfil y configuración de acceso al ciclo. Las excepciones de release deben quedar visibles en el historial de cambios para los usuarios. |
| Registro de nuevos módulos | **Híbrido**: el despliegue crea un registro "pendiente" con metadatos técnicos (id, versión, rutas, health, capacidades declaradas); la activación funcional y la configuración la decide un administrador global. |
| Lanzador Access (`Shell`, `/cmd`, copia/ejecutable, UNC) | **Retirar** — el menú global web permission-aware sustituye la ejecución desktop. |
| Distinción oficina / fuera de oficina | **Retirar** — el control pre-producción es ahora el ciclo UAT moderno. |
| Gestión técnica de backend, rutas y contraseñas de backend | **Retirar** como capacidad de administración; la configuración se externaliza a adapters. |
| Formación (vídeos, cuestionarios, visionados) | **Retirar** — el histórico queda como evidencia archivada. |
| Telemetría de SSID / ubicación física / coordenadas | **Retirar** — sigue minimización de datos. |
| Cola de correo, tasks flags, ejecución batch vía Lanzadera | **Fusionar** con el servicio unificado de notificaciones y el scheduler compartido. |

## Fuentes de autoridad

1. Documentación en `C:\00repos\documentacion\OPENSPEC\00_LANZADERA`.
2. Código en `C:\00repos\codigo\00_LANZADERA\00_main`.
3. Dysflow en modo solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Toda fila del catálogo debe llevar fuente, fecha y, si aplica, evidencia Dysflow.
- Los roles y permisos se citan con su tabla origen; no se infieren desde el nombre del rol.
- Las excepciones documentadas (por ejemplo, cuentas técnicas) se etiquetan como tales.

## Checklist

- [ ] Cada afirmación lleva etiqueta `Verified-*` / `Intended` / `Likely` / `Divergent`.
- [ ] Las referencias a tablas y procedimientos citan módulo y nombre exacto.
- [ ] No se mezclan permisos de entornos (producción, staging, sandbox) sin marcarlo.

## Siguiente paso

Poblar este documento al ejecutar el Lote 1 de `exploration.md`; detener y solicitar aprobación antes de pasar al Lote 2.
