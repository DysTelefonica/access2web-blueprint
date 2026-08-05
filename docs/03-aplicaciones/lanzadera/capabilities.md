# Lanzadera · Capacidades descubiertas

Fecha de evidencia: **2026-08-04**. `staging` es la baseline funcional; `main` solo sirve como referencia publicada. La clasificación es semilla de debate, no disposición final.

## Baseline y delta staging/main

- `staging`: `63ba5e01617fdda857503d43151f06bb7bc11829` (`staging`), con índice CodeGraph-VBA propio.
- `main`: `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` (`main`), referencia publicada con índice propio.
- Delta funcional observable: no hay diferencias versionadas en formas, módulos, clases, consultas o informes entre ambos commits; la divergencia de `staging` es documental/operativa (`AGENTS.md`). Esto no demuestra equivalencia del binario no versionado.

| Capacidad | Evidencia actual | Clasificación de uso | Pregunta para debatir |
|---|---|---|
| Autenticación y sesión | `Form_FormLogin.ComandoIniciarSesion_Click` → `Constructor.getUsuario` → `Login` → `LoginCorrecto`; `m_ObjUsuarioConectadoLogin` | Activa: 10.360 conexiones; última observada 2024-02-02 | Mantener, modernizar o reemplazar mecanismo |
| Alta, baja y perfil de usuario | `Form_FormUsuariosGestion`, `Form_FormUsuarioDatosGenerales`, clase `Usuario` | Unknown: no métrica de operaciones agregada inspeccionada | Debatir mantener, modernizar o retirar |
| Restablecimiento y cambio de contraseña | `Form_FormObtenerContraseña`, `Form_FormUsuarioCambioPass`, `Usuario.ResetearPass` | Rara/indirecta: 455 entradas históricas; última 2026-05-11 | Debatir mantener, modernizar o reemplazar mecanismo |
| Catálogo de aplicaciones | `Aplicacion`, `Constructor.getAplicaciones`, `TbAplicaciones` | Activa: 20 entradas; 11 con aperturas observadas | Mantener contrato, fusionar o retirar entradas |
| Visibilidad por ubicación y estado | `Form_FormAplicacionesOficina/FueraOficina`; `EjecucionEnOficina`, `EnPruebas`, `ConIconoEnLanzadera` | Unknown: flags de catálogo, sin uso directo medido | Debatir mantener, modernizar o retirar |
| Lanzamiento y paso de identidad | `Funciones Genreales.Lanzar`; `Aplicacion.Abrir`; `Shell`, `/cmd`, `CorreoUsuario` | Activa: 11 aplicaciones con 1.737 aperturas; última 2024-12-10 | Mantener capacidad y reemplazar mecanismo |
| Asignación usuario-aplicación | `Usuario.ColAplicacionesPermisos`; `TbUsuariosAplicacionesPermisos` | Activa: 18 aplicaciones, 610 asignaciones | Fusionar como servicio compartido o modernizar localmente |
| Roles y permisos granulares | `UsuarioAplicacionPermisos` lee/escribe `TbUsuariosAplicacionesPermisos`; flags `Administrador`, `Calidad`, `Economía`, `Secretaría`, `Técnico`, `SinAcceso` y `CalidadAvisos` | Activa: 610 asignaciones agregadas | Debatir fusionar servicio o modernizar |
| Administrador global | `Usuario.EsAdministradorCalculado`; el menú/perfil usa el estado global y la colección de aplicaciones activas | Unknown: política estática, sin uso de administración medido | Debatir mantener y modernizar enforcement |
| Menús por perfil | `Menu`, `Form_FormMenuPrincipalAdmin/Calidad/Usuario` | Unknown: navegación estática no instrumentada | Debatir reemplazar mecanismo o mantener contrato |
| Contenido formativo | clases `Video`, `Categoria`, `Visionado`; forms `FormVideos*`; `TbVideos*` | Rara: 138 visionados; última 2023-12-05 | Mantener, modernizar, posponer o retirar |
| Conexiones y auditoría | `Conexion.RegistroConexion`; `TbConexiones*`; `AplicacionApertura` | Activa histórica: 10.360 conexiones; sin actividad posterior a 2024-02-02 en esta tabla | Fusionar, conservar para trazabilidad o reemplazar mecanismo |
| Cola de correo | `Correo.EnviarCorreo`; `TbUsuariosCorreosEnvio` y referencias a `TbCorreosEnviados` | Rara/stale: 1.042 registros; último creado 2023-04-27 | Fusionar, reemplazar mecanismo o retirar |
| Configuración de backend/entorno | `Variables Globales.getdb`, `LeeConfiguracionLocal`; `TbConfiguracionBackends` | Unknown: contrato técnico, sin telemetría de cambios | Debatir fusionar servicio o reemplazar mecanismo |
| Tareas, batch y exportaciones | campos `ParaTareasProgramadas`, `DiaEnvioTareas`, helpers de correo; no se probó ejecución | Unknown: 41/156 usuarios marcados para tareas; ejecución no demostrada | Mantener, modernizar, posponer o retirar |

**Ventana y límites:** se inspeccionaron fechas disponibles en el backend compartido `C:\00repos\datos\Lanzadera_Datos.accdb`; los máximos van de 2023-04-27 a 2024-12-10 según categoría, por lo que “Activa” significa uso observado en el histórico disponible, no actividad actual. “No observado” no se asigna sin una tabla/periodo fiable.

## Ledger de confianza

- **Verified-runtime:** existencia, rutas, inventario, esquemas y relaciones devueltos por Dysflow; consultas `SELECT` agregadas de uso por categorías. No significa test funcional superado.
- **Verified-static:** eventos, llamadas, navegación, reglas de permisos, SHA-256, `Shell`, manejo de errores y contratos observados en source.
- **Verified-static adicional (CodeGraph-VBA):** `Login` tiene un caller (`Form_FormLogin`); `getdb` tiene 76 callers internos; `UsuarioAplicacionPermisos` tiene 7 callers en clases y formularios. El grafo es per-repositorio: no prueba callers en otras aplicaciones.
- **Intended:** arquitectura declarada en OpenSpec (`Formulario → ViewModel → Servicio → Repositorio`) y documentación README sobre SSO, historial de contraseñas y bloqueo.
- **Divergent:** README afirma 12 clases y 8 módulos, pero el inventario actual devuelve 14 clases y 17 módulos estándar, incluyendo harness de tests; README también afirma autenticación SHA-256, que el código soporta, pero no prueba por sí sola la política efectiva de todos los registros.

## Riesgos de negocio observados

La autorización se materializa en controles de interfaz y helpers de dominio, no en un único enforcement server-side. El estado de administrador global se calcula desde `Usuario.EsAdministrador`; el grafo no demuestra por sí solo cómo se construye cada colección de aplicaciones. Expedientes se fuerza como accesible para usuarios en dos formularios aunque no exista fila de permiso. El ID 51 aparece como referencia estática/documental pendiente de reconciliar con el catálogo efectivo; no se atribuye todavía a una regla concreta.

## Disposiciones finales (post-Lote 1)

Decisiones posteriores a este descubrimiento que ya están **APROBADAS** y que cierran el debate abierto por capacidad. No se han medido nuevos datos: se aplica lo decidido en `docs/09-arquitectura-objetivo-y-principios.md` y en `08-decisiones-y-preguntas-abiertas.md`. Cada fila enlaza con su `topic_key` de Engram.

| Capacidad | Disposición | Estado | Origen |
|---|---|---|---|
| Formación (vídeos, cuestionarios, visionados, ActiveX) | **Retirar** del producto web; el histórico queda como evidencia archivada, no se migra al módulo operativo. | APROBADO | `product/lanzadera-training-disposition` |
| Lanzador Access (`Shell`, `/cmd`, copia/ejecutable, UNC) | **Retirar** mecanismo desktop; el catálogo semántico se conserva y se sustituye por entradas permission-aware en el menú global web. | APROBADO | `product/lanzadera-launcher-disposition` |
| Distinción oficina / fuera de oficina | **Retirar** segmentación por ubicación; el control pre-producción se reemplaza por un ciclo UAT moderno (ver `arch23905`). | APROBADO | `product/lanzadera-location-visibility` |
| Gestión de rutas de backend, contraseñas de backend, `TbConfiguracionBackends` y tablas vinculadas | **Retirar** como capacidad de administración; la configuración técnica se externaliza a adapters/ports hexagonales. | APROBADO | `product/lanzadera-backend-config-disposition` |
| Auditoría de autenticación y apertura de aplicación | **Preservar y modernizar** mediante logs canónicos y el subsistema de auditoría de la plataforma. | APROBADO | `product/lanzadera-audit-disposition` |
| Telemetría de SSID, ubicación física y coordenadas | **Retirar**; los datos de máquina/dispositivo también siguen minimización y solo se recogen si surge un requisito de seguridad. | APROBADO | `product/lanzadera-audit-disposition` |
| Catálogo de aplicaciones, registro de usuarios y asignaciones usuario-aplicación | **Preservar y modernizar** como núcleo de administración de plataforma. Los identificadores semánticos se conservan; el enforcement se rediseña centralmente. | APROBADO | `product/lanzadera-core-capabilities` |
| Cola de correo Lanzadera | **Reemplazar mecanismo**; pasa a ser el adaptador transitorio del servicio unificado de notificaciones (ver `informes-exportaciones-y-correo.md`). | APROBADO | `architecture/notification-delivery-adapter-v1` |
| Flags `ParaTareasProgramadas` y ejecución batch nocturna | **Fusionar** con el scheduler compartido; ya no son un flag de Lanzadera. | APROBADO | `architecture/shared-scheduler-operations` |
| Hashes heredados de contraseña | **Preservar y migrar**; verificación legacy inicial con rehash transparente al primer login exitoso. | APROBADO | `product/credential-migration` + `architecture/opportunistic-password-rehash` |

Las decisiones de ciclo de credencial (caducidad, bloqueo, umbrales, activación, notificación manual de cambios de permiso) viven en `docs/06-seguridad-y-trazabilidad.md` y `09-arquitectura-objetivo-y-principios.md`; las de UAT, registro de aplicaciones y suplantación también. Esta tabla registra **qué se hace con cada capacidad detectada**, no su diseño detallado.
