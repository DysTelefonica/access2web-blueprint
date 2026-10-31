# 07 · Migración — Matriz legacy → web

## Propósito

Matriz viva que asocia capacidades legacy con futuros componentes web. **No** define el stack objetivo ni la arquitectura destino: solo deja trazada la correspondencia para que una fase SDD posterior pueda descomponerla en `proposal/spec/design/tasks`.

## Qué va aquí

- Una fila por capacidad: `capacidad legacy · componente web candidato · riesgo principal · dependencias · excepciones · confianza`.
- Identificación de capacidades con dependencia crítica de UI Access (lógica a extraer a servicio/use-case).
- Identificación de capacidades candidatas a conservar como reporte o solo-lectura en web.
- Trazabilidad a la ficha de capacidad correspondiente en `05-capacidades/`.

## Estado del contenido

La matriz traduce a componente web candidato las capacidades de las **siete aplicaciones cuyo estudio tiene la fase 5 cerrada** en `main`: Expedientes (#1073), Lanzadera (#1074), Condor (#1075), Brass (#1078), Gestión de Riesgos (#1079), HPS Solicitudes (#1077) y No Conformidades (#1080). Cada fila resume y enlaza; la evidencia con cita `ruta:línea` vive en el PRD de la aplicación y no se reproduce aquí.

**Sigue en curso y queda fuera de esta versión:** HPS (#1076). Su estudio no ha entregado todavía la clasificación de capacidades, de modo que no hay filas que traducir. La excepción `Nemotecnico` de No Conformidades entra con la fila de alta de NC de proyecto (CAP-NC-7).

**Pendiente transversal:** el ERD de origen de las ocho aplicaciones sigue sin extraerse. La extracción de solo lectura con `dysflow-erd` depende del runner Windows del propietario y bloquea, entre otras cosas, las claves, los índices y las relaciones que la coexistencia necesita para afirmar nada sobre PostgreSQL vinculado (`D-EXP-4`, `D-CON-5`, `D-BRA-1`, `D-GR-4`). Ninguna fila de esta matriz depende de esa extracción.

## Fuentes de autoridad

1. `05-capacidades/`.
2. `07-migracion/modelo-dominio-agnostico.md`.
3. `03-aplicaciones/<app>/` por si una capacidad es exclusiva de una aplicación.
4. Dysflow solo lectura, como validación puntual.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- Un componente web candidato se nombra solo si hay capacidad legacy que lo justifique.
- Las capacidades con lógica en evento de formulario se marcan como `lógica a extraer` y se referencian a su módulo/handler.
- Las excepciones (`Nemotecnico` en No Conformidades, etc.) viajan con su fila, no se difuminan.

## Cómo se lee la matriz

- **Origen de cada fila.** El identificador y la descripción corta de cada capacidad son los que declara el PRD de la aplicación (`03-aplicaciones/<app>/prd/capacidades.md`). El resumen de los siete estudios está en [`05-capacidades/capabilities-index.md`](../05-capacidades/capabilities-index.md).
- **Componente web candidato.** Para Expedientes, el candidato es el requisito del spec vigente (`openspec/changes/expedientes-web-migration/specs/`) y el módulo web que ya existe en `app/src/modules/expedientes/`: la épica histórica de Expedientes está *superseded* como plan de producto por el ledger normativo y se conserva solo como inventario. Para el resto de aplicaciones, el candidato es la feature `F1..Fn` de la épica (`03-aplicaciones/<app>/epic.md`) o el módulo web ya escrito, cuando existe.
- **Riesgo principal.** Es el riesgo que registran el PRD, la épica o la coexistencia de esa capacidad, con su identificador (`H-LAN-*`, `D-*`, `P-*`). Cuando ninguna fuente registra un riesgo propio de esa capacidad, la celda va con `—`.
- **Dependencias.** Es la dependencia principal que el PRD sostiene. `—` cuando el PRD no declara ninguna.
- **Excepción.** Es la desviación, la ausencia o la disposición que el PRD registra. `—` cuando no hay ninguna.
- **Confianza.** Solo Expedientes tiene un vocabulario de confianza declarado: la columna «Evidencia» del [ledger normativo](../03-aplicaciones/expedientes/capability-ledger.md), con las marcas `VS` (evidencia estática auditada), `VR-hist` (evidencia runtime histórica, no repetida) e `INT` (contrato pretendido o integración por confirmar). Se transcriben como `Verified-static`, `Verified-runtime (histórico)` e `Intended`. Los PRD de Lanzadera, Condor, Brass, Gestión de Riesgos, HPS Solicitudes y No Conformidades no declaran nivel de confianza por capacidad, de modo que sus celdas van con `—`.
- **`sin candidato declarado`** significa que ninguna pieza del repositorio (épica, spec o módulo web) justifica hoy un candidato. No es una señal de que la capacidad se retire.

## Expedientes

Fuente: [`prd/capacidades.md`](../03-aplicaciones/expedientes/prd/capacidades.md) · commit fijado `ec6a8c3` · identificadores `EXP-CAP-001..063` · [ledger normativo](../03-aplicaciones/expedientes/capability-ledger.md).

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| EXP-CAP-001 · Alta transaccional del expediente | `specs/lifecycle.md` (EXP-CAP-001) · `application/create_expediente` | P-EXP-3: la vía de vínculos por ODBC sigue abierta; las operaciones de catálogos no son transaccionales | expediente | — | Verified-static |
| EXP-CAP-002 · Edición del agregado con invariantes y concurrencia | `specs/lifecycle.md` · `application/edit_expediente` | Sin versión de fila: el último que guarda gana (coexistencia, apartado 7) | expediente | — | Verified-static |
| EXP-CAP-003 · Baja condicionada y auditable | `specs/lifecycle.md` · `application/delete_expediente` | `Eliminar` toca 13 tablas y las reglas de «en uso» no tienen clave foránea que las respalde | expediente | — | Verified-static |
| EXP-CAP-004 · Cambio de tipo con validación de elegibilidad | `specs/lifecycle.md` | Elegibilidad y efectos derivados en `ExpedienteOperaciones.cls` | expediente | — | Verified-static |
| EXP-CAP-005 · Estado calculado y transiciones con fechas y garantía | `specs/lifecycle.md` · `application/transition_expediente` | `ESTADO` y `FechaFinGarantia` se recalculan en VBA (coexistencia, apartado 6) | expediente | — | Verified-static |
| EXP-CAP-006 · Jerarquía de acuerdos marco, lotes y basados | `specs/lifecycle.md` | Integridad padre-hijo sin clave foránea | expediente | — | Verified-static |
| EXP-CAP-007 · Derivación ordinal funcional | `specs/lifecycle.md` | — | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-008 · Hitos y sus reglas temporales | `specs/related-data.md` · `application/register_hito` | Los hitos no se guardan en transacción | expediente | — | Verified-static |
| EXP-CAP-009 · Modificados, cambios de campo e historial | `specs/related-data.md` · `application/register_modificado` | `RegistrarUltimoCambio` escribe fuera de la transacción de guardado | expediente | — | Verified-static |
| EXP-CAP-010 · Anexos y referencias documentales | `specs/related-data.md` · `application/anexos` | Retención y borrado controlados; el puerto `ports/document_storage.py` no tiene adaptador declarado | expediente | — | Verified-static |
| EXP-CAP-011 · Anualidades vinculadas al expediente | `specs/related-data.md` · `application/register_anualidad` | Escritura sin transacción | expediente | — | Verified-static |
| EXP-CAP-012 · Responsables por rol y jefaturas | `specs/related-data.md` · `application/register_responsable` | El invariante de responsables solo lo aplica VBA (coexistencia, apartado 6) | expediente | — | Verified-static |
| EXP-CAP-013 · Relaciones con entidades y jurídicas | `specs/related-data.md` · `application/register_juridica` | Unicidad por consulta previa, sin índice | expediente | — | Verified-static |
| EXP-CAP-014 · Suministradores, UTE y jerarquía del árbol | `specs/related-data.md` · `application/register_suministrador` | La jerarquía vive en `ExpedienteSuministradorServicio.cls` | expediente | — | Verified-static |
| EXP-CAP-015 · Comerciales | `specs/catalogs.md` · `application/catalogs` | Unicidad por consulta previa; el registro «N/A» se protege por nombre literal | — | — | Verified-static |
| EXP-CAP-016 · CPV | `specs/catalogs.md` · `domain/cpv` | La comprobación de «en uso» no tiene clave foránea | — | — | Verified-static |
| EXP-CAP-017 · Ejércitos | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-018 · Suministradores (catálogo) | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-019 · Lugares de ejecución | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-runtime (histórico) + Verified-static |
| EXP-CAP-020 · PECAL | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-021 · RAC | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-runtime (histórico) + Verified-static |
| EXP-CAP-022 · Grados de clasificación | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-023 · Órganos de contratación | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-024 · Oficinas de programa | `specs/catalogs.md` · `application/catalogs` | — | — | — | Verified-static |
| EXP-CAP-025 · Bandeja paginada con filtros | `specs/query-and-tasks.md` · `application/queries` | Único `RowSource` con SQL sobre el motor del frontend (`Form_FormE2ECola.cls:137`) | expediente | — | Verified-static |
| EXP-CAP-026 · Búsqueda avanzada | `specs/query-and-tasks.md` · `application/queries` | La semántica de filtros hay que preservarla | expediente | — | Verified-runtime (histórico) + Verified-static |
| EXP-CAP-027 · Búsqueda técnica sujeta a permiso efectivo | `specs/query-and-tasks.md` · `application/queries` · `application/authorization` | El permiso efectivo no está trazado como contrato (ver [`06-seguridad-y-trazabilidad.md`](../06-seguridad-y-trazabilidad.md)) | expediente | — | Verified-static |
| EXP-CAP-028 · Exportación Excel reproducible y autorizada | `specs/query-and-tasks.md` | Sin módulo web declarado | expediente | — | Verified-static |
| EXP-CAP-029 · Tareas calculadas por reglas del expediente | `specs/query-and-tasks.md` | Sin tabla de tareas que inventar | expediente | — | Intended + Verified-static |
| EXP-CAP-030 · Autosave de datos generales y fechas | `specs/write-resilience.md` · `application/autosave_general` | Retirada en RC1 con `GuardadoAutomaticoHelper.bas` | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-031 · Autosave de hitos y modificados | `specs/write-resilience.md` · `application/autosave_related` | Retirada en RC1 con `GuardadoAutomaticoHelper.bas` | expediente | — | Verified-static |
| EXP-CAP-032 · Idempotencia y feedback de progreso | `specs/write-resilience.md` · `application/feedback_guard` | Sin equivalente comprobado en Access | — | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-033 · DTO de dominio estable | `specs/e2e.md` · `application/e2e_dto` | — | — | — | Verified-static |
| EXP-CAP-034 · Exportación JSON completa y determinista | `specs/e2e.md` · `application/json_canonical` | — | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-035 · Gestión y ejecución batch E2E | `specs/e2e.md` · `application/batch_e2e` | `Form_FormE2EGestionBatch` está retirado en RC1 | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-036 · Hash de contenido versionado | `specs/e2e.md` · `application/e2e_hash` · `domain/hash/versioning` | — | — | — | Verified-runtime (histórico) |
| EXP-CAP-037 · Generación del paquete de exportación E2E | `specs/e2e.md` · `application/export_package` | — | — | — | Verified-runtime (histórico) |
| EXP-CAP-038 · Trazabilidad entre expediente, batch, artefacto y resultado | `specs/e2e.md` · `application/batch_e2e` | Capacidad añadida en RC1 | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-039 · Selección manual de expedientes para batch | `specs/e2e.md` · `application/batch_e2e` | Capacidad añadida en RC1 (`Form_FormE2EGestionEnlaces`) | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-040 · Destino de exportación por usuario | `specs/e2e.md` · `application/e2e_session` | Sin equivalente comprobado en Access | expediente | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) |
| EXP-CAP-041 · Sesión de exportación, reanudación y cierre | `specs/e2e.md` · `application/e2e_session` | Sin equivalente comprobado en Access | expediente | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) |
| EXP-CAP-042 · Ordinal E2E único y consistente | `specs/e2e.md` · `application/ordinal_e2e` | — | expediente | — | Verified-runtime (histórico) |
| EXP-CAP-043 · Autorización deny-by-default por capacidad | `specs/access-control.md` · `application/authorization` | La autorización vive hoy en controles de interfaz (`H-LAN-12`) | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-044 · Login y principal actual desde Lanzadera | `specs/access-control.md` · `application/current_principal` | La identidad se consume del backend de Lanzadera | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-045 · Estado de usuario y sesión sin globals | `specs/access-control.md` · `application/session_state` | Globals y singletons mutables (D178) | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-046 · Auditoría de accesos, cambios y exportaciones | `specs/access-control.md` · `application/audit` | El actor del último cambio no está trazado (coexistencia, apartado 8) | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-047 · Arranque y readiness con fallos diagnosticables | `specs/runtime.md` · `di/readiness.py` | Sin equivalente comprobado en Access | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) |
| EXP-CAP-048 · Configuración tipada por entorno | `specs/runtime.md` · `di/config.py` | La ruta y la clave del backend salen de una tabla local del frontend | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static |
| EXP-CAP-049 · Caché invalidable y read-models reconstruibles | `specs/runtime.md` | Las cachés en memoria por sesión no las invalida la web (coexistencia, apartado 6) | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) |
| EXP-CAP-050 · Binding de backend por despliegue | `specs/runtime.md` · `application/runtime_policies` | El selector interactivo de backend y las rutas embebidas están para retirar | contrato de identidad de Lanzadera | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) |
| EXP-CAP-051 · Alta y consulta desde HPS | `specs/integrations.md` · `application/hps` · `ports/hps.py` | El contrato autenticado e idempotente está por confirmar | HPS | — | Verified-static + Intended |
| EXP-CAP-052 · Intercambio con AGEDYS | `specs/integrations.md` | El ownership y los errores del intercambio no están definidos | AGEDYS | — | Intended |
| EXP-CAP-053 · Enlace con Gestión de Riesgos | `specs/integrations.md` · `application/riesgos` · `ports/riesgos.py` | El identificador estable no está comprobado | Gestión de Riesgos | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-static + Intended |
| EXP-CAP-054 · Enlace con No Conformidades | `specs/integrations.md` · `application/nc` · `ports/nc.py` | `TbNoConformidades` es una tabla vinculada a otro `.accdb` | No Conformidades | concepto de la web; sin objeto equivalente identificado en RC1 | Verified-runtime (histórico) + Intended |
| EXP-CAP-055 · Notificaciones por correo | `specs/integrations.md` · `ports/notification_delivery.py` | — | — | — | Verified-static + Intended |
| EXP-CAP-056 · Documentos en SharePoint u object storage | `specs/integrations.md` · `ports/document_storage.py` | La referencia opaca, la retención y la autorización no tienen adaptador declarado | SharePoint u object storage | — | Verified-static + Intended |
| EXP-CAP-057 · Resolución de red Win32 | Puerto de infraestructura y health/readiness (`di/readiness.py`) | — | — | mecanismo legacy a retirar | — |
| EXP-CAP-058 · Gestión de procesos Win32 | Jobs o servicios supervisados y observables | — | — | mecanismo legacy a retirar | — |
| EXP-CAP-059 · Drag & drop OLE y controles ActiveX | Interacción web accesible que conserva la jerarquía | — | — | mecanismo legacy a retirar | — |
| EXP-CAP-060 · Globals y singletons mutables | Estado por request o sesión y dependencias explícitas (`application/session_state`) | D178 | — | mecanismo legacy a retirar | — |
| EXP-CAP-061 · Cambio interactivo de backend y rutas embebidas | Configuración tipada por despliegue (`di/config.py`) | — | — | mecanismo legacy a retirar | — |
| EXP-CAP-062 · Popup modal anti-spam y busy | Claves de idempotencia, controles deshabilitados y `aria-busy` (`application/feedback_guard`) | — | — | mecanismo legacy a retirar | — |
| EXP-CAP-063 · Menú JSON-hub y navegación acoplada a formularios | Rutas SSR autorizadas por capability | — | — | mecanismo legacy a retirar | — |

## Lanzadera

Fuente: [`prd/capacidades.md`](../03-aplicaciones/lanzadera/prd/capacidades.md) · commit fijado `1474e8e` · identificadores `C-LAN-1..15`.

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| C-LAN-1 · Autenticación y sesión | `app/src/modules/lanzadera/application/login.py` · `domain/session.py` · `openspec/changes/lanzadera-mvp/specs/auth-core` | `H-LAN-1`: un administrador de Windows entra sin contraseña; `H-LAN-2`: el bloqueo nunca caduca | identidad de Lanzadera | — | — |
| C-LAN-2 · Alta, baja y perfil de usuario | `application/create_user.py` · `application/disable_user.py` · `specs/users` | `H-LAN-14`: nunca se envía la clave inicial; `H-LAN-15`: el usuario se localiza por correo | identidad de Lanzadera | — | — |
| C-LAN-3 · Restablecimiento y cambio de contraseña | `application/set_password.py` · `domain/services/issue_reset_token.py` · `domain/services/consume_reset_token.py` · `specs/auth-reset` | `H-LAN-3` y `H-LAN-16`: la clave de un solo uso se guarda y viaja en claro | identidad de Lanzadera | — | — |
| C-LAN-4 · Catálogo de aplicaciones | `application/create_app.py` · `application/update_app.py` · `application/disable_app.py` · `specs/apps` | `H-LAN-8`: el alta parece inalcanzable y sobrescribe la primera fila de la tabla | — | — | — |
| C-LAN-5 · Visibilidad por ubicación y estado | `application/list_effective_apps.py` · `specs/apps` | `H-LAN-13`: las etiquetas de versión de SICA y AGEDO muestran la versión de otra aplicación | — | disposición aprobada: retirar la segmentación por ubicación | — |
| C-LAN-6 · Lanzamiento y paso de identidad | `application/get_my_apps.py` · `delivery/http/auth_routes.py` (`/auth/me/apps`) | `H-LAN-12`: `Lanzar` no comprueba permisos y el botón de EXPEDIENTES se habilita para cualquiera | identidad de Lanzadera | disposición aprobada: retirar el lanzador de escritorio | — |
| C-LAN-7 · Asignación usuario-aplicación | `application/assign_profile.py` · `domain/assignment.py` · `specs/assignments` | `H-LAN-17`: en Perfil, «Eliminar» actúa sobre la aplicación del combo y no sobre la fila elegida | permisos de Lanzadera | — | — |
| C-LAN-8 · Roles y permisos granulares | `domain/assignment.py` · `specs/assignments` | El enforcement central está pendiente de la disposición aprobada | permisos de Lanzadera | — | — |
| C-LAN-9 · Administrador global | `application/grant_global_admin.py` · `application/revoke_global_admin.py` · `application/bootstrap_global_admins.py` · `specs/global_admins` | — | permisos de Lanzadera | — | — |
| C-LAN-10 · Menús por perfil | F2 · Menú Admin, F3 · Menú Calidad, F4 · Menú Usuario (épica) | `H-LAN-5`: los menús nunca se desactivan; `H-LAN-7`: la navegación oculta fallos con `On Error Resume Next` | permisos de Lanzadera | disposición aprobada: sustituir por menú global con permisos | — |
| C-LAN-11 · Contenido formativo | sin candidato declarado | `H-LAN-18` a `H-LAN-22`: el árbol de vídeos solo muestra una categoría y la carga de `FormVideosUsuarios` está comentada | — | disposición aprobada: retirar; el histórico queda archivado y no se migra | — |
| C-LAN-12 · Conexiones y auditoría | `application/track_presence.py` · `application/get_connected_users.py` · `application/audit_append.py` · `specs/audit` | La trazabilidad la dan las tablas de aperturas, ediciones y conexiones | identidad de Lanzadera | disposición aprobada: retirar la telemetría de SSID y ubicación | — |
| C-LAN-13 · Cola de correo | `ports/notification_delivery.py` | — | — | disposición aprobada: reemplazar el mecanismo por un servicio de notificaciones | — |
| C-LAN-14 · Configuración de backend y entorno | sin candidato declarado | La clave del backend vive en la tabla local `TbConfiguracionBackends` y `Config_BackendHelper.bas` la escribe | — | disposición aprobada: retirar como capacidad de administración | — |
| C-LAN-15 · Tareas, batch y exportaciones | sin candidato declarado | — | — | disposición aprobada: fusionar con el scheduler compartido | — |

## Condor

Fuente: [`prd/capacidades.md`](../03-aplicaciones/condor/prd/capacidades.md) · commit fijado `c88347e` · identificadores `C-CON-1..22`.

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| C-CON-1 · Arranque, identidad y validación crítica | F7 · Vinculación con Lanzadera y Expedientes vía adaptadores unificados | D93: contraseña hardcodeada como fallback en `GetPasswordDB` | identidad de Lanzadera | — | — |
| C-CON-2 · Roles y jerarquía | sin candidato declarado | `JerarquiaRolesHelper.bas` no existe en el commit; la jerarquía vive en `UsuarioServicio.DeterminarRol` | permisos de Lanzadera | — | — |
| C-CON-3 · Configuración de backend | sin candidato declarado | La configuración de conexión vive en el código (`FUNCIONES UTILES.bas:233-252`) | — | retirada | — |
| C-CON-4 · Modo de pruebas y sandbox seguro | sin candidato declarado | El aislamiento pasa a copia de ficheros, no a enrutado de `getdb()` | — | — | — |
| C-CON-5 · Login por comando de Lanzadera | F7 · Vinculación con Lanzadera y Expedientes vía adaptadores unificados | — | identidad de Lanzadera | — | — |
| C-CON-6 · Gestión de solicitudes | F1 · Solicitudes: alta, edición, baja, cambio de estado, snapshot, búsqueda y filtros | D94: claves foráneas conceptuales sin garantía; P-CON-4: no existe reversión de corte | — | — | — |
| C-CON-7 · Tipos de solicitud | F2 · Tipos de solicitud con datos específicos | El tipo `PC_SUB` no existe en el commit: quedan tres tipos | — | el tipo `PC_SUB` no existe en el commit; quedan tres tipos | — |
| C-CON-8 · Catálogo de estados y transiciones | F3 · Workflow declarativo con transiciones y roles requeridos | D96: el motor es data-driven y necesita semilla y evolución de `transiciones` | — | — | — |
| C-CON-9 · Validación de calidad | F5 · Validación de calidad con hash de integridad | P-CON-6: el algoritmo de hash debe sobrevivir al cambio | — | — | — |
| C-CON-10 · Historial de rechazos | F4 · Recepción de rechazo e historial de resoluciones | `tbHistorialRechazos` no existe; la unicidad de «un rechazo activo» pasa a ser restricción | — | — | — |
| C-CON-11 · Adjuntos por etapa | F6 · Adjuntos asociados a etapa del workflow | Una etapa admite un adjunto por la fecha de subida más reciente: dos subidas simultáneas, gana la última | — | — | — |
| C-CON-12 · Visor de estado embebido | sin candidato declarado | El control de navegador sigue presente y el HTML se genera con caché en disco | — | — | — |
| C-CON-13 · Motor de workflow | F3 · Workflow declarativo con transiciones y roles requeridos | `EjecutarTransicion` comprueba precondiciones y luego escribe: dos sesiones pueden transicionar la misma solicitud | — | — | — |
| C-CON-14 · Vinculación con expediente | F7 · Vinculación con Lanzadera y Expedientes vía adaptadores unificados | P-CON-2: cambiar el formato del código de solicitud invalida los códigos ya emitidos | Expedientes | — | — |
| C-CON-15 · Vinculación con no conformidades | F8 · Vinculación con NoConformidades vía adaptador | D95: la relación es por código y NoConformidades debe migrarse antes que Condor; P-CON-3 | No Conformidades | — | — |
| C-CON-16 · Suministradores | F1 · Solicitudes | — | — | — | — |
| C-CON-17 · Documentos y PDF | sin candidato declarado | Los documentos, los PDF y las instantáneas se escriben en rutas locales, fuera de todo control de acceso | — | — | — |
| C-CON-18 · Notificaciones | sin candidato declarado | La simulación envía notificaciones reales: falta un doble explícito | — | — | — |
| C-CON-19 · Auditoría en tres tablas | sin candidato declarado | — | — | — | — |
| C-CON-20 · Mapeo de campos | sin candidato declarado | D97: las 183 reglas de `tbMapeoCampos` deben sobrevivir como datos de runtime | — | — | — |
| C-CON-21 · Hashing y snapshot | F5 · Validación de calidad con hash de integridad | — | — | — | — |
| C-CON-22 · Batería de pruebas | sin candidato declarado | 166 procedimientos de prueba sin arnés equivalente en la web | — | — | — |

## Brass

Fuente: [`prd/capacidades.md`](../03-aplicaciones/brass/prd/capacidades.md) · commit fijado `3f77d58` · identificadores `CAP-1..12` (provisionales en el PRD).

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| CAP-1 · Arranque, identidad y permisos | F1 · Menú principal | D104: contraseña real hardcodeada; `EVE` aborta si el usuario no es administrador ni técnico | identidad de Lanzadera | sin formulario de login propio | — |
| CAP-2 · Opciones y configuración de la BD | F1 · Menú principal, sub-menú «Otras opciones» | D-BRA-3: contraseña literal en el código, en tres apariciones | — | — | — |
| CAP-3 · Eventos (alta, edición, copia, franqueo, SLA, alta masiva) | F9 · Eventos (maestro), F10 · Alta y edición, F11 · Franqueo, F12 · SLA, F13 · Copia, F14 · Alta masiva por NS | D107: el SLA tiene tres estados y el «inconsistente» no es un error; D111: 33 botones en el formulario de gestión | — | — | — |
| CAP-4 · Actividades y tipos de actividad | F19 · Actividades, F20 · Tipos de actividad | D125: el identificador se calcula por `Max + 1`, con carrera en concurrencia | — | — | — |
| CAP-5 · Equipos, equipos de medida y calibraciones | F6 · Equipos de medida, F7 · Calibraciones, F8 · Equipos calibrables, F15 · Equipos de medida del evento | D122: estado compartido por globals; D124: comunicación por `WithEvents` y `RaiseEvent` | — | — | — |
| CAP-6 · Planificación y partes | F17 · Planificación preventiva, F18 · Cuadro de mando de planificación, F43 · Partes de trabajo | D118: un parte firmado por cada Centro dentro de una factura | — | — | — |
| CAP-7 · Materiales, reparaciones y repuestos | F21 · Materiales reparación, F22 · Materiales repuesto, F23 · Materiales intervención, F24 · Router Material | D112: filtro `EquipoID = 538` hardcodeado en el alta de reparación | — | — | — |
| CAP-8 · Facturación, gastos y subcontratación | F27 · Alta de factura, F28 · Resultado de factura, F29 · Gestión de facturas, F30 · Gastos y dietas, F42 · Subcontrataciones | D117: seis tablas de hechos que no se deben consolidar; D120: IVA y recargo hardcodeados | — | — | — |
| CAP-9 · Técnicos, ausencias, fiestas y precios | F35 · Técnicos, F36 · Tipos de técnico, F37 · Precios de tipo de técnico, F38 · Ausencias, F39 · Fiestas, F40 · Consulta de horas, F41 · Calendario laboral | D113: jornada de 7,5 h fija; D119: precios como SCD tipo 2 manual | — | — | — |
| CAP-10 · Informes (SLA, RAC, desde lista, por factura) | F16 · Informe SLA, F31 · Informe desde lista, F32 · Informe para RAC, F33 · Informe por factura | D115: `TbAuxEventosParaInforme` es una tabla temporal sin feeder identificado | — | — | — |
| CAP-11 · Anexos | F25 · Anexos polimórficos | D114: anexos en el sistema de ficheros, sin versionado ni análisis antivirus | — | — | — |
| CAP-12 · Estructura y catálogos (BUI, nodos, subsistemas, ubicación, originador) | F2 · Gestión de nodos, F3 · Gestión de BUI, F4 · Gestión de subsistemas, F5 · Gestión de ubicaciones, F44 · Originadores | — | — | — | — |

## Gestión de Riesgos

Fuente: [`prd/capacidades.md`](../03-aplicaciones/gestion-riesgos/prd/capacidades.md) · commit fijado `36fcf60` · el estudio no numera las capacidades, de modo que el identificador va con `—`.

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| — · Arranque e identidad | F6 · Splash y autenticación dual | — | identidad de Lanzadera | — | — |
| — · Roles (administrador, calidad, técnico) | F6 · Splash y autenticación dual | La identidad se resuelve leyendo el backend de Lanzadera, no un servicio | identidad de Lanzadera | la identidad se resuelve leyendo el backend de Lanzadera, no un servicio | — |
| — · Configuración temporal | sin candidato declarado | El modelo conmutable se fuerza a «nuevo» en el arranque | — | — | — |
| — · Proyecto | F3 · Gestión de Proyectos y UTE | — | — | — | — |
| — · Edición | F11 · Alta y edición directa de riesgo, F12 · Gestión de Riesgos | — | — | — | — |
| — · Riesgo (alta, edición, borrado) | F11 · Alta y edición directa de riesgo, F12 · Gestión de Riesgos, F24 · Gestión principal de riesgos | D140: sin locks en concurrencia y `UPDATE` directo sobre las ediciones | — | — | — |
| — · Estados del riesgo (14 estados) | F23 · Estado del riesgo | D126: la máquina de estados se calcula en runtime desde fechas | — | pendiente de ficha | — |
| — · Causa raíz y biblioteca | F5 · Biblioteca de Riesgos, F7 · CRUD de RiesgosBiblioteca | — | — | — | — |
| — · Riesgo externo | F21 · Riesgo Externo (Oferta) | — | — | — | — |
| — · Materialización | F15 · Materialización del riesgo, F35 · Capturar plan contingencia | El formulario de alta no está en `main` | — | — | — |
| — · Planes de mitigación (PM) | F47 · Alta de plan individual | — | — | — | — |
| — · Planes de contingencia (PC) | F35 · Capturar plan contingencia, F47 · Alta de plan individual | — | — | — | — |
| — · Evidencias y anexos | F53 · Anexos polimórficos | D142: tres copias de la misma implementación, no polimorfismo | — | — | — |
| — · Suministradores por edición | F2 · Gestión de suministros, F37 · Sincronización de suministros, F44 · Evidencias de suministros | — | — | — | — |
| — · Control de cambios | F52 · Gestión de cambios de la aplicación | D134: audita las versiones de la aplicación, no los riesgos | — | — | — |
| — · Publicabilidad | F38 · Flujo de publicación de calidad, F40 · Ejecutar publicación, F48 · Integración con AGEDO, F49 · Informes de publicación, F51 · Auditoría de publicaciones | D128: workflow dual técnico-calidad; D129: integración con AGEDO; D137: publicar no es transaccional | — | — | — |
| — · Calidad — tareas | F27 · Dashboard de calidad, F28 · Detalle de edición pendiente, F29 · Visado de aceptación y retiro | D138: `WithEvents` sobre-amplia, con 32 aristas | — | — | — |
| — · Técnico — tareas | F56 · Tareas pendientes del técnico, F57 · Detalle de tarea, F58 · Riesgo aceptado o retirado, F59 · Explicación de nodo | D139: el estado del árbol se derrama entre formularios por un global | — | — | — |
| — · Indicadores | F4 · Dashboard de Indicadores | Se apoyan en consultas efímeras, sin consulta guardada que replicar | — | — | — |
| — · Informes | F49 · Generación de informes de publicación | La salida es un fichero en disco, no una descarga autenticada | — | — | — |
| — · Histórico de estados del riesgo | F23 · Estado del riesgo | — | — | pendiente de ficha | — |
| — · Catálogos | F17 · Mitigación, F20 · Vinculación con No Conformidades, F48 · Integración con AGEDO | — | — | — | — |
| — · Exportación e impresión | F49 · Generación de informes de publicación | La salida es un fichero local | — | — | — |
| — · Integraciones | F20 · Vinculación con No Conformidades, F48 · Integración con AGEDO, F54 · Búsqueda cross-app de Expedientes | Se abre un backend por integración, con 15 tablas vinculadas de seis backends; D132: handshake HTTP/JSON sin auth documentado | seis backends externos | se abre un backend por integración: 15 tablas vinculadas de seis backends | — |
| — · Acoplamiento DAO directo | sin candidato declarado | D88 y D-GR-2: 362 llamadas a `getdb` en 42 ficheros de producción | — | — | — |
| — · Concurrencia sin atomicidad | sin candidato declarado | Un único `BeginTrans` y ningún `RollbackTrans`; identificadores por `Max + 1` | — | — | — |

## No Conformidades

Fuente: [`prd/capacidades.md`](../03-aplicaciones/no-conformidades/prd/capacidades.md) · commit fijado `1d7a2f8` · identificadores `CAP-NC-1..45`.

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| CAP-NC-1 · Arranque con cálculo del rol | F3 · Splash y autenticación dual | La identidad se determina por el entorno del equipo y el proveedor de identidad de la web está pendiente de decisión (coexistencia, apartado 6) | identidad de Lanzadera | modificada: los formularios son iguales y las variables globales y el entorno difieren del snapshot | — |
| CAP-NC-2 · Identidad y roles | F3 · Splash y autenticación dual | El rol técnico se asigna por exclusión en `main` y su decisión sigue abierta (P-NC-1) | identidad de Lanzadera | la identidad y los permisos no se reespecifican aquí: su propietario es Lanzadera | — |
| CAP-NC-3 · Configuración de backend e `IDAplicacion` | sin candidato declarado | D-NC-3: el ERD de origen no está extraído y el destino de la configuración queda sin decidir | — | `IDAplicacion` tiene un valor por entorno; sin configuración legible se informa del error | — |
| CAP-NC-4 · Menús de entrada | F1 · Menú principal, F2 · Menú técnico | Los enlaces sin handler en `main` se documentan como defecto, no como requisito (P-NC-2) | — | mixta: el menú raíz no cambia y los dos submenús sí | — |
| CAP-NC-5 · Selección de auditoría y alta o edición de NC de auditoría | F16 · Listado de NCs de auditoría, F17 · Gestión de NCs de auditoría, F18 · Alta y edición de NC de Auditoría | `RegistrarDatosUnicos` rechaza combinaciones no válidas: la precondición solo la aplica VBA (coexistencia, apartado 4) | — | modificada: difieren del snapshot el formulario general de auditoría y sus operaciones | — |
| CAP-NC-6 · Gestión de NC de auditoría | F17 · Gestión de NCs de auditoría | Ninguna escritura de negocio está dentro de una transacción de base de datos (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-7 · Alta y edición de NC de proyecto | F33 · Alta y edición de NC de Proyecto | El identificador y el código de NC se calculan en cliente con `Max + 1`, con carrera en concurrencia (coexistencia, apartados 2 y 4) | Expedientes | modificada; el identificador funcional `CodigoNC`/`Nemotecnico` necesita decisión de unicidad y de vacíos | — |
| CAP-NC-8 · Búsqueda y listado de NC de proyecto | F32 · Gestión y listado de NCs de Proyecto | El SQL se construye por concatenación y el filtrado se aplica en memoria tras leer todas las filas (coexistencia, apartado 2) | — | modificada: el formulario tiene código distinto del snapshot | — |
| CAP-NC-9 · Borrado lógico y habilitación de NC | F32 · Gestión y listado de NCs de Proyecto, F17 · Gestión de NCs de auditoría | Ninguna apertura declara bloqueo explícito y el borrado con motivo no está en transacción (coexistencia, apartados 3 y 5) | — | modificada: ambas clases de operaciones difieren del snapshot | — |
| CAP-NC-10 · Alta de NC vinculada y borrado en cascada | F34 · Alta y edición de NC de Proyecto con vínculo a NC existente | El alta escribe la NC y el log en pasos sin transacción (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-11 · Acciones correctivas | F35 · Acción Correctiva NC-Proyecto, F19 · Registro de Acción Correctiva de NC de Auditoría, F39 · Acciones NC-Proyecto, F23 · Acciones de NC de Auditoría | El alta de una acción correctiva escribe alta, log y recálculo sin límite transaccional (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-12 · Acciones realizadas | F36 · Acción Realizada NC-Proyecto, F20 · Registro de Acción Realizada de NC de Auditoría, F39 · Acciones NC-Proyecto, F23 · Acciones de NC de Auditoría | El borrado de una acción realizada actualiza los datos calculados sin transacción (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-13 · Historial de acciones y recálculo de estados | F39 · Acciones NC-Proyecto, F23 · Acciones de NC de Auditoría | El recálculo de estados vive en las clases de acciones y no en el servidor (coexistencia, apartado 4) | — | modificada | — |
| CAP-NC-14 · Plan de control de eficacia | F37 · Control de Eficacia NC-Proyecto, F38 · Alta Control de Eficacia NC-Proyecto, F21 · Control de Eficacia de NC de Auditoría, F22 · Alta de Control de Eficacia de NC de Auditoría | El plan se graba junto con la NC, sin transacción (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-15 · Resultado del control de eficacia | F37 · Control de Eficacia NC-Proyecto, F21 · Control de Eficacia de NC de Auditoría | Solo es posible con un plan previo y esa precondición está por verificar en la ficha (CAP-NC-15) | — | modificada | — |
| CAP-NC-16 · Motivo de no requerir control de eficacia | F6 · Catálogo de motivos «no requiere Control de Eficacia» | El catálogo es compartido con Brass y Gestión de Riesgos y su relación concreta queda pendiente (coexistencia, apartado 8) | Brass y Gestión de Riesgos (catálogo compartido, D86) | sin cambios en el formulario; el texto se devuelve por evento y no escribe tablas | — |
| CAP-NC-17 · Replanificación de acciones y tareas | F43 · Replanificación NC-Proyecto, F27 · Replanificación de NC de Auditoría | El registro de replanificación se valida en las operaciones, sin transacción (coexistencia, apartado 3) | — | modificada; el alta por formulario no existe (a verificar) | — |
| CAP-NC-18 · Consulta de replanificaciones | F43 · Replanificación NC-Proyecto, F27 · Replanificación de NC de Auditoría | — | — | modificada: los formularios no cambian y se clasifica como modificada por depender de CAP-NC-17 | — |
| CAP-NC-19 · Documentos de NC | F40 · Documentos NC-Proyecto, F24 · Documentos de NC de Auditoría | El borrado de un documento deja filas y ficheros sin atomicidad (P-NC-8) | — | modificada | — |
| CAP-NC-20 · Documentos de evidencia de AR | F9 · Documentos AR Auditoría, F10 · Documentos AR Proyecto | Las evidencias se gestionan sobre la acción realizada activa, sin transacción (coexistencia, apartado 3) | — | modificada | — |
| CAP-NC-21 · Documentos de auditoría | F15 · Documentos de auditoría | Antes del alta los anexos se acumulan en memoria y se graban al registrar la auditoría (P-NC-8) | — | modificada | — |
| CAP-NC-22 · Nota libre de la NC | F41 · Notas NC-Proyecto, F25 · Notas de NC de Auditoría | El saneado de texto se aplica al alta de acciones y no a las notas (coexistencia, apartado 4) | — | modificada por las operaciones | — |
| CAP-NC-23 · Motivo de borrado de la NC | F42 · Motivo de eliminación NC-Proyecto, F26 · Motivo de eliminación de NC de Auditoría | La corrección del motivo de una NC ya borrada no está en transacción (coexistencia, apartado 3) | — | modificada por las operaciones | — |
| CAP-NC-24 · Seguimiento de tareas | F44 · Seguimiento NC-Proyecto, F46 · Seguimiento Tareas NC-Proyecto, F28 · Seguimiento de NC de Auditoría, F30 · Seguimiento Tareas de NC de Auditoría | Las listas se alimentan de colecciones del entorno y no se refrescan ante cambios de otro usuario (coexistencia, apartado 5) | — | modificada | — |
| CAP-NC-25 · Revisión de estados de NC | F45 · Seguimiento NC NC-Proyecto, F29 · Seguimiento NC de NC de Auditoría | Los filtros se aplican sobre colecciones del entorno, sin refresco automático (coexistencia, apartado 5) | — | modificada | — |
| CAP-NC-26 · Log de actividad por NC | sin candidato declarado | Los logs no registran el cliente (Access o web) y las escrituras con `On Error Resume Next` pierden el error (coexistencia, apartado 6) | — | modificada: el log difiere en `Funciones Generales.bas` | — |
| CAP-NC-27 · Alta y edición de cabecera de auditoría | F13 · Alta y edición de cabecera de auditoría | El alta valida y graba en `Registrar`, sin transacción (coexistencia, apartado 3) | — | sin cambios en el formulario | — |
| CAP-NC-28 · Catálogo de auditorías | F14 · Gestión de auditorías | El borrado de una auditoría deja documentos huérfanos (P-NC-11) | — | modificada | — |
| CAP-NC-29 · Catálogo de tipologías de NC | F47 · Gestión de tipologías NC-Proyecto, F48 · Alta y edición de tipología NC-Proyecto | El identificador se calcula con `Max + 1` (coexistencia, apartado 2) | — | sin cambios | — |
| CAP-NC-30 · Vinculación con riesgos | F7 · Vinculación con Gestión de Riesgos | El vínculo borra y reinserta los riesgos sin transacción: si falla un `INSERT`, el vínculo queda a medias (coexistencia, apartado 3) | Gestión de Riesgos | modificada: `RiesgoServicio1.bas` es nuevo respecto al snapshot y está vacío en `main` | — |
| CAP-NC-31 · Búsqueda de expedientes | F8 · Búsqueda cross-app con Expedientes | La búsqueda se hace con un `SELECT *` filtrado en VBA (coexistencia, apartado 8) | Expedientes | sin cambios; los expedientes no se editan aquí | — |
| CAP-NC-32 · Indicadores semestrales | F4 · Dashboard de Indicadores | Los indicadores se calculan sobre las tablas, sin caché, con `Format` y `#` no portables (coexistencia, apartado 2) | — | sin cambios | — |
| CAP-NC-33 · Correo de la NC | F5 · Envío de correo manual | El informe HTML se construye con `Format` y `#` no portables a PostgreSQL (coexistencia, apartado 2) | — | sin cambios | — |
| CAP-NC-34 · Aviso al administrador ante errores | F5 · Envío de correo manual | La contraseña del backend viaja en la cadena de conexión (coexistencia, apartado 6) | — | modificada | — |
| CAP-NC-35 · Informe Word y exportación a Excel | sin candidato declarado | El SQL del informe se construye por concatenación (coexistencia, apartado 2) | — | modificada | — |
| CAP-NC-36 · Informe HTML de NC | sin candidato declarado | Las fechas con `#` y `Format` no son portables a PostgreSQL (coexistencia, apartado 2) | — | sin cambios | — |
| CAP-NC-37 · Utilidades del instalador | sin candidato declarado | Su alcance en la web está por decidir (P-NC-16) | — | sin cambios | — |
| CAP-NC-38 · Clase de informe de NC de auditorías | sin candidato declarado | No hay comportamiento que aceptar: la clase solo declara el error (CAP-NC-38) | — | añadida: no está en el snapshot | — |
| CAP-NC-39 · Caché de indicadores | sin candidato declarado | Retirada en `main`: no hay objeto que portar (P-NC-17) | — | retirada | — |
| CAP-NC-40 · Caché de listados de NC | sin candidato declarado | Retirada en `main`: no hay objeto que portar (P-NC-17) | — | retirada | — |
| CAP-NC-41 · Kill switch y bancos de prueba | sin candidato declarado | Retirada en `main`: no hay objeto que portar (P-NC-17) | — | retirada | — |
| CAP-NC-42 · Ayudantes y servicios retirados | sin candidato declarado | Retirada en `main`: la lógica vive en los formularios y en las clases de operaciones (P-NC-17) | — | retirada | — |
| CAP-NC-43 · Viewmodels de lista y detalle | sin candidato declarado | Retirada en `main`: la web define sus propios DTO (P-NC-17) | — | retirada | — |
| CAP-NC-44 · Caché por NC | sin candidato declarado | La lectura desde caché está desactivada y solo se invalida; su portabilidad depende del ERD global (P-NC-18) | — | modificada | — |
| CAP-NC-45 · Mantenimiento de la caché | sin candidato declarado | Los comandos de mantenimiento no tienen equivalente web declarado (P-NC-18) | — | modificada | — |

## HPS Solicitudes

Fuente: [`prd/capacidades.md`](../03-aplicaciones/hps-solicitudes/prd/capacidades.md) · commit fijado `cfa1937` · identificadores `C-HPS-S-1..18`.

| capacidad | componente web candidato | riesgo principal | dependencias | excepción | confianza |
|---|---|---|---|---|---|
| C-HPS-S-1 · Arranque e identidad | F1 · Arranque e identidad | D-HPS-8: la identidad vive en las tablas vinculadas `TbUsuarios` de `HPST.accdb` y consolidar o federar está sin decidir | identidad compartida con Lanzadera | disposición aprobada: preservar y modernizar; integrar con la identidad única de la web | — |
| C-HPS-S-2 · Configuración de backend | F2 · Configuración de backend | D-HPS-1: las claves y las tablas del backend no están extraídas | — | disposición aprobada: preservar `TbConfiguracion` en backend (singleton, decisión del propietario) | — |
| C-HPS-S-3 · Flags de operación | F13 · Configuración runtime | Los flags son `TEXT(2)` con valores `Sí`/`No` y su traducción a `BOOLEAN` está planteada pero no aprobada (coexistencia, apartado 4) | — | disposición aprobada: mantener como flags de aplicación durante la coexistencia | — |
| C-HPS-S-4 · Roles | F1 · Arranque e identidad | H-HPS-S-6: el rol se calcula en el arranque, en `Variables Globales.bas`, y no en un servicio | identidad compartida con Lanzadera | disposición aprobada: preservar y modernizar; el modelo de dos niveles se valida en cada endpoint | — |
| C-HPS-S-5 · Solicitudes | F3 · CRUD Solicitud | D-HPS-10: no hay versión de fila, de modo que dos escrituras simultáneas se pisan sin aviso | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-6 · Traspasos a ONS | F5 · Traspasos a ONS | D-HPS-11: `RegistroEnvioTraspasoONS` escribe sin transacción y la operación multi-tabla no es atómica | — | disposición aprobada: preservar y modernizar; transacción atómica obligatoria | — |
| C-HPS-S-7 · Adjuntos Excel | F6 · Adjuntos Excel del solicitante | H-HPS-S-18: el nombre del fichero se interpola en el SQL sin parametrizar | — | disposición aprobada: preservar y modernizar; validación de tipo MIME y tamaño | — |
| C-HPS-S-8 · Plantillas HTML | F7 · Plantillas HTML y vistas web embebidas | Las plantillas se sirven por una ruta UNC y la web necesita URL firmada por tiempo limitado (coexistencia, apartado 4) | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-9 · Renovación | F3 · CRUD Solicitud, F4 · Workflow regulatorio de HPS con 20 fechas | D-HPS-2: 20 columnas `SHORT_DATE_TIME` con la decisión de zona horaria pendiente | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-10 · Cambio de tipo | F3 · CRUD Solicitud | H-HPS-S-29: `Form_FormCambioTipo` no tiene cableado en `src/` | — | disposición aprobada: preservar y modernizar; validación de rol | — |
| C-HPS-S-11 · Justificaciones | F8 · Justificaciones | — | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-12 · Responsables | F9 · Responsables | D99: la relación con la solicitud es por `email` y no por identificador | usuarios de la web (FK a `responsables.id`, decisión P-EXP-3) | disposición aprobada: preservar y modernizar; la FK por `email` se sustituye por FK numérica | — |
| C-HPS-S-13 · Motivos HPS | F10 · Motivos HPS | D-HPS-9: `MotivoHPS.EditarMotivo` actualiza cuatro tablas sin transacción; H-HPS-S-23 y H-HPS-S-29: los formularios no están cableados en `src/` | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-14 · Entidades y expedientes | F11 · Entidades y búsqueda de expedientes | El filtro de expedientes concatena en `InStr` (H-HPS-S-39) y la búsqueda accede a la tabla vinculada `TbExpedientes` (coexistencia, apartado 1) | Expedientes | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-15 · Tareas pendientes del tramitador | F12 · Tareas pendientes del tramitador | Once consultas casi idénticas en `constructor.bas` que la web debe sustituir por una vista materializada; H-HPS-S-49: los `RowSource` cargan nombres y empresas reales | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-16 · Log general | sin candidato declarado | D-HPS-3: `TbLogsGeneral` escribe `DNI` y `Usuario` en plano | — | disposición aprobada: preservar y modernizar; la épica declara REPLACE web-native (D27) | — |
| C-HPS-S-17 · Correos | F14 · Correos automáticos | D-HPS-12: `CDO.Message` se invoca sin TLS explícito y la contraseña del SMTP es un literal en el código | — | disposición aprobada: preservar y modernizar | — |
| C-HPS-S-18 · Registro en HPS | F15 · Registro automático en HPS | P-EXP-3: la equivalencia entre el flag `RegistroEnHPS` y el endpoint `POST /hps/registro` está por verificar | HPS | disposición aprobada: preservar y modernizar; integración por API con autenticación federada | — |

## Lógica a extraer

Capacidades cuya lógica vive en un evento de formulario y necesita un servicio o un caso de uso. El módulo o handler que las sostiene se cita junto a la fila; el destino web es el candidato declarado en la tabla anterior.

### Expedientes

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| EXP-CAP-009 | `Form_FormExpedienteModificados.cls`, `Form_FormModificado.cls` | `application/register_modificado` |
| EXP-CAP-016 | `Form_FormCPV.cls`, `Form_FormCPVsGestion.cls` | `application/catalogs` |
| EXP-CAP-027 | `Form_FormExpedientesGestionTecnica.cls` | `application/queries` con `application/authorization` |
| EXP-CAP-030, EXP-CAP-031 | `Form_FormExpediente.cls` y subformularios; el autosave lived en `GuardadoAutomaticoHelper.bas`, retirado en RC1 | `application/autosave_general`, `application/autosave_related` |
| EXP-CAP-032 | Sin objeto equivalente identificado en RC1: concepto de la web | `application/feedback_guard` |
| EXP-CAP-040, EXP-CAP-041 | Sin objeto equivalente identificado en RC1: concepto de la web | `application/e2e_session` |
| EXP-CAP-051 | `Form_FormExpedienteAltaParaHPS.cls` | `application/hps` |
| EXP-CAP-052 | Sin módulo declarado | sin candidato declarado |
| EXP-CAP-053, EXP-CAP-054 | Sin objeto equivalente identificado en RC1: concepto de la web | `application/riesgos`, `application/nc` |
| EXP-CAP-056 | `Form_FormExpedienteDocumentacion.cls` | `ports/document_storage.py` |
| EXP-CAP-047 a EXP-CAP-050 | Sin objeto equivalente identificado en RC1: concepto de la web | `di/readiness.py`, `di/config.py`, `application/runtime_policies`, `application/session_state` |

Las capacidades de identidad, autorización y auditoría (EXP-CAP-043 a EXP-CAP-046) no tienen objeto equivalente en Access: su contrato se define en la web y se apoya en el contrato de identidad de Lanzadera (ver [`06-seguridad-y-trazabilidad.md`](../06-seguridad-y-trazabilidad.md)).

### Lanzadera

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| C-LAN-1 | `Form_FormLogin.cls`, `Funciones Genreales.bas` (`Login`), `Variables Globales.bas` (`LoginCorrecto`) | `application/login.py`, `domain/session.py` |
| C-LAN-2 | `Form_FormUsuariosGestion.cls`, `Form_FormUsuarioGeneral.cls`, `Form_FormUsuarioDatosGenerales.cls`, `Form_FormPerfil.cls` | `application/create_user.py`, `application/disable_user.py` |
| C-LAN-3 | `Form_FormObtenerContraseña.cls`, `Form_FormUsuarioCambioPass.cls`, `Usuario.cls` (`ResetearPass`) | `application/set_password.py`, `domain/services/issue_reset_token.py`, `domain/services/consume_reset_token.py` |
| C-LAN-4 | `Form_FormAplicacionesGestion.cls`, `Form_FormAplicacionGeneral.cls`, `Form_FormAplicacionDatosGenerales.cls` | `application/create_app.py`, `application/update_app.py`, `application/disable_app.py` |
| C-LAN-5, C-LAN-6 | `Form_FormAplicaciones*.cls`, `Funciones Genreales.bas` (`Lanzar`) | `application/list_effective_apps.py`, `application/get_my_apps.py` |
| C-LAN-7, C-LAN-8 | `Form_FormUsuarioPerfilAplicaciones.cls`, `UsuarioAplicacionPermisos.cls` | `application/assign_profile.py`, `domain/assignment.py` |
| C-LAN-9 | `Usuario.cls` (`EsAdministradorCalculado`) | `application/grant_global_admin.py`, `application/revoke_global_admin.py`, `application/bootstrap_global_admins.py` |
| C-LAN-10 | `Menu.cls`, `Form_FormMenuPrincipalAdmin.cls`, `Form_FormMenuPrincipalCalidad.cls`, `Form_FormMenuPrincipalUsuario.cls` | F2, F3 y F4 · Menús (sin módulo web declarado) |
| C-LAN-11 | `Form_FormVideosGestion.cls` y las otras seis fichas de vídeos | sin candidato declarado (disposición aprobada: retirar) |
| C-LAN-12 | `Form_FormUsuariosConectados.cls`, `Conexion.cls`, `AplicacionApertura.cls` | `application/track_presence.py`, `application/get_connected_users.py`, `application/audit_append.py` |
| C-LAN-13 | `Correo.cls` | `ports/notification_delivery.py` |
| C-LAN-14 | `Variables Globales.bas` (`getdb`), `Config_BackendHelper.bas` | sin candidato declarado |

### Condor

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| C-CON-6, C-CON-12 | `Form_frmGestionSolicitud.cls` (handlers de alta, edición, cambio de estado y `webInfo`) | F1 · Solicitudes; el visor no tiene feature declarada |
| C-CON-14 | `Form_frmGestionSolicitud.cls`, `ExpedienteServicio.cls` | F7 · Adaptadores unificados |
| C-CON-15 | `Form_frmGestionSolicitud.cls`, `NoConformidadServicio.cls` | F8 · Adaptador con No Conformidades |

### Brass

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| CAP-1, CAP-2 | `Constructor.bas`, `Entorno.cls`, `Form0BDOpciones`, `Form0BDOpcionesOtras` | F1 · Menú principal |
| CAP-3 | `Evento.cls`, `FormEventoGestion`, `FormEventoAlta`, `FormEventoEdicion`, `FormEventoCopia`, `FormEventoFranqueo`, `FormEventoSLA`, `FormEventosAltaMasivaPorNS` | F9 a F14 |
| CAP-4 | `Actividad.cls`, `FormActividadAlta`, `FormActividadEdicion`, `FormActividadGestion`, `FormActividadTipo*` | F19, F20 |
| CAP-5 | `EquipoMedida.cls`, `EquipoMedidaCalibracion.cls`, `FormEquipoMedida*`, `FormCalibracion*` | F6, F7, F8, F15 |
| CAP-6 | `Parte.cls`, `Constructor.bas`, `FormPlanificacion*`, `FormParteGestion` | F17, F18, F43 |
| CAP-7 | `Material.cls`, `MaterialSeguimiento.cls`, `FormMaterial*` | F21 a F24 |
| CAP-8 | `Factura.cls`, `Gasto.cls`, `Subcontratacion.cls`, `LIbranza.cls`, `FormFactura*`, `FormGasto*`, `FormSubContratacion*` | F27 a F30, F42 |
| CAP-9 | `Tecnico.cls`, `TecnicoHoras.cls`, `TipoTecnicoPrecio.cls`, `Fiesta.cls`, `FormTecnico*` | F35 a F41 |
| CAP-10 | `Informe.cls`, `GestorInforme.cls`, `SLAHTMLService.bas`, `FormInformeSLA`, `FormInformes*` | F16, F31 a F33 |
| CAP-11 | `Anexo.cls`, `FormAnexos` | F25 |
| CAP-12 | `Constructor.bas`, `Funciones Generales.bas`, `FormNodoGestion`, `FormBUIGestion`, `FormSubSistema*`, `FormUbicacionGestion`, `FormOriginadorGestion` | F2 a F5, F44 |

### Gestión de Riesgos

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| — · Arranque e identidad, — · Roles | `Form_frmSplash.cls`, `Usuario.cls`, `UsuarioAplicacionPermisos.cls` | F6 · Splash y autenticación dual |
| — · Proyecto, — · Edición, — · Riesgo | `Proyecto.cls`, `Edicion.cls`, `Riesgo.cls`, `FormRiesgo`, `FormGestionRiesgos` | F11, F12, F24 |
| — · Estados del riesgo, — · Histórico de estados | `Riesgo.cls`, `FormRiesgosGestionRiesgo` | F23 · Estado del riesgo |
| — · Materialización, — · Planes de contingencia | `RiesgoMaterializacion.cls`, `FormRiesgoMaterializado`, `FormMaterializacionPlanContingencia` | F15, F35 |
| — · Publicabilidad | `PublicabilidadEdicion.bas`, `PublicabilidadRiesgo.bas`, `PublicacionLog.cls`, `FormPublicacion*` | F38 a F51 |
| — · Calidad — tareas | `TareasCalidad.cls`, `ArbolTareasCalidad.cls`, `FormCalidadTareas*`, `FormCalidadRiesgoAceptadoRetiradoVisado` | F27 a F36 |
| — · Técnico — tareas | `TareasTecnico.cls`, `ArbolTareasTecnico.cls`, `FormTecnicoTareas*` | F56 a F60 |
| — · Indicadores | `mIndicador.bas`, `modIndicadorDashboard.bas`, `FormIndicador`, `formIndicadorProyectos` | F4 · Dashboard de Indicadores |

### No Conformidades

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| CAP-NC-1, CAP-NC-2 | `Variables Globales.bas` (`EVE`), `Entorno.cls`, `constructor.bas` (`getUsuario`), `Form_Form0BDTecnicos.cls` | F3 · Splash y autenticación dual |
| CAP-NC-3 | `Variables Globales.bas` (configuración local e `IDAplicacion`) | sin candidato declarado |
| CAP-NC-4 | `Form_Form0BDOpciones.cls`, `Form_Form0BDOpcionesParteProyectos.cls`, `Form_Form0BDOpcionesAuditorias.cls`, `Funciones Generales.bas` | F1, F2 · Menús |
| CAP-NC-5 | `Form_FormAuditoriaSeleccion.cls`, `Form_FormNCAuditoriaGeneral.cls`, `NCaUDITORIAOperaciones.cls` | F16, F17, F18 |
| CAP-NC-6 | `Form_FormNCAuditoriaGestion.cls` | F17 |
| CAP-NC-7 | `Form_FormNCProyectoGeneral.cls`, `NCProyectoOperaciones.cls` | F33 · Alta y edición de NC de Proyecto |
| CAP-NC-8, CAP-NC-9 | `Form_FormNCProyectoGestion.cls`, `NCProyectoOperaciones.cls`, `NCaUDITORIAOperaciones.cls` | F32, F17 |
| CAP-NC-10 | `Form_FormNCProyectoGeneralConVinculoNC.cls`, `NCProyectoOperaciones.cls` | F34 · Alta y edición de NC de Proyecto con vínculo a NC existente |
| CAP-NC-11 a CAP-NC-13 | `Form_FormNCProyectoAcciones.cls`, `Form_FormNCAuditoriaAcciones.cls`, `ACProyectoOperaciones.cls`, `ARProyectoOperaciones.cls` | F19, F20, F23, F35, F36, F39 |
| CAP-NC-14 a CAP-NC-16 | `Form_FormNCProyectoControlEficacia.cls`, `Form_FormNCProyectoControlEficaciaAlta.cls`, `Form_FormNCAuditoriaControlEficacia.cls`, `Form_FormNCAuditoriaControlEficaciaAlta.cls`, `Form_FormMotivosNoRequiereControlEficacia.cls` | F6, F21, F22, F37, F38 |
| CAP-NC-17, CAP-NC-18 | `ACProyectoOperaciones.cls`, `ARProyectoOperaciones.cls`, `ReplanificacionesProyectoOperaciones.cls`, `ReplanificacionesAuditoriaOperaciones.cls` | F27, F43 |
| CAP-NC-19 a CAP-NC-21 | `Form_FormNCProyectoDocumentos.cls`, `Form_FormNCAuditoriaDocumentos.cls`, `Form_FormAuditoriaDocumentos.cls`, `DocumentoProyectoOperaciones.cls`, `DocumentoAuditoriaOperaciones.cls` | F9, F10, F15, F24, F40 |
| CAP-NC-22 | `Form_FormNCProyectoNota.cls`, `Form_FormNCAuditoriaNota.cls` | F25, F41 |
| CAP-NC-23 | `Form_FormNCProyectoMotivoEliminado.cls`, `Form_FormNCAuditoriaMotivoEliminado.cls` | F26, F42 |
| CAP-NC-24 | `Form_FormNCProyectoSeguimientoTareas.cls`, `Form_FormNCAuditoriaSeguimientoTareas.cls`, `SegTareasProyecto.cls` | F28, F30, F44, F46 |
| CAP-NC-25 | `Form_FormNCProyectoSeguimientoNC.cls`, `Form_FormNCAuditoriaSeguimientoNC.cls` | F29, F45 |
| CAP-NC-26 | `Funciones Generales.bas` (escritura del log), `LogNCProyecto.cls`, `LogNCAuditoria.cls` | sin candidato declarado |
| CAP-NC-27 | `Form_FormAuditoria.cls`, `AuditoriaOperaciones.cls` | F13 · Alta y edición de cabecera de auditoría |
| CAP-NC-28 | `Form_FormAuditoriasGestion.cls` | F14 · Gestión de auditorías |
| CAP-NC-29 | `Form_FormNCProyectoTipologiaGestion.cls`, `TipologiaNCProyectos.cls` | F47, F48 |
| CAP-NC-30 | `Form_formRiesgosSeleccion.cls`, `RiesgoRepositorio.bas`, `RiesgoServicio1.bas` | F7 · Vinculación con Gestión de Riesgos |
| CAP-NC-31 | `Form_FormExpedientesBusqueda.cls`, `Expediente.cls` | F8 · Búsqueda cross-app con Expedientes |
| CAP-NC-32 | `Form_FormIndicadores.cls`, `IndicadorServicio.cls`, `IndicadorRepositorio.bas` | F4 · Dashboard de Indicadores |
| CAP-NC-33 | `Form_FormCorreo.cls`, `Correo.cls` | F5 · Envío de correo manual |
| CAP-NC-34 | `Funciones Generales.bas` (aviso al administrador ante errores) | F5 · Envío de correo manual |
| CAP-NC-35, CAP-NC-36 | `Informe.cls`, `HTML.bas`, `Funciones Generales.bas` | sin candidato declarado |
| CAP-NC-37 | `Instalador.bas` | sin candidato declarado |
| CAP-NC-38 | `InformeNCAuditorias.cls` | sin candidato declarado |
| CAP-NC-39 a CAP-NC-43 | Sin objeto equivalente: retiradas en `main` | sin candidato declarado |
| CAP-NC-44, CAP-NC-45 | `CacheNCProyecto.bas`, `InicializadorCache.bas` | sin candidato declarado |

### HPS Solicitudes

| capacidades | módulo o handler que las sostiene | candidato web |
|---|---|---|
| C-HPS-S-1 | `Variables Globales.bas` (`EVE`), `constructor.bas` (`getUsuario`), `Form_frmSplash` | F1 · Arranque e identidad |
| C-HPS-S-2 | `constructor.bas` (`getConfiguracion`), `Configuracion.cls`, `ConfiguracionOperaciones.cls`, `Form_FormConfiguracion` | F2 · Configuración de backend |
| C-HPS-S-3 | `Variables Globales.bas` (flags de operación) | F13 · Configuración runtime |
| C-HPS-S-4 | `Variables Globales.bas` (`EsUsuarioTecnicoCalculado`), `Form_Form0BDOpciones` | F1 · Arranque e identidad |
| C-HPS-S-5 | `Form_FormSolicitudAlta`, `Solicitud.cls`, `SolicitudOperaciones.cls`, `SolicitudFechas.cls`, `SolicitudFechasOperaciones.cls` | F3 · CRUD Solicitud |
| C-HPS-S-6 | `Form_FormSolicitudesAltaTraspasoDatos`, `Form_FormSolicitudesTraspasoFechas`, `Form_FormAdjuntaTraspasoONS`, `SolicitudOperaciones.RegistroEnvioTraspasoONS` | F5 · Traspasos a ONS |
| C-HPS-S-7 | `Form_FormAdjuntarExcelSolicitante`, `SolicitudOperaciones.RegistroEnvioExcel` | F6 · Adjuntos Excel del solicitante |
| C-HPS-S-8 | `Form_FormPlantillasHTML`, `Form_FormWeb`, `Form_FormCorreo`, `Form_FormEnvioCorreo`, `Form_FormCorreoPruebas` | F7 · Plantillas HTML y vistas web embebidas |
| C-HPS-S-9 | `Form_FormSolicitudRenovacion`, `Form_FormSolicitudRenovacionDatos`, `Form_FormSolicitudRenovacionFechas`, `Form_FormSolicitudesAltaRenovacionFechas` | F3, F4 · CRUD Solicitud y workflow regulatorio |
| C-HPS-S-10 | `Form_FormCambioTipo` | F3 · CRUD Solicitud |
| C-HPS-S-11 | `Justificacion.cls`, `JustificacionServicio.cls`, `Form_frmJustificacion`, `Form_frmBuscarJustificacion` | F8 · Justificaciones |
| C-HPS-S-12 | `Responsable.cls`, `ResponsableOperaciones.cls`, `Form_FormResponsable`, `Form_FormResponsablesGestion` | F9 · Responsables |
| C-HPS-S-13 | `MotivoHPS.cls`, `Form_FormMotivoHPS`, `Form_FormMotivosHPS` | F10 · Motivos HPS |
| C-HPS-S-14 | `Form_FormEntidades`, `Form_FormExpedientesBusqueda`, `Form_FormExpedienteDetalle` | F11 · Entidades y búsqueda de expedientes |
| C-HPS-S-15 | `Form_FormTareasTramitadorPendientes`, `constructor.bas` (once consultas de tareas) | F12 · Tareas pendientes del tramitador |
| C-HPS-S-16 | `LogGeneral.cls`, `LogGeneralOperaciones.cls` | sin candidato declarado |
| C-HPS-S-17 | `Correo.cls`, `CorreoOperaciones.cls`, `CorreoServicio.cls`, `Form_FormCorreo`, `Form_FormCorreosGestion`, `Form_FormEnvioCorreo`, `Form_FormCorreoPruebas` | F14 · Correos automáticos |
| C-HPS-S-18 | Sin módulo declarado: `TbSolicitudes.IDUsuarioHPS` y el flag `RegistroEnHPS` | F15 · Registro automático en HPS |

## Riesgos transversales

1. **Lanzadera es el origen de identidad y permisos.** Su PRD lo dice sin ambigüedad: es propietaria de las tablas de identidad que vinculan Expedientes, HPS, Condor y No Conformidades, y Gestión de Riesgos vincula además `TbAplicaciones`; un cambio de estructura o de claves afecta a seis de las ocho aplicaciones durante toda la coexistencia (`prd/coexistencia.md`, apartado 6). Lo confirman Brass (`getdbLanzadera` lee `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`), Gestión de Riesgos (10 llamadas repartidas en cuatro ficheros y tres tablas), Condor (`getdbLanzadera`, 7 llamadas) y Expedientes (EXP-CAP-043 a EXP-CAP-046 dependen del contrato de identidad de Lanzadera). HPS Solicitudes se suma con D-HPS-8, que deja abierta la consolidación de su identidad con la de Lanzadera.
2. **Expedientes es la referencia de expediente.** Condor compone el código de la solicitud a partir de un texto del expediente y de ahí sale el RAC (P-CON-2); Gestión de Riesgos busca expedientes por HTTP/JSON sin autenticación documentada (D132); No Conformidades busca expedientes con un `SELECT *` filtrado en VBA (CAP-NC-31) y HPS Solicitudes llega a ellos por la tabla vinculada `TbExpedientes` (C-HPS-S-14); y el propio Expedientes abre cuatro contratos de salida: HPS, AGEDYS, Gestión de Riesgos y No Conformidades (EXP-CAP-051 a EXP-CAP-054).
3. **No Conformidades precede a Condor en el orden de migración.** La verificación de la no conformidad antes de eliminar usa el código de la concesión, de modo que la relación es por código y no por clave foránea directa; el orden estricto es No Conformidades primero y Condor después (D95, P-CON-3). Gestión de Riesgos se enlaza a la vez por `TbRiesgosNC.IDNC` (F20) y Expedientes por EXP-CAP-054.
4. **El identificador es el punto de colisión entre Access y web.** Expedientes, Lanzadera, Condor, Brass, Gestión de Riesgos y No Conformidades calculan los identificadores con `Max + 1` sin bloqueo ni transacción. Mientras Access y web escriban a la vez sobre las mismas tablas, ambos pueden obtener el mismo valor; la fuente única de identificadores es una decisión pendiente en la coexistencia de cada aplicación.
5. **Las reglas de negocio que hoy solo aplica VBA necesitan una implementación.** Expedientes las enumera (validación `MotivoNoOK`, invariante de responsables, unicidad por consulta previa, comprobaciones de «en uso», permisos en controles); Brass llega a la misma conclusión: con dos clientes, cada regla necesita una sola implementación, y hasta que se decida, toda regla nueva se replica a mano en VBA. La web no puede omitirlas porque PostgreSQL no las aplica si no se declaran como restricciones.
6. **Hay secretos en el código de varias aplicaciones.** Condor tiene una contraseña como fallback en `GetPasswordDB` (D93) y Brass una contraseña real en tres sitios (D104, D-BRA-3); el patrón puede existir también en Lanzadera, Gestión de Riesgos, No Conformidades, HPS y HPS Solicitudes (D104 cross-cutting). HPS Solicitudes lo documenta: la contraseña del SMTP para `CDO.Message` vive como literal en `Variables Globales.bas` (coexistencia, apartado 6). Es un riesgo compartido y no exclusivo de una aplicación.
7. **La reversión de corte no está diseñada.** Condor lo registra de forma expresa (P-CON-4): si el destino falla a mitad de una transición, el estado de la solicitud queda en un backend y el log en el otro, y no hay operación que los vuelva a alinear. No Conformidades y HPS Solicitudes lo registran igual: la reversión del corte en la Lanzadera está por definir en ambas coexistencias (apartados 8 y 9). Sin reversión no se puede planificar la salida de ninguna aplicación.
8. **No Conformidades no tiene ninguna escritura de negocio en transacción.** Su coexistencia lo mide: 0 `BeginTrans` y 0 `CommitTrans` de DAO, y las transacciones que existen pertenecen a su caché, no al motor. El alta de una NC, el alta de una acción correctiva y el vínculo con riesgos son pasos múltiples sin atomicidad, y el identificador se calcula en cliente. Es el caso más severo de los siete estudios y el que más rápido invalida el resto del plan.
9. **HPS Solicitudes abre ficheros `.accdb` vinculados y escribe datos personales en claro.** Su identidad vive en las tablas vinculadas `TbUsuarios`, `TbUsuariosEntidades` y `TbUsuariosHistoricos` de `HPST.accdb`, y los expedientes en `TbExpedientes` de `Expedientes_datos.accdb` (D-HPS-8). Además, `TbSolicitudes` guarda ocho columnas con datos personales (D98), `TbLogsGeneral` escribe `DNI` y `Usuario` en plano (D-HPS-3) y la relación con los responsables es por `email` y no por identificador (D99). Consolidar la identidad es previa a cualquier corte.

## Pendientes y siguiente paso

- **ERD de origen.** Sigue pendiente en las ocho aplicaciones: la extracción de solo lectura con `dysflow-erd` depende del runner Windows del propietario (`D-EXP-4`, `D-CON-5`, `D-BRA-1`, `D-GR-4`). El ERD objetivo PostgreSQL global, en `erd-global-postgresql.md`, se consolida cuando estén los ocho ERD de origen; hasta entonces solo pueden proponerse vistas por aplicación.
- **Una aplicación en curso.** HPS (#1076) sigue con la fase 5 abierta y no tiene filas aquí. HPS Solicitudes (#1077) y No Conformidades (#1080) han entregado su clasificación de capacidades en `main` y ya tienen filas en esta matriz; la excepción `Nemotecnico` de No Conformidades entra con la fila de alta de NC de proyecto (CAP-NC-7).
- **Confianza de las seis aplicaciones sin vocabulario.** El PRD de Lanzadera, Condor, Brass, Gestión de Riesgos, HPS Solicitudes y No Conformidades clasifica el estado frente al estudio previo, pero no declara nivel de confianza por capacidad. Cuando esos PRD adopten el vocabulario del ledger, sus celdas dejarán de ir con `—`.
- **Capacidades sin candidato web declarado.** En No Conformidades (CAP-NC-3, CAP-NC-26, CAP-NC-35 a CAP-NC-38 y CAP-NC-39 a CAP-NC-45) y en HPS Solicitudes (C-HPS-S-16) ninguna feature de la épica las sostiene: sus celdas van con `sin candidato declarado` hasta que la épica las declare. No es una señal de que la capacidad se retire.
- **Siguiente paso.** Cruzar con `08-decisiones-y-preguntas-abiertas.md` cuando aparezcan divergencias o preguntas de mayor valor.

## Checklist

- [x] Cada fila cita la ficha de capacidad y la aplicación que la sostiene.
- [x] Las capacidades marcadas como `lógica a extraer` enlazan con `06-seguridad-y-trazabilidad.md` si tocan permisos.
- [x] Ningún componente web introduce todavía decisiones de stack.
