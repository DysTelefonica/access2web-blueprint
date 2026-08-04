# 09 · Arquitectura objetivo y principios

## Propósito

Consolida las decisiones de producto y arquitectura **ya aprobadas** para el estudio de modernización del ecosistema legacy, junto con los baselines provisionales, las opciones futuras que permanecen abiertas y las preguntas que requieren respuesta. Sirve como referencia única para los siguientes lotes de discovery y para futuras fases SDD; **no** propone todavía proposal/spec/design/tasks concretos.

Las decisiones se han extraído de Engram (`project: access2web-blueprint`) con sus `topic_key`. Cada fila del registro consolidado lleva ese identificador para poder auditar el origen.

## Estado del descubrimiento

El discovery legacy (Fase 1 por lotes, ver `openspec/changes/legacy-ecosystem-discovery/exploration.md`) **está incompleto**. La topología general y la matriz inicial de dependencias son provisionales; las capacidades reales se obtendrán al ejecutar los Lotes 1 a 9. Mientras tanto, este documento fija la **forma objetivo** del producto web y sus principios, no la réplica exacta del legacy.

El legacy define el **suelo mínimo de capacidad de negocio**, no el objetivo de UX ni de implementación. La forma moderna debe preservar la intención y adaptarla a patrones web nativos.

## Ruta rápida

1. Plataforma única, modular y permission-aware.
2. Navegación app-first anidada; Lanzadera queda restringida a administración.
3. Hexagonal real en todos los módulos; web y CLI como adaptadores driving; PostgreSQL, cola de correo por tabla, object storage, scheduler, autenticación, caché y almacenamiento de archivo como adaptadores driven.
4. Persistencia objetivo: PostgreSQL compartido con aislamiento por esquema.
5. Notificaciones: servicio unificado; v1 solo email sobre la cola por tabla como adaptador transitorio.
6. Adjuntos: object storage S3-compatible detrás de un puerto; sin versionado de contenido; papelera con retención 30 días; restauración por el borrado o por administrador global.
7. CLI exclusivo para administrador global; reutiliza los mismos casos de uso y autorización que la web; sin secretos en argumentos ni en logs.
8. Logs canónicos industriales; retención por niveles configurable (90 días hot + 1 año total provisional).
9. Diagnóstico bajo demanda y programado; remediación siempre con aprobación explícita del administrador global.

## Forma de la plataforma

| Decisión | Estado | Origen |
|---|---|---|
| Producto único, modular y permission-aware: cada usuario ve solo los módulos para los que tiene permiso. | APROBADO | `architecture/target-platform-shape` |
| Navegación app-first anidada: barra global con módulos visibles; cada módulo expone su menú interno. El expediente no es punto de entrada. | APROBADO | `architecture/navigation-model` |
| Lanzadera queda restringida a administración de plataforma (usuarios y permisos por aplicación). Los usuarios normales se autentican una vez y solo ven sus módulos autorizados. | APROBADO | `architecture/identity-and-admin-module` |

## Arquitectura hexagonal global

| Decisión | Estado | Origen |
|---|---|---|
| Hexagonal es obligatorio en toda la plataforma y en cada módulo, no solo en autenticación. Las capacidades dependen de puertos, no de mecanismos de infraestructura. | APROBADO | `architecture/global-hexagonal-principle` |
| Adaptadores **driving**: interfaz web y CLI administrativo. | APROBADO | consolidación `ai-cli-first` + navegación |
| Adaptadores **driven** reemplazables: PostgreSQL, cola de correo por tabla, object storage, scheduler, proveedores de autenticación, caché y almacenamiento de archivo. | APROBADO | consolidación `global-hexagonal-principle`, `attachment-storage`, `authentication-ports-adapters`, `notification-delivery-adapter-v1`, `audit-retention`, `scheduled-health-checks`, `performance-and-cache` |
| Autenticación: adaptador inicial email/password migrado desde `Lanzadera_Datos.accdb` (contraseñas como hashes); un adaptador futuro e intercambiable podrá integrarse con SiteMinder y JWT unificado si la plataforma se aloja en OCP corporativo. SiteMinder y OCP son **opciones futuras**, no dependencias iniciales. | APROBADO (adaptador inicial) / FUTURO (SiteMinder/OCP) | `architecture/authentication-ports-adapters` |
| La cola de correo por tabla del legacy se mantiene como **adaptador transitorio** del servicio unificado de notificaciones; se sustituirá por la integración corporativa cuando IT defina su contrato. | APROBADO | `architecture/notification-delivery-adapter-v1` |
| El CLI nunca duplica lógica de negocio: invoca los mismos casos de uso y pasa por la misma autorización server-side que la web. | APROBADO | `architecture/ai-cli-first` |

## Persistencia y rendimiento

| Decisión | Estado | Origen |
|---|---|---|
| Persistencia objetivo preferida: una base de datos PostgreSQL compartida; cada módulo dueño de su esquema; servicios comunes en esquema(s) dedicado(s). | APROBADO (preferencia/dirección) | `architecture/target-database-topology` |
| Una base de datos física compartida no implica propiedad compartida: los límites de esquema y las reglas de acceso deben preservar el aislamiento hexagonal y modular. | APROBADO | `architecture/target-database-topology` |
| Rendimiento y caché deliberada como preocupación de producto y arquitectura de primer orden; se elimina la percepción de lentitud del legacy. | APROBADO | `architecture/performance-and-cache` |
| Tecnología concreta de caché, topología de despliegue, stack exacto y descomposición en servicios. | ABIERTO | — |

## Notificaciones

| Decisión | Estado | Origen |
|---|---|---|
| Un único servicio unificado de notificaciones compartido por todos los módulos. Los módulos aportan el disparador de negocio, destinatarios y mensaje semántico; canales, reintentos, plantillas, observabilidad e integración con proveedores quedan detrás del puerto. | APROBADO | `architecture/unified-notification-service` |
| Primera versión de release: **solo email** (paridad mínima con el canal de comunicación legacy exigido). Notificaciones in-app y otros canales quedan fuera del alcance inicial. | APROBADO | `product/notification-v1-scope` |
| Proveedor de email corporativo concreto. | ABIERTO (FUTURO considerado) | — |

## Adjuntos

| Decisión | Estado | Origen |
|---|---|---|
| Almacenamiento en object storage S3-compatible detrás de un puerto de almacenamiento; los módulos no dependen de APIs, semánticas de bucket ni URLs S3. | APROBADO (puerto y compatibilidad S3) / ABIERTO (proveedor final) | `architecture/attachment-storage` |
| Sin historial de versiones de contenido: solo se conserva el fichero activo por referencia de adjunto. | APROBADO | `product/attachment-versioning` |
| Los adjuntos borrados pasan a una papelera con retención limitada antes del borrado permanente. | APROBADO | `product/attachment-deletion` |
| Ventana de recuperación: **30 días**; pasado ese plazo, un proceso programado purga definitivamente el contenido y el estado recuperable. | APROBADO | `product/attachment-retention` |
| Durante la papelera solo pueden restaurar: la persona que borró el adjunto o un **administrador global** de plataforma. Otros usuarios con permiso de edición del registro padre **no** pueden restaurar. | APROBADO | `product/attachment-restore-authorization` |

## Autorización y roles

| Decisión | Estado | Origen |
|---|---|---|
| Existen dos scopes: **administrador global** de plataforma y **administrador de aplicación** limitado al módulo. Además, cada aplicación tiene una persona responsable. | APROBADO | `architecture/authorization-roles` |
| Las responsabilidades del administrador de aplicación se definen **por módulo**, no por un único contrato universal. Ejemplo: en Gestion_Riesgos el administrador de aplicación asigna qué usuarios trabajan en cada instancia de riesgo; otros módulos exponen permisos delegados distintos y algunos módulos pueden no exponer funciones de admin de aplicación en su versión inicial. | APROBADO | `architecture/module-admin-capabilities` |
| Solo el **administrador global** nombra administradores de aplicación, y solo para los módulos donde aplique administración delegada. Los nombramientos son acciones auditables independientes de las capacidades administrativas del módulo. | APROBADO | `architecture/module-admin-assignment` |
| Para la recuperación de adjuntos, el término "administrador" significa **administrador global** de plataforma. | APROBADO | `architecture/authorization-roles` + `product/attachment-restore-authorization` |

## CLI administrativo para IA

| Decisión | Estado | Origen |
|---|---|---|
| Cada módulo expone un CLI operable por un agente IA para ejecutar **cualquier** acción soportada por el módulo, con credenciales por variables de entorno. | APROBADO | `architecture/ai-cli-first` |
| El CLI es exclusivo del **administrador global** de plataforma. Administradores de aplicación, responsables y usuarios normales no reciben acceso CLI. | APROBADO | `architecture/cli-access-control` |
| Acciones destructivas o de impacto masivo invocadas por CLI requieren una **confirmación explícita adicional** aun con credencial de administrador global autorizada. | APROBADO | `architecture/cli-destructive-confirmation` |
| Los secretos nunca aparecen como argumentos ni se registran en logs; entran por variables de entorno o por adaptadores de proveedor de secretos reemplazables. | APROBADO | `architecture/ai-cli-first` |
| La autorización se aplica **server-side** y se audita en cada acción CLI; el CLI no es un superusuario compartido sin trazabilidad. | APROBADO | `architecture/cli-access-control` |

## Observabilidad y auditoría

| Decisión | Estado | Origen |
|---|---|---|
| Toda la plataforma emite logs estructurados canónicos de calidad industrial, con correlación de acciones de usuario/IA a través de web, CLI, módulos, servicios compartidos, adaptadores, colas y jobs. | APROBADO | `architecture/observability-audit-logs` |
| Eventos canónicos con: ID de correlación/traza, actor, acción, objetivo, resultado, timestamp, módulo y contexto de error seguro; credenciales y cargas sensibles redactadas. | APROBADO | `architecture/observability-audit-logs` |
| Retención por niveles configurable (no hard-coded): recientes en hot consultable; antiguos comprimidos y archivados en object storage S3-compatible a través de un puerto de archivo reemplazable. | APROBADO (mecanismo) | `architecture/audit-retention` |
| Baseline provisional: 90 días en hot consultable + 1 año de retención total, con archivo en object storage. | PROVISIONAL | `product/audit-retention-periods` |
| Periodos de retención definitivos por cumplimiento normativo o IT corporativa. | ABIERTO | — |

## Operaciones y salud

| Decisión | Estado | Origen |
|---|---|---|
| El CLI administrativo expone diagnósticos operativos explícitos: inspección de logs canónicos, salud actual, evaluación de corrección, detección de anomalías/degradación, explicación de evidencia y soporte de remediación proactiva. | APROBADO | `architecture/ai-operations-cli` |
| El CLI puede diagnosticar, correlacionar evidencia y proponer acciones correctivas, pero **siempre se detiene y espera aprobación explícita** del administrador global antes de ejecutar cualquier remediación, incluso en acciones clasificadas como seguras. | APROBADO | `architecture/ai-remediation-approval` |
| Inspecciones de salud disponibles **bajo demanda** (CLI) y **programadas** (cron-like); los jobs programados solo miden calidad y salud, no remedian automáticamente. | APROBADO | `architecture/scheduled-health-checks` |
| Cada aplicación expone un área de administración donde se configuran los destinatarios de anomalías de ese módulo. | APROBADO | `architecture/application-operations-settings` |
| Solo el **administrador global** configura health-checks (destinatarios, checks activos, frecuencia, severidad mínima de alerta). Los administradores de aplicación no gestionan esta configuración operativa. | APROBADO | `architecture/health-check-configuration-authorization` |

## Resultado del estudio

| Decisión | Estado | Origen |
|---|---|---|
| El estudio termina con un **roadmap global** + un **roadmap detallado y plan de implementación por herramienta**. | APROBADO (resultado objetivo) | `product/modernization-principles` |
| El comportamiento legacy marca el suelo mínimo de capacidad de negocio, no un requisito de paridad exacta de funcionalidad o UI. Las funcionalidades se preservan en intención y se adaptan a flujos web nativos. | APROBADO | `product/modernization-principles` |
| Decisiones de producto e implementación requieren cuestionamiento iterativo y validación explícita del usuario. | APROBADO | `product/modernization-principles` |

## Mapa rápido de decisiones

| # | Tema | Estado | Origen |
|---|---|---|---|
| D5 | Plataforma única modular permission-aware | APROBADO | `architecture/target-platform-shape` |
| D6 | Navegación app-first anidada | APROBADO | `architecture/navigation-model` |
| D7 | Lanzadera = solo administración | APROBADO | `architecture/identity-and-admin-module` |
| D8 | Hexagonal global | APROBADO | `architecture/global-hexagonal-principle` |
| D9 | Autenticación: adaptador inicial email/password | APROBADO | `architecture/authentication-ports-adapters` |
| D10 | SiteMinder/OCP/JWT unificado como adaptador futuro | FUTURO | `architecture/authentication-ports-adapters` |
| D11 | Notificación unificada como servicio compartido | APROBADO | `architecture/unified-notification-service` |
| D12 | Notificación v1 solo email | APROBADO | `product/notification-v1-scope` |
| D13 | Cola de correo por tabla como adaptador transitorio | APROBADO | `architecture/notification-delivery-adapter-v1` |
| D14 | PostgreSQL compartido con esquemas por módulo | APROBADO (dirección) | `architecture/target-database-topology` |
| D15 | Rendimiento y caché deliberada como prioridad de primer orden | APROBADO | `architecture/performance-and-cache` |
| D16 | Adjuntos en object storage S3-compatible detrás de puerto | APROBADO | `architecture/attachment-storage` |
| D17 | Adjuntos sin versionado de contenido | APROBADO | `product/attachment-versioning` |
| D18 | Adjuntos con papelera de retención limitada | APROBADO | `product/attachment-deletion` |
| D19 | Retención de papelera = 30 días | APROBADO | `product/attachment-retention` |
| D20 | Restauración de papelera por borrado o administrador global | APROBADO | `product/attachment-restore-authorization` |
| D21 | Roles: administrador global + administrador de aplicación por módulo | APROBADO | `architecture/authorization-roles` |
| D22 | Capacidades de admin de aplicación definidas por módulo | APROBADO | `architecture/module-admin-capabilities` |
| D23 | Solo administrador global nombra administradores de aplicación | APROBADO | `architecture/module-admin-assignment` |
| D24 | CLI operable por IA para todas las acciones de aplicación | APROBADO | `architecture/ai-cli-first` |
| D25 | CLI exclusivo para administrador global | APROBADO | `architecture/cli-access-control` |
| D26 | Confirmación explícita en acciones destructivas o masivas | APROBADO | `architecture/cli-destructive-confirmation` |
| D27 | Logs estructurados canónicos con correlación | APROBADO | `architecture/observability-audit-logs` |
| D28 | Retención por niveles configurable con archivo en object storage | APROBADO | `architecture/audit-retention` |
| D29 | Retención provisional 90 días hot + 1 año total | PROVISIONAL | `product/audit-retention-periods` |
| D30 | CLI soporta diagnósticos operativos proactivos | APROBADO | `architecture/ai-operations-cli` |
| D31 | Toda remediación requiere aprobación del administrador global | APROBADO | `architecture/ai-remediation-approval` |
| D32 | Inspecciones de salud bajo demanda + programadas | APROBADO | `architecture/scheduled-health-checks` |
| D33 | Destinatarios de anomalías configurables por aplicación | APROBADO | `architecture/application-operations-settings` |
| D34 | Solo administrador global configura health-checks | APROBADO | `architecture/health-check-configuration-authorization` |
| D35 | Resultado del estudio = roadmap global + roadmap y plan por herramienta | APROBADO | `product/modernization-principles` |

## Decisiones aún no tomadas (ABIERTO)

No se han decidido y **no se inventan** en este documento:

- Tecnología concreta de caché (el "qué" del adaptador de caché).
- Topología de despliegue (on-premise, nube corporativa, OCP u otro).
- Stack exacto de implementación (framework, librerías, runtime).
- Descomposición en servicios / monolitio modular / microservicios.
- Periodos definitivos de retención por cumplimiento normativo o políticas de IT.
- Integraciones corporativas concretas (correo, identidad, monitorización, etc.).
- Calidad y algoritmo de los hashes de contraseña del legacy (decide si se migran tal cual o se exige reset).
- Comportamiento exacto de "responsable de aplicación" (capacidades operativas fuera de la administración delegada).
- Forma final de los health-checks (catálogo concreto, métricas, umbrales).
- Estrategia de migración de datos desde los `.accdb` a PostgreSQL.
- Catálogo definitivo de funciones de admin de aplicación por módulo (más allá del ejemplo de Gestion_Riesgos).

## Reglas del documento

- Este fichero no sustituye a `08-decisiones-y-preguntas-abiertas.md`: lo complementa con el detalle aprobado por el usuario.
- Toda mención a SiteMinder, OCP, Redis, S3, frameworks o colas corporativas se trata como **opción futura** o **abierto**, nunca como decisión final.
- El legacy no es paridad de UX ni de implementación: es suelo mínimo de capacidad de negocio.
- Los IDs de Engram (`topic_key`) son el anclaje autoritativo; cualquier cambio futuro debe actualizar simultáneamente este documento y la observación correspondiente.
- APAP y APAP_WEB no aparecen ni se mencionan en este fichero.
- Documentación en castellano de España, registro técnico-profesional.

## Checklist

- [ ] Cada nueva decisión aprobada por el usuario añade una fila al mapa rápido (D36, D37, …) y actualiza su sección temática.
- [ ] Ningún elemento se mueve de ABIERTO/FUTURO a APROBADO sin confirmación explícita.
- [ ] Los baselines PROVISIONALes se marcan con su condición de revisión.
- [ ] Las referencias a SiteMinder, OCP, Redis, S3, frameworks o colas corporativas se mantienen como FUTURO/ABIERTO.

## Siguiente paso

Cruzar este documento con `08-decisiones-y-preguntas-abiertas.md` para incorporar las nuevas decisiones y preguntas derivadas. Las decisiones de stack, despliegue y plazos quedan para una fase SDD posterior; el siguiente lote de discovery (Lote 1 – Lanzadera) no se bloquea con ellas.
