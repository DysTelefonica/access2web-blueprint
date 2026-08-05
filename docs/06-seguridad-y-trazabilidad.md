# 06 · Seguridad y trazabilidad

## Propósito

Documenta el modelo de seguridad del producto objetivo (identidad y ciclo de credencial, autenticación, autorización, separación de entornos, manejo de secretos, observabilidad y auditoría) y la trazabilidad mínima de cada afirmación. Las decisiones aquí recogidas son **APROBADAS**, **PROVISIONALES**, **FUTURAS** o **ABIERTO**, con su `topic_key` de Engram como anclaje autoritativo. Este documento se complementa con `09-arquitectura-objetivo-y-principios.md`, que consolida el mapa global de decisiones.

## Modelo de identidad y autenticación

- **Adaptador inicial (APROBADO)**: email/password migrado desde `Lanzadera_Datos.accdb`. Los hashes heredados se conservan y se verifican con un **adapter de verificación legacy versionada**. Los campos con credenciales en claro y el histórico de contraseñas no se migran como secretos.
- **Rehash transparente (APROBADO)**: tras el primer login exitoso con un hash heredado, el adaptador calcula el hash con la política moderna y sustituye el registro de forma atómica e idempotente. El usuario no realiza ninguna acción; la operación audita el algoritmo nuevo, no el secreto.
- **Adaptador futuro (FUTURO)**: integración corporativa con SiteMinder y JWT unificado si la plataforma se aloja en OCP corporativo. SiteMinder y OCP son **opciones futuras**, no dependencias iniciales.
- **Hexagonal obligatorio**: la autenticación se invoca siempre a través de un port; los módulos no conocen el mecanismo concreto.

## Ciclo de credencial

- **Caducidad (APROBADO)**: el administrador global configura el periodo de caducidad de contraseña. La política puede representar "sin caducidad periódica". Puede eximir a un usuario concreto.
- **Lockout (APROBADO)**:
  - **Umbral configurable** por el administrador global, con valor por defecto **cinco** intentos fallidos.
  - **Duración de una hora** por defecto; el administrador global puede **desbloquear manualmente** antes.
  - La política de reintentos y la ventana de reset del contador son detalles abiertos del catálogo de health-checks / políticas de auth.
- **Notificación de bloqueo (APROBADO)**: cada bloqueo dispara un aviso por el servicio unificado de notificaciones a los administradores globales configurados. El aviso lleva contexto de diagnóstico seguro e identificador de correlación; nunca contraseñas, hashes, tokens o datos personales innecesarios.
- **Activación y desactivación (APROBADO)**: solo el administrador global registra usuarios, los da de baja y crea roles por aplicación. El resto del ciclo (asignación de usuarios activos a roles existentes) puede ser de un administrador de aplicación.
- **Aviso de bienvenida y aviso manual de cambios de permiso (APROBADO)**: ningún cambio de permisos dispara correo automático. El envío lo inicia explícitamente un administrador autorizado con plantilla y destinatarios confirmados; el cambio aplica de inmediato, solo la comunicación es manual.

## Suplantación para pruebas (APROBADO)

- **Solo administrador global** inicia sesión suplantada.
- Ámbito previsto: pruebas/UAT.
- La sesión muestra la **doble identidad** (actor real y persona impersonada) de forma visible.
- Auditoría completa: iniciador, persona impersonada, parámetros, marcas temporales, módulo y resultado.

## Jerarquía de administración

- **Administrador global** (cross-platform): autoridad de plataforma. Nombra administradores de aplicación, configura health checks, opera el CLI administrativo.
- **Administrador de aplicación** (limitado al módulo): nombrado solo donde aplique administración delegada. Sus capacidades concretas las define cada módulo; las funciones de admin global sensibles (registro/baja de usuarios, creación de roles, configuración de health checks, configuración operativa del scheduler) **no** delegan.
- **Persona responsable** de cada aplicación: figura separada del administrador de aplicación, sin capacidades automáticas derivadas.

## Capabilities, políticas y autorización

- **Capabilities declaradas por módulo (APROBADO)**: cada módulo declara las **capacidades estables** que ofrece. Los **grupos de capabilities** componen roles verificables.
- **Políticas contextuales en código (APROBADO)**: las reglas que dependen de estado, recurso o condición de negocio se evalúan en el código del módulo, no en el catálogo. Capabilities y políticas coexisten: capabilities son el "qué soy" y la política es el "qué puedo hacer aquí y ahora".
- **Backend autoritativo, UI derivada (APROBADO)**: la UI consume el endpoint de capabilities/políticas. El backend rechaza operaciones no autorizadas aunque la UI las muestre.
- **Vistas por módulo (APROBADO)**: cada módulo decide si usa una vista adaptativa o varias especializadas. Por defecto, vista única; especializada cuando hay diferencias materiales (ejemplo: Expedientes tiene gestión avanzada para mantenimiento y vista técnica simplificada). La selección se guía por capabilities/políticas, no por nombre de rol.

## CLI administrativo

- **Acceso restringido** exclusivamente al administrador global. Administradores de aplicación, responsables y usuarios normales **no** tienen acceso CLI.
- **Mismas reglas que la web**: el CLI invoca los mismos casos de uso y pasa por la misma autorización server-side. No hay duplicación de lógica de negocio.
- **Confirmación explícita** para acciones destructivas o de impacto masivo, aun con credencial autorizada. La confirmación liga el plan exacto (alcance, parámetros, nonce).
- **Secretos**: nunca como argumento de comando ni en logs. Entrada por variables de entorno o por un adapter de proveedor de secretos reemplazable.
- **Remediación** siempre con aprobación explícita del administrador global, incluso para acciones clasificadas como seguras.

## Observabilidad y auditoría

- **Logs canónicos estructurados** de calidad industrial en toda la plataforma (web, CLI, módulos, servicios compartidos, adaptadores, colas y jobs). Eventos con ID de correlación/traza, actor, acción, objetivo, resultado, timestamp, módulo y contexto de error seguro.
- **Redacción**: credenciales y cargas sensibles nunca se registran en logs.
- **Retención por niveles configurable**: recientes en hot consultable; antiguos comprimidos y archivados en object storage S3-compatible a través de un puerto de archivo reemplazable.
- **Baseline provisional (PROVISIONAL)**: **90 días en hot consultable + 1 año de retención total**, archivo posterior vía adapter de object storage. Sujeto a revisión cuando IT o cumplimiento definan los periodos definitivos.
- **Misma retención** aplica a artefactos de informe generados y a evidencia de entrega de informes (ver `04-integraciones-y-operacion/informes-exportaciones-y-correo.md`).

## UAT, releases y excepciones

- **Gobernanza UAT solo administrador global (APROBADO)**: el administrador global define quién participa, qué perfil recibe cada participante y la configuración de acceso del ciclo. Los administradores de aplicación no configuran participación en UAT.
- **Asignación explícita por ciclo (APROBADO)**: cada ciclo UAT declara explícitamente sus participantes y perfil de aplicación. UAT **no** espeja automáticamente el acceso de producción.
- **Excepciones auditadas (APROBADO)**: un administrador global puede liberar con casos UAT fallidos o sin UAT ejecutado cuando las circunstancias lo justifiquen. Cada excepción es explícita, razonada, atribuible, sellada temporalmente y conservada con la evidencia del release.
- **Excepciones visibles para usuarios (APROBADO)**: cuando un release publica con casos UAT fallidos o sin UAT, la excepción y su justificación aparecen en el historial de cambios al alcance del usuario.
- **Diseño UAT detallado (ABIERTO)**: workflow, visibilidad, entorno, aprobaciones y promoción se diseñan en fase posterior.

## Manejo de secretos y configuración

- Ningún secreto aparece en documentación ni en logs; solo se describe dónde se almacena (variables de entorno o adapter de secretos).
- Las configuraciones Dysflow del legacy no se mutan en esta fase (ver `00-alcance-y-evidencia.md`).
- Las cuentas técnicas se etiquetan como tales y nunca se mezclan con cuentas personales.

## Separación de entornos

- Los roles por entorno (desarrollo, staging, producción) se documentan por separado; no se mezclan sin marcarlo.
- El discovery opera en modo solo lectura sobre checkout `00_main` y Dysflow.
- La navegación UAT/producción simultánea está **abierta** como diseño posterior (`topic open/dual-environment-navigation`).

## Decisiones aún no tomadas (ABIERTO)

- Periodos definitivos de retención por cumplimiento normativo o política corporativa de IT.
- Estrategia concreta de credenciales operativas para el CLI (gestión de secretos, rotación, ámbito detallado).
- Diseño exacto del dashboard de operaciones de notificación (UX, filtros, acceso a contenido sensible, umbrales de alerta, acciones operativas).
- Catálogo definitivo de health checks (métricas, umbrales, severidades).
- Forma final de la navegación dual UAT/producción cuando aplique.
- Forma final de la política de reintentos y de escalado de alertas en el servicio unificado de notificaciones.

## Fuentes de autoridad

1. Decisiones en Engram (`project: access2web-blueprint`, `topic_key` indicados en `09-arquitectura-objetivo-y-principios.md`).
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación para contrastar intención legacy.
3. `C:\00repos\codigo\<app>\00_main` por aplicación para comportamiento observable.
4. Dysflow en modo solo lectura como validación puntual.
5. Engram como contexto histórico, nunca como fuente única.

## Reglas de evidencia

- Cada afirmación lleva etiqueta `APROBADO`, `PROVISIONAL`, `FUTURO` o `ABIERTO`, con origen (`topic_key` de Engram) y fecha.
- Las decisiones tomadas por el usuario son autoridad; los placeholders heredados se reemplazan por decisión o se retiran.
- APAP y APAP_WEB no aparecen en este documento.

## Checklist

- [ ] Cada rol y permiso cita su origen (Engram `topic_key`) y su estado.
- [ ] Los secretos se describen sin mostrarse.
- [ ] Cada baseline provisional lleva su condición de revisión.
- [ ] Las opciones futuras se mantienen etiquetadas como `FUTURO`, nunca como `APROBADO`.

## Siguiente paso

Cruzar con `09-arquitectura-objetivo-y-principios.md` para verificar coherencia del mapa de decisiones, y con `02-topologia-ecosistema/lanzadera-identidad-permisos.md` cuando se ejecute el Lote 2 para no duplicar el detalle legacy.
