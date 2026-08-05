# 04 · Integraciones y operación — Procesos batch y automatizaciones

## Propósito

Inventario de los procesos batch, scripts, tareas programadas y automatizaciones que dan soporte al ecosistema (sincronizaciones, cargas, depuraciones, reporting nocturno, health checks). Describe el **qué** y el **scheduler unificado** que los soporta; no diseña la migración interna de cada job.

## Scheduler unificado (APROBADO)

- **Plataforma única** que sustituye los flags de tareas de Lanzadera y los mecanismos específicos por aplicación. Los jobs programados y los procesos batch se ejecutan a través del servicio compartido; el motor concreto (cron-like inicial; reemplazable a futuro) vive detrás de un port.
- **Gobernanza**: solo el administrador global configura y supervisa los jobs. La periodicidad, frecuencia y severidad mínima de alerta son decisiones del administrador global, no por aplicación.
- **Lógica de negocio en el módulo**: cada job de informe o de mantenimiento carga su lógica desde el módulo que lo origina. El scheduler no contiene reglas de negocio editables.
- **Inspección bajo demanda + programada**: el CLI del administrador global permite ejecutar un trabajo configurado al momento, además de la ejecución programada. La inspección programada solo mide calidad y salud, **no** ejecuta remediaciones.
- **Remediación con aprobación**: toda remediación se detiene y espera aprobación explícita del administrador global, incluso cuando la acción está clasificada como segura. Las acciones destructivas o de impacto masivo requieren confirmación adicional.

## Evidencia de automatizaciones legacy

### Scripts VBS (evidencia sanitizada)

- Los once scripts `.vbs` originales se trasladaron, previa sanitización, a `inputs/automatizaciones-legacy/raw/`. Esa carpeta está **ignorada por Git** (ver `.gitignore`) y constituye evidencia de sólo lectura para análisis de migración.
- `inputs/automatizaciones-legacy/README.md` documenta el flujo sin reproducir valores sensibles: entrypoint `script.vbs`, ciclos laborales, mapeo aplicación-script y listado de peligros conocidos (contraseña OLEDB compartida, host SMTP hardcoded, UNC internos y listas de distribución internas). Esos secretos **no** se exponen en este repositorio.
- El entrypoint orquesta el batch diario (lunes a viernes tras hora laboral) y, en cada ciclo, los mailers de tareas y de pendientes.
- Los dos scripts marcados como "huérfanos" no se invocan desde el entrypoint.

### Bucle de despacho de cola de correo (≈ 5 min)

El daemon externo invoca scripts de despacho cada cinco minutos aproximadamente para inspeccionar la cola por tabla y enviar mensajes pendientes. La responsabilidad se separa en dos a partir del producto web:

| Responsabilidad hoy | En el producto web |
|---|---|
| Programación de informes diarios | Scheduler unificado, configurado por el administrador global. |
| Despacho de correos pendientes | Servicio unificado de notificaciones, adaptador v1 de cola por tabla. |

### Flags de tareas y aplicación

Las marcas como `ParaTareasProgramadas` y los helpers de correo de Lanzadera se retiran como mecanismo de aplicación. Su semántica de negocio (qué informe, qué destinatarios) se conserva en el catálogo de jobs del módulo; la operativa (cuándo, con qué reintentos, qué alerta) la gobierna el scheduler compartido.

## Health checks e inspecciones programadas (APROBADO)

- **Bajo demanda**: el CLI del administrador global ejecuta inspecciones de salud puntuales.
- **Programado**: jobs cron-like miden calidad y salud. Solo miden; **no remedian**.
- **Configuración**: solo el administrador global configura destinatarios, checks activos, frecuencia y severidad mínima. Los administradores de aplicación no gestionan esta configuración operativa.
- **Destinatarios por aplicación**: cada aplicación expone un área de administración donde los administradores (global o de aplicación, según política del módulo) configuran los destinatarios de anomalías para ese módulo. Esto no contradice la configuración global de health checks; es una afinidad de routing.
- **Anomalías**: ante una anomalía se notifica a los destinatarios del módulo afectado. La remediación requiere aprobación explícita del administrador global.

## Buenas prácticas operativas

- **Confirmación explícita** para jobs destructivos o de impacto masivo, incluso con credencial autorizada. La confirmación liga el plan exacto (alcance, parámetros, nonce).
- **Secretos** fuera del plan del job: entran por variables de entorno o por un adapter de proveedor de secretos.
- **Idempotencia**: cada job documenta su comportamiento idempotente. La re-ejecución no debe romper el estado del sistema.
- **Logs canónicos**: cada ejecución emite logs estructurados con ID de correlación, actor, acción, objetivo, resultado, timestamp, módulo y contexto de error seguro.

## Decisiones aún no tomadas (ABIERTO)

- Catálogo definitivo de health checks (métricas, umbrales, severidades).
- Política exacta de reintentos y de escalado de alertas en el servicio unificado de notificaciones.
- Detalle del dashboard de operaciones de notificación (ver `informes-exportaciones-y-correo.md`).

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA para trazado de símbolos.
5. Engram como contexto histórico.

## Reglas de evidencia

- Cada automatización lleva su propósito declarado y su planificador.
- Los efectos sobre datos (escritura, borrado, archivo) se documentan con su tabla destino y, si aplica, el job que los gobierna.
- Las automatizaciones entre aplicaciones se marcan como `cross-app` y se referencian desde ambas fichas.
- APAP y APAP_WEB no aparecen.

## Checklist

- [ ] Cada job automatizado declara módulo, schedule y destinatarios.
- [ ] Los efectos sobre datos están citados, no asumidos.
- [ ] La política de remediación automática está alineada con `06-seguridad-y-trazabilidad.md`.

## Siguiente paso

Cruzar con `04-integraciones-y-operacion/informes-exportaciones-y-correo.md` para mantener coherencia entre scheduler y notificaciones, y con `tablas-vinculadas-y-backends.md` para mantener coherencia de rutas.
