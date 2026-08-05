# 04 · Integraciones y operación — Informes, exportaciones y correo

## Propósito

Catálogo operativo del servicio unificado de notificaciones (email v1), del scheduler de informes, del manejo de artefactos de informe y del comportamiento manual/on-demand. No introduce capacidades funcionales — esas viven en `05-capacidades/`; aquí solo se documentan los **puntos de salida** observables y la plataforma que los soporta.

## Servicio unificado de notificaciones (APROBADO)

- **Arquitectura**: servicio compartido único, invocado por todos los módulos a través de un port. Los módulos aportan el disparador de negocio, destinatarios y mensaje semántico. Plantillas, reintentos, observabilidad e integración con proveedores quedan detrás del port.
- **v1 solo email (APROBADO)**: paridad mínima con el canal de comunicación legacy. Notificaciones in-app y otros canales quedan fuera del alcance inicial.
- **Proveedor de email corporativo (ABIERTO)**: la cola por tabla es el adaptador transitorio hasta que IT confirme el contrato final. No se prefija SMTP ni hosts.
- **Adaptador transitorio de cola por tabla (APROBADO)**: la cola por tabla del legacy se mantiene como primer adaptador del servicio. El adaptador escribe a través del port, por lo que se sustituye por la integración corporativa sin tocar los módulos.

### Evidencia del patrón legacy (queue + dispatcher 5 min)

El descubrimiento del Lote 1 confirmó el patrón extremo a extremo: las aplicaciones/jobs insertan registros de correo pendientes, y el bucle externo de orquestación invoca scripts de despacho aproximadamente **cada cinco minutos** para inspeccionar la cola y enviar los mensajes pendientes. La transición del legacy a la web debe separar dos responsabilidades hoy acopladas en un mismo daemon:

- Programación de informes (scheduler).
- Despacho de la cola de email (notification dispatcher).

Ambas pasarán a estar detrás de sus puertos respectivos; la cola por tabla actúa como adapter v1 mientras conviven los binarios Access.

## Scheduler unificado de informes y reportes (APROBADO)

- **Plataforma única**: scheduler compartido para jobs programados y procesos batch; el adaptador (cron-like inicial, reemplazable) queda detrás de un port.
- **Gobernanza**: solo el administrador global configura y supervisa los jobs; cada módulo conserva la lógica de negocio del informe.
- **Configuración por separado del contenido**: el administrador global configura horarios y destinatarios desde el panel de operaciones; el contenido del informe, su selección de datos y reglas de negocio viven en el módulo que lo origina.
- **Catálogo de jobs**: incluye informes diarios/semanales/mensuales de cada módulo, jobs de salud y jobs de mantenimiento del adaptador.

### Ejecución de informes

| Modo | Comportamiento | Quién |
|---|---|---|
| **Programado** | Modo habitual. Ejecuta la lógica del informe, adjunta el artefacto, lo envía a los destinatarios configurados y registra la entrega. | Automático |
| **Manual bajo demanda** | El administrador global lanza un trabajo configurado. Identifica al iniciador, parámetros, destinatarios objetivo, reintentos e identificadores de correlación. El envío sigue requiriendo confirmación explícita cuando es destructivo o de impacto masivo. | Administrador global |
| **Vista previa sin envío** | Genera el artefacto exacto del informe sin seleccionar destinatarios, sin encolar email ni enviar. Recomendado cuando el contenido es incierto o nuevo. | Administrador global |
| **Generar y enviar directamente** | El administrador global puede saltar la vista previa para casos urgentes o conocidos. La acción muestra destinatarios y parámetros antes de la confirmación explícita. | Administrador global |

Los informes generados se almacenan en el port de object storage, con checksum y referencia inmutable a la ejecución. Cada artefacto queda ligado a su ejecución y entrega.

### Auditoría de entrega

Para cada ejecución (programada o manual) se registra:

- Identificador de correlación / traza.
- Informe ejecutado y parámetros.
- Lista exacta de destinatarios.
- Resultado de la entrega (exitoso/fallido), reintentos y error seguro.
- Referencia al artefacto generado (id, checksum y origen).

La auditoría es del administrador global. La notificación de bloqueos de cuenta (ver `identidad-arranque-y-permisos.md`) sigue el mismo camino de logs canónicos.

### Retención de artefactos y evidencia de informes

- **Política común (PROVISIONAL)**: 90 días en hot consultable + un año de retención total, archivo posterior vía adapter de object storage. Misma línea base que la auditoría de plataforma.
- Los periodos definitivos dependen de cumplimiento normativo o política corporativa de IT; **ABIERTO**.

## Operaciones manuales y automatizaciones de notificación

- **Aviso manual de cambios de permiso** (APROBADO): el envío lo dispara un administrador autorizado, no es automático.
- **Aviso de bloqueo de cuenta** (APROBADO): disparo automático por la política de lockout hacia el servicio unificado; el contenido evita secretos.
- **Aviso de bienvenida al activar usuario** (APROBADO): disparo explícito por el administrador global.

## Evidencia legacy observada

- `Lanzadera` mantiene la cola en `TbUsuariosCorreosEnvio` (con referencias a `TbCorreosEnviados`); la tabla sigue el contrato "pendiente → enviado" que el dispatcher externo consume.
- Los scripts legacy de despacho (`EnviarCorreoTareas.vbs`, `EnviarCorreoNoEnviado.vbs` y otros) viven fuera de este repositorio y de `inputs/automatizaciones-legacy/raw/`; el READMÉ sanitizado en `inputs/automatizaciones-legacy/README.md` describe el flujo sin reproducir valores sensibles.
- No se identificaron informes Access exportables en el frontend de Lanzadera (`export_queries` devolvió cero consultas); no se demostró ejecución batch, exportación Excel, macros ni API externa.

## Dashboard de operaciones de notificación (APROBADO dirección, diseño detallado ABIERTO)

- **Aprobado**: la plataforma ofrece un dashboard para el administrador global con salud de la cola, estadísticas de entrega/fallo y reintentos controlados.
- **ABIERTO** (diseño detallado): UX exacta, filtros, acceso a contenido sensible, política de retención, umbrales de alerta y acciones operativas. La discusión posterior requerirá su propio lote.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA para trazado de símbolos.
5. Engram como contexto histórico.

## Reglas de evidencia

- Cada informe, exportación o envío de correo cita módulo y disparador (scheduled, batch, manual).
- Los adjuntos se describen por nombre y origen; no se exponen credenciales SMTP, hosts ni direcciones internas.
- Las automatizaciones se cruzan con `procesos-batch-y-automatizaciones.md`.
- APAP y APAP_WEB no aparecen.

## Checklist

- [ ] Cada salida lleva disparador y evidencia.
- [ ] Los informes de auditoría de informes comparten política con la auditoría general.
- [ ] El dashboard operativo se marca como dirección aprobada y diseño ABIERTO.

## Siguiente paso

Cruzar con `procesos-batch-y-automatizaciones.md` para mantener coherencia de jobs y puntos de salida, y con `08-decisiones-y-preguntas-abiertas.md` para mantener vivos los puntos ABIERTOS.
