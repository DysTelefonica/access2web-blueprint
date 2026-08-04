# 06 · Seguridad y trazabilidad

## Propósito

Documenta el modelo de seguridad del producto objetivo (identidad, autenticación, autorización, separación de entornos, manejo de secretos, observabilidad y auditoría) y la trazabilidad mínima de cada afirmación. Las decisiones aquí recogidas son **APROBADAS** por el usuario o **PROVISIONALES**; las opciones aún abiertas se listan al final.

Este documento se complementa con `09-arquitectura-objetivo-y-principios.md`, que consolida todas las decisiones de producto y arquitectura. Las decisiones de seguridad viven aquí en su forma detallada; el mapa global con su `topic_key` de Engram está en el documento 09.

## Modelo de identidad y autenticación

- **Adaptador inicial (APROBADO)**: email/password migrado desde `Lanzadera_Datos.accdb`, donde las contraseñas se almacenan como hashes. La idoneidad del algoritmo de hash del legacy se evalúa antes de decidir si se migra tal cual o se exige reset de credenciales.
- **Adaptador futuro (FUTURO)**: SiteMinder corporativo y login JWT unificado, integrado si la plataforma se aloja en OCP corporativo. SiteMinder y OCP son opciones, no dependencias iniciales.
- **Hexagonal obligatorio**: la autenticación se invoca siempre a través de un puerto; los módulos no conocen el mecanismo concreto.

## Roles y autorización

- **Administrador global** de plataforma: autoridad cross-platform (nombrar administradores de aplicación, configurar health-checks, operar el CLI).
- **Administrador de aplicación**: rol limitado al módulo, nombrado por el administrador global solo donde aplique administración delegada.
- **Persona responsable** de cada aplicación: figura separada del admin de aplicación, sin capacidades automáticas derivadas.
- **Capacidades de admin de aplicación por módulo**: cada módulo define sus propias funciones delegadas. Ejemplo: Gestion_Riesgos permite asignar qué usuarios trabajan en cada instancia de riesgo. No existe un contrato universal.

## CLI administrativo

- **Acceso restringido** exclusivamente al administrador global. Administradores de aplicación, responsables y usuarios normales **no** tienen acceso CLI.
- **Mismas reglas que la web**: el CLI invoca los mismos casos de uso y pasa por la misma autorización server-side. No hay duplicación de lógica de negocio.
- **Confirmación explícita** para acciones destructivas o de impacto masivo, aun con credencial autorizada. La confirmación se liga al plan exacto (alcance, parámetros, nonce), no a un "sí" genérico.
- **Secretos**: nunca como argumento de comando ni en logs. Entrada por variables de entorno o por un adaptador de proveedor de secretos reemplazable.

## Observabilidad y auditoría

- **Logs canónicos estructurados** de calidad industrial en toda la plataforma (web, CLI, módulos, servicios compartidos, adaptadores, colas y jobs). Eventos con ID de correlación/traza, actor, acción, objetivo, resultado, timestamp, módulo y contexto de error seguro.
- **Redacción**: credenciales y cargas sensibles nunca se registran en logs.
- **Retención por niveles configurable**: recientes en hot consultable; antiguos comprimidos y archivados en object storage S3-compatible a través de un puerto de archivo reemplazable.
- **Baseline provisional (PROVISIONAL)**: 90 días en hot consultable + 1 año de retención total. Sujeto a revisión cuando IT o cumplimiento definan los periodos definitivos.

## Manejo de secretos y configuración

- Ningún secreto aparece en documentación ni en logs; solo se describe dónde se almacena (variables de entorno o adaptador de secretos).
- Las configuraciones Dysflow del legacy no se mutan en esta fase (ver `00-alcance-y-evidencia.md`).
- Las cuentas técnicas se etiquetan como tales y nunca se mezclan con cuentas personales.

## Separación de entornos

- Los roles por entorno (desarrollo, staging, producción) se documentan por separado; no se mezclan sin marcarlo.
- El discovery opera en modo solo lectura sobre checkout `00_main` y Dysflow.

## Decisiones aún no tomadas (ABIERTO)

- Periodos definitivos de retención por cumplimiento normativo o política corporativa de IT.
- Calidad y algoritmo de los hashes de contraseña heredados (¿se migran tal cual o se fuerza reset?).
- Estrategia concreta de credenciales operativas para el CLI (gestión de secretos, rotación, ámbito).

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

Cruzar con `09-arquitectura-objetivo-y-principios.md` para verificar coherencia del mapa de decisiones, y con `02-topologia-ecosistema/lanzadera-identidad-permisos.md` cuando se ejecute el Lote 1 para no duplicar el detalle legacy.
