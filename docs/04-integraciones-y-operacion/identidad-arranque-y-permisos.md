# 04 · Integraciones y operación — Identidad, arranque y permisos

## Propósito

Recoge los patrones compartidos de arranque, autenticación y resolución de permisos en el producto objetivo: contrato de identidad, resolución de permisos por módulos, comportamiento por capacidades y trazabilidad. Distingue el camino objetivo (post-Lote 1) del inventario legacy que sigue vigente hasta la migración.

## Modelo objetivo (producto web)

### Identidad y sesión

- **Adaptador inicial (APROBADO)**: email/password migrado desde `Lanzadera_Datos.accdb`. Las contraseñas heredadas se conservan como hashes y se verifica con un adapter de verificación legacy versionada. Los campos en claro y el histórico de contraseñas no se migran como secretos.
- **Rehash transparente (APROBADO)**: tras el primer login exitoso con un hash heredado, el adaptador calcula el hash con la política moderna y sustituye el registro; la operación es atómica, idempotente y auditable. El usuario no realiza ninguna acción.
- **Lockout (APROBADO)**: el umbral de intentos fallidos es configurable por el administrador global, con valor por defecto cinco. Alcanzado el umbral, la cuenta queda bloqueada durante una hora por defecto. El administrador global puede desbloquear antes.
- **Notificación de lockout (APROBADO)**: cada bloqueo dispara un aviso por el servicio unificado de notificaciones a los administradores globales configurados. El aviso contiene contexto de diagnóstico seguro y un identificador de correlación, nunca contraseñas, hashes, tokens o datos personales innecesarios.
- **Caducidad de contraseña (APROBADO)**: el periodo de caducidad es configurable globalmente; el administrador global puede eximir a un usuario concreto. La política puede representar "sin caducidad periódica". Los cambios de política se auditan y aplican con efectos definidos para credenciales existentes.
- **Activación y desactivación (APROBADO)**: solo el administrador global registra usuarios, los da de baja y crea roles por aplicación. Los administradores de aplicación asignan usuarios activos a roles existentes dentro de su módulo.
- **Notificación de cambios de permiso (APROBADO)**: el cambio de acceso aplica de inmediato; el correo de resumen se envía solo cuando un administrador autorizado lo dispara explícitamente.
- **Suplantación para pruebas (APROBADO)**: solo el administrador global inicia sesión suplantada. Ámbito previsto: pruebas/UAT. La sesión muestra la doble identidad y registra auditoría completa con actor real y persona impersonada.

### Permisos, capabilities y vistas

- **Scopes administrativos (APROBADO)**: administrador global (cross-platform) y administrador de aplicación (limitado al módulo). Cada aplicación tiene una persona responsable separada de la administración.
- **Capabilities declaradas por módulo (APROBADO)**: cada módulo declara las **capacidades estables** que ofrece (p. ej. `risk.assign`, `nc.sign-off`, `expediente.manage`). Los **grupos de capabilities** se usan para componer roles verificables.
- **Políticas contextuales en código (APROBADO)**: las políticas que dependen de estado, recurso o condición de negocio se evalúan en el código del módulo, no como strings en el catálogo de capabilities. Capabilities y políticas coexisten; capabilities son el "qué soy" y la política contextual es el "qué puedo hacer aquí y ahora".
- **Backend autoritativo, UI derivada (APROBADO)**: la UI consume el endpoint de capabilities/políticas y no enumera acciones por nombre de rol. El backend rechaza operaciones no autorizadas aunque la UI las muestre por error.
- **Variantes de UI por módulo (APROBADO)**: cada módulo decide por sí mismo si usa una vista adaptativa o varias vistas especializadas. Vista única por defecto; especializada cuando existan diferencias materiales de flujo (ejemplo: Expedientes tiene gestión avanzada y vista técnica simplificada; la selección se guía por capabilities/políticas efectivas, no por nombre de rol).

### Arranque de aplicación (target)

1. Identidad (correo electrónico o suplantación validada) entra al shell unificado.
2. El shell consulta `id-origen/permisos` por capability para la sesión efectiva.
3. El módulo destino recibe el contexto de sesión con el conjunto de capabilities/políticas resuelto.
4. Las vistas y los endpoints se proyectan según ese conjunto. El backend aplica la última verificación.
5. Los eventos canónicos de `startup`, `module-open`, `deny` y `deny-fallback` quedan registrados.

## Evidencia legacy de arranque

La cadena observada al iniciar cada aplicación legacy (identidad → usuario → `IDAplicacion` → permisos → normalización → precedencia → perfil efectivo → primera vista → rechazo/fallback/error) está consolidada en [`docs/06-autorizacion-legacy-matriz.md`](../06-autorizacion-legacy-matriz.md). Sigue siendo referencia mientras conviven los binarios Access; el modelo objetivo no clona estos patrones, los reemplaza por servicios centralizados y capabilities declaradas.

Las decisiones que aplican uniformemente a las ocho aplicaciones legacy están en `docs/09-arquitectura-objetivo-y-principios.md`. Las excepciones operativas (Contraseña plana, BCC fijo, `PassIncialPlana`, `EjecucionEnOficina`, etc.) se tratan en `03-aplicaciones/lanzadera/integrations-security.md` y en las fichas individuales por aplicación.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. Inspección Dysflow solo lectura.
3. CodeGraph-VBA para trazado de símbolos.
4. Engram como contexto histórico, nunca como verdad actual.

## Reglas de evidencia

- Los permisos se documentan por capability/tabla origen; no por nombre de rol.
- El patrón de arranque se describe como familia, no como implementación única.
- Las cuentas técnicas y excepciones se etiquetan como tales.
- APAP y APAP_WEB no aparecen.

## Checklist

- [ ] Cada capability/política lleva módulo y caso de uso que la invoca.
- [ ] Los roles se referencian al catálogo de capacidades, no se duplican.
- [ ] Las decisiones de seguridad transversal (lockout, activación, suplantación) coinciden con `06-seguridad-y-trazabilidad.md`.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/lanzadera-identidad-permisos.md` para evitar duplicación de disposiciones sobre el rol transversal de Lanzadera.
