# Issues propuestas — Lanzadera

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Issue:** #1074 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Unidades de trabajo que se derivan del PRD. **No se ha creado ninguna issue**: son propuestas para que el mantenedor las apruebe, las fusione con las existentes o las descarte. No se abre ninguna issue en el repositorio legacy de Lanzadera. Cada una debe publicarse con etiqueta `app/lanzadera`; las marcadas «transversal» además con `cross-cutting` y coordinadas con las demás aplicaciones.

## Reconciliación con el MVP web existente

El módulo [`app/src/modules/lanzadera`](../../../../app/src/modules/lanzadera) ya implementa una parte relevante. Lo que sigue compara con el legacy a nivel de capacidad; no es una revisión del código web.

| Capacidad legacy | Equivalente web (ficheros) | Diferencia observada |
|---|---|---|
| Autenticación y sesión (C-LAN-1) | `application/login.py`, `application/logout.py`, `delivery/http/auth_routes.py`, `auth_middleware.py` | La web fija una política de bloqueo explícita (umbral 5, duración 3 600 s por defecto, `domain/session.py`); en el legacy el contador de intentos no se incrementa y el bloqueo no caduca (H-LAN-1 a H-LAN-2) |
| Contraseñas (C-LAN-3) | `application/set_password.py`, `adapters/crypto/credential_hasher_argon2id.py`, repositorio de tokens de reseteo | Algoritmo distinto del legacy (que usa la API criptográfica de Windows); la migración de hashes heredados con rehash al primer login es una disposición aprobada y su implementación no se ha verificado en este estudio |
| Alta y gestión de usuarios (C-LAN-2) | `application/create_user.py`, `disable_user.py`, `delivery/http/admin_routes_users.py` | El legacy no guarda responsable, permiso de pruebas ni fecha de baja en el alta (H-LAN-14); la web debe decidir si los conserva |
| Catálogo de aplicaciones (C-LAN-4) | `application/create_app.py`, `update_app.py`, `disable_app.py`, `delivery/http/admin_routes_apps.py` | Alta de aplicación inalcanzable en el legacy (H-LAN-8) |
| Asignaciones y permisos (C-LAN-7, 8) | `application/assign_profile.py`, `revoke_assignment.py`, `adapters/persistence/repositories/assignment_repository_pg.py`; migración `0001` | El legacy tiene siete indicadores de rol por usuario y aplicación; la web usa perfiles con capacidades (`domain/legacy_role_map.py`) |
| Administrador global (C-LAN-9) | `application/grant_global_admin.py`, `revoke_global_admin.py`, `bootstrap_global_admins.py` | — |
| Conexiones (C-LAN-12) | `application/track_presence.py`, `get_connected_users.py`, migración `0008` | La web no recoge SSID ni ubicación, acorde con la disposición aprobada |
| Menús, lanzamiento y visibilidad (C-LAN-5, 6, 10) | `application/get_my_apps.py`, `list_effective_apps.py` | La web lista las aplicaciones efectivas del usuario; el lanzador de escritorio se retira por disposición aprobada |
| Formación (C-LAN-11) | sin equivalente | Retirada aprobada |

## Propuestas

| Id | Propuesta | Tipo | Transversal | Depende de | Solapa con |
|---|---|---|---|---|---|
| P-LAN-1 | Extraer del backend los metadatos de solo lectura (tablas, claves, índices, relaciones, nulabilidad) | verificación | no | máquina Windows del propietario con `dysflow` | — |
| P-LAN-2 | Verificar en el binario los hallazgos de autenticación y permisos | verificación | no | RC compilado y entorno aislado | — |
| P-LAN-3 | Fijar el contrato de identidad compartido con las otras aplicaciones (`TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbAplicaciones`) | decisión | sí | P-LAN-1 | PRD de Expedientes, HPS, Condor, No Conformidades y Gestión de Riesgos |
| P-LAN-4 | Verificar la estrategia de migración de credenciales (hashes heredados y reinicio) con datos sintéticos | verificación | no | P-LAN-1 | #625 (lectura del `.accdb` con contraseña, cerrada) |
| P-LAN-5 | Definir qué significa «lanzar» durante la coexistencia (lanzador de escritorio hasta que migren todas) | decisión | sí | — | — |
| P-LAN-6 | Fuente única de identificadores compartida por Access y la web | decisión | sí | P-LAN-3 | P-EXP-4 de Expedientes |
| P-LAN-7 | Aprobar la política objetivo de bloqueo y caducidad de contraseña | decisión | no | P-LAN-2 | — |
| P-LAN-8 | Disposición de los formularios huérfanos, de las tablas de formación y de las tablas técnicas del backend | decisión | no | P-LAN-1 | — |
| P-LAN-9 | Plan y pruebas de aceptación bidireccionales de inicio de sesión y permisos | pruebas | no | P-LAN-3 | — |
| P-LAN-10 | Incorporar al estudio previo las capacidades candidatas | documentación | no | — | — |

## Detalle

### P-LAN-1 · Extracción de metadatos
- **Evidencia:** el ERD de origen no se ha extraído (D-LAN-3 del [README](README.md)); la [matriz](matriz-campos.md) usa columnas del snapshot sin claves ni nulabilidad.
- **Criterios de aceptación:** metadatos de las 35 tablas del backend, extraídos por el propietario con `dysflow` en solo lectura, en un artefacto versionado; diferencias frente a `ERD/schema.sql` listadas.

### P-LAN-2 · Verificación de hallazgos de impacto alto
- **Evidencia:** [hallazgos del README](README.md#hallazgos-del-legacy): H-LAN-1 (administrador de Windows entra sin contraseña), H-LAN-2 (bloqueo que no caduca), H-LAN-3 y H-LAN-16 (clave de un solo uso en claro), H-LAN-12 (`Lanzar` sin comprobación de permisos), H-LAN-14 (alta de usuario sin clave inicial) y H-LAN-8 (alta de aplicación inalcanzable).
- **Criterios de aceptación:** por hallazgo, resultado reproducido o descartado; ninguna corrección en el legacy. Por su sensibilidad, los resultados de autenticación se comunican al propietario, no se publican con detalle.

### P-LAN-3 · Contrato de identidad compartido
- **Evidencia:** otras cuatro aplicaciones vinculan estas tablas del backend de Lanzadera ([coexistencia](coexistencia.md#6-identidad-permisos-y-credenciales)).
- **Criterios de aceptación:** estructura, claves y propietario de cada tabla acordados con todas las aplicaciones afectadas y reflejados en el ERD global; ningún cambio de estructura sin esa coordinación.

### P-LAN-4 · Credenciales
- **Criterios de aceptación:** con datos sintéticos, un hash heredado se verifica y se reemplaza por el de la web en el primer inicio de sesión correcto; ninguna credencial real se usa ni se registra; los parámetros del algoritmo heredado no se publican.

### P-LAN-5 · Significado de «lanzar»
- **Evidencia:** Lanzadera abre y modifica los `.accdb` de otras aplicaciones y los copia entre ubicaciones ([coexistencia](coexistencia.md#3-lo-específico-de-lanzadera-gestiona-los-frontends-de-las-demás-aplicaciones)); el lanzador de escritorio está aprobado para retirar.
- **Criterios de aceptación:** decisión escrita sobre qué hace la web mientras convivan aplicaciones de escritorio y web, incluida la reversión.

### P-LAN-6 · Identificadores
- **Evidencia:** `DameID` calcula `Max + 1` (12 llamadas en 7 ficheros, [coexistencia](coexistencia.md#5-recordsets-identidad-transacciones-y-tipos)); el esquema web usa otros tipos de identificador.
- **Criterios de aceptación:** una única fuente de identificadores y su correspondencia entre Access y web.

### P-LAN-7 · Política de bloqueo y caducidad
- **Evidencia:** la web fija umbral 5 y 3 600 s (`domain/session.py`); el legacy no aplica ninguno de los dos; las decisiones de ciclo de credencial viven en `docs/06-seguridad-y-trazabilidad.md`.
- **Criterios de aceptación:** política objetivo aprobada y aplicada igual en ambos lados mientras coexistan.

### P-LAN-8 · Disposición de huérfanos y tablas
- **Evidencia:** cinco formularios sin apertura desde `src` ([capacidades](capacidades.md)); formación retirada por disposición aprobada; tablas técnicas del backend sin uso en `src`.
- **Criterios de aceptación:** disposición por elemento (archivar, migrar o descartar) aprobada por el propietario, sin borrar datos.

### P-LAN-9 · Aceptación bidireccional
- **Criterios de aceptación:** las siete pruebas de la [coexistencia](coexistencia.md#8-pruebas-de-aceptación-bidireccionales-propuestas), en entorno aislado y nunca sobre producción.

### P-LAN-10 · Candidatas
- **Criterios de aceptación:** decisión sobre las tres candidatas de [capacidades](capacidades.md#capacidades-de-código-sin-entrada-propia-en-el-estudio).

## Pendientes

- Aprobar, fusionar o descartar cada propuesta antes de crear issues.
- Confirmar con las demás aplicaciones que P-LAN-3 y P-LAN-6 no duplican trabajo ya propuesto.
