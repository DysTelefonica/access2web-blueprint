# Fase 1. Plan de aplicación del saneamiento en producción

> **Épica:** [#1545](../epica-coexistencia-postgresql.md#fase-1-saneamiento-en-access) · **Estado:** propuesta del 2026-11-24, pendiente de validar por el propietario · **Etiqueta de las issues:** `saneamiento:coexistencia`

## Propósito

Este documento fija el orden y el procedimiento para aplicar en producción las issues del saneamiento sin romper ninguna aplicación ni la cola de correo. Las propuestas, su coste y sus comprobaciones viven en el [catálogo de accionables](../fase-0/catalogo-accionables.md) y en las [comprobaciones del grupo (a)](../fase-0/comprobaciones-grupo-a.md); este plan no las repite.

Los cambios en el legacy los aplica el propietario (decisión 28).

## Etiqueta de las issues

Las 182 issues del saneamiento llevan la etiqueta `saneamiento:coexistencia` en los catorce repositorios: las 169 que enlaza el catálogo, las once issues de los fronts de la épica de `IDCorreo` ([#1740](https://github.com/DysTelefonica/access2web-blueprint/issues/1740)), la del backend ([#1741](https://github.com/DysTelefonica/access2web-blueprint/issues/1741)) y [Lanzadera#26](https://github.com/DysTelefonica/Lanzadera/issues/26). La etiqueta `estrategia:coexistencia` se mantiene para el resto de la migración.

## La cola de correo y Automatización

Hay tres tablas `TbCorreosEnviados` distintas. Solo la de `Correos_datos` entra en el saneamiento.

| Cola | Quién da de alta | Quién envía | En el saneamiento |
|---|---|---|---|
| `Correos_datos` (común) | Los once fronts de [#1740](https://github.com/DysTelefonica/access2web-blueprint/issues/1740) | `bats/EnviarCorreoNoEnviado.vbs` de Automatización, tarea `Automatizacion-Correo`, cada 5 minutos | Sí: decisión 22 |
| `Tareas_datos1` | Seis scripts vivos de Automatización (`BRASS.vbs`, `TareaExpedientes.vbs`, `GestionRiesgos_bat.vbs`, `Tareas.VBS`, `NoConformidades_bat.vbs`, `HPS_SOLICITUDES.vbs`) y el archivado `_legacy/CopiaDatosTE.vbs`, con `Max(IDCorreo) + 1` | `bats/EnviarCorreoTareas.vbs`, que marca por `IDCorreo` | No: Tareas está fuera del alcance (BFA-S4) |
| `Solicitudes_HPS_datos` | El front de Solicitudes HPS | `bats/vbs/EnviarCorreoHPS/EnviarCorreosHPS.vbs` | Pendiente: decisión 26 |

Ningún script de Automatización da de alta correos en `Correos_datos`. `EnviarCorreoNoEnviado.vbs` lee las filas con `FechaEnvio Is Null` y marca cada una con `.Update` sobre la fila abierta del recordset. Pide bloqueo pesimista, pero con cursor de cliente ADO lo trata como estático con bloqueo optimista: el `.Update` se convierte en un UPDATE que localiza la fila por sus columnas clave, y con `IDCorreo` como clave primaria la localizará por `IDCorreo`. Es inferencia del código; la confirma el ensayo. Es el consumidor que hay que proteger al cambiar la tabla.

Evidencia: repositorio `DysTelefonica/Automatizacion`, HEAD `9cb6fb5` (2026-11-23), `bats/EnviarCorreoNoEnviado.vbs:58`, `:86-109`, `:262-265`; `config/env-config.json:323-328`; `config/ScheduledTasks.ps1:34-79`.

## 1. Antes de tocar producción

1. **Ensayo de `IDCorreo` en una copia.** Lo hace Automatización en la parte 1 de [Automatizacion#53](https://github.com/DysTelefonica/Automatizacion/issues/53): ensayo de `EnviarCorreoNoEnviado.vbs` contra una copia de `Correos_datos` con la tabla recreada con `IDCorreo` autonumérico y clave primaria, y un front sin cambiar. Se comprueba:
   - que un front que todavía asigna `Max + 1` da de alta un correo con la columna ya autonumérica ([#1740](https://github.com/DysTelefonica/access2web-blueprint/issues/1740) lo afirma, pero no consta ensayado);
   - que `EnviarCorreoNoEnviado.vbs` envía y marca bien con la clave primaria nueva (no documentado);
   - si los fronts necesitan refrescar los vínculos tras el cambio;
   - el procedimiento del cambio: Access no convierte a autonumérico un campo con datos, de modo que se crea la tabla nueva con `IDCorreo` autonumérico y clave primaria, se anexan las filas conservando su identificador y se intercambian los nombres.

   **Método del ensayo, abierto.** El perfil `production-test` no puede ejecutar el emisor: lleva `AUTOMATIZACION_MAIL_MODE = QUEUE_ONLY` (`config/env-config.json:358`), `config/Environment.ps1:117-122` solo admite `SMTP` en `production`, y con `QUEUE_ONLY` `EnviarCorreoNoEnviado.vbs:306-308` termina sin abrir la cola. Cómo se ensaya lo decide el propietario con Automatización. Una vía que ya existe es la de `tests/vbs/EnviarCorreoNoEnviado.Tests.ps1`: una cola `.accdb` temporal con `IDCorreo` como clave primaria (`:66`) y un servidor SMTP falso (`:133-218`).
2. **Perfil de Automatización. Resuelto (propietario, 2026-11-25):** el servidor real usa `production` (`AUTOMATIZACION_PROFILE=production` en `C:\Automatizacion-Prod`, en CUEVA), que envía por SMTP contra los backends de producción (`config/env-config.json:307-351`). `production-test` apunta a `\\datoste\...\00_datos_prueba` (`:353-397`). El valor del `.env` del servidor lo afirma el propietario; el repositorio de Automatización no lo puede probar.
3. **Ventanas de mantenimiento.** Cambiar el diseño de un `.accdb` exige acceso exclusivo. Las tareas de `\Automatizacion\` (`config/ScheduledTasks.ps1:34-79`):
   - `Automatizacion-Correo`: cada 5 minutos, todos los días. Envía las tres colas.
   - `Automatizacion-Diaria`: cada 5 minutos, de lunes a viernes, de 07:00 a 20:00.
   - `Automatizacion-Resumen`: cada 5 minutos, todos los días. Abre `Tareas_datos1` y `Correos_datos` en solo lectura (`Exportar-ResumenBackends.ps1:121`, `:132-133`).
   - `Automatizacion-Doctor`: solo bajo demanda. `Automatizacion-Festivos` (los lunes a las 06:30) y `Automatizacion-Panel` (cada minuto, solo si se instala) no abren backends.

   No hay ninguna hora sin tareas que abran backends: todo cambio se hace fuera de horario y con las tareas pausadas.

## 2. Procedimiento de cada cambio

1. Pausar las tareas, guardando antes cuáles no estaban deshabilitadas, y comprobar que ninguna sigue en `Running` (`Disable-ScheduledTask` no para una instancia en curso):

   ```powershell
   $activas = Get-ScheduledTask -TaskPath '\Automatizacion\' | Where-Object State -ne 'Disabled'
   $activas | Disable-ScheduledTask
   ```
2. Comprobar que no hay nadie dentro del backend: no existe su `.laccdb`.
3. Si el cambio toca la cola, esperar a que no queden correos pendientes (`FechaEnvio Is Null`) o anotar cuáles quedan.
4. Copia de seguridad del backend, con su SHA-256.
5. Repetir la comprobación de datos de la propuesta contra producción tal como está ese día. Las del grupo (a) se hicieron sobre la [copia de referencia](../fase-0/copia-referencia.md) del 2026-11-05 y no autorizan aplicar.
6. Aplicar el cambio en el backend.
7. Comprobarlo: abrir el front y hacer la operación afectada; si toca la cola, dar de alta un correo de prueba.
8. Ejecutar el doctor de Automatización y comprobar que pasa. Después, reactivar solo las tareas guardadas en el paso 1, en la misma sesión de PowerShell, y comprobar que un ciclo de correo termina bien:

   ```powershell
   $activas | Enable-ScheduledTask
   ```

   El doctor (`config/DoctorResources.ps1:350-369`, con las tablas de `config/uat-bundle-manifest.json`) falla si falta o se ha renombrado una tabla que lista, y entonces termina toda pasada, también la de correo (lectura del código, no ejecutada). Las tareas siguen pausadas hasta que el cambio esté completo y el doctor pase.
9. Publicar el front adaptado: subir `VersionAplicacion` en el `.ini` del servidor para que la Lanzadera copie el `.accde` a cada usuario.
10. **Marcha atrás:** restaurar la copia de seguridad solo antes de reabrir a los usuarios. Después, restaurar borra lo grabado desde entonces: se corrige hacia delante.

## 3. Orden

### Bloque A. `IDCorreo` ([#1740](https://github.com/DysTelefonica/access2web-blueprint/issues/1740))

Va primero porque toca a las once aplicaciones. **Orden (cambiado el 2026-11-26): primero los once fronts en versión tolerante, después el backend.** En ACE, un alta con `IDCorreo` explícito mueve el contador autonumérico a ese valor + 1, así que un front con `Max + 1` conviviendo con la columna autonumérica puede bloquear a los demás ([evidencia de Automatización](https://github.com/DysTelefonica/Automatizacion/issues/53#issuecomment-7193543305)). Un front tolerante detecta si `IDCorreo` es autonumérico: si lo es, no da el número; si no, sigue con `Max + 1`. Se publica cuando le toque, sin ventana, y #1741 va cuando el último esté en producción.

1. Comprobación previa de la decisión 22: ninguna tabla guarda los valores repetidos: 22 en producción el 2026-11-26 (19 en la copia de referencia), y siguen apareciendo, así que se repite el día de la ventana de #1741. Se miran sobre todo las de AGEDYS Técnico, AGEDYS ECO y Gestión de Riesgos, que guardan `IDCorreo`.
2. Los once fronts, en versión tolerante, cada uno en su ciclo (en las aplicaciones con una RC1 pendiente de UAT, en una versión posterior). Los que guardan `IDCorreo` en otra tabla guardan, con la columna autonumérica, el que asigna la base.
   - Primero los que guardan `IDCorreo` en otra tabla: [AGEDYS#5](https://github.com/DysTelefonica/AGEDYS/issues/5), [AGEDYS_ECO#12](https://github.com/DysTelefonica/AGEDYS_ECO/issues/12) y [GESTION_RIESGOS#338](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/338).
   - Después: [Lanzadera#27](https://github.com/DysTelefonica/Lanzadera/issues/27), [Expedientes#279](https://github.com/DysTelefonica/Expedientes/issues/279), [condor#29](https://github.com/DysTelefonica/condor/issues/29), [seguridad#10](https://github.com/DysTelefonica/seguridad/issues/10), [HPS#52](https://github.com/DysTelefonica/HPS/issues/52), [HPS_SOLICITUDES#2](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/2), [No_conformidades#186](https://github.com/DysTelefonica/No_conformidades/issues/186) y [REGISTROS#2](https://github.com/DysTelefonica/REGISTROS/issues/2).
3. **Requisitos previos:** #1741 no se aplica en producción hasta que los once fronts tolerantes estén instalados y el ensayo de la parte 1 de [Automatizacion#53](https://github.com/DysTelefonica/Automatizacion/issues/53) esté documentado en esa issue.
4. [#1741](https://github.com/DysTelefonica/access2web-blueprint/issues/1741), en el backend, con el procedimiento del apartado 2: renumerar el segundo correo de cada pareja, `IDCorreo` autonumérico y clave primaria.
   - **Lo que necesita `EnviarCorreoNoEnviado.vbs`** (`9cb6fb5`). Accede a cada columna por su nombre, nunca por posición: lee `IDCorreo`, `Aplicacion`, `Destinatarios`, `DestinatariosConCopia`, `DestinatariosConCopiaOculta`, `Asunto`, `Cuerpo` y `URLAdjunto` (`:95-102`) y escribe `FechaEnvio` (`:108-109`). El orden de las columnas no importa. Abre `Correos_datos.accdb` directamente (`:262-265`), así que en Automatización no hay vínculos que refrescar; los de los fronts se comprueban en el ensayo. El procedimiento tiene que conservar nombres y tipos; pasar `IDCorreo` de entero largo a autonumérico es compatible.
   - **Lo que le falta a #1741:** definir con rigor «la misma estructura» (nombres, tipos, tamaños, «requerido» y «permitir longitud cero» de todas las columnas, con una comprobación escrita antes del intercambio; `FechaGrabacion` la lee `Exportar-ResumenBackends.ps1:132`); Automatización y la pausa de sus tareas, con la cola sin correos pendientes; el requisito previo de Automatizacion#53; y la marcha atrás de las filas grabadas en la tabla nueva después del intercambio, que se pierden si se vuelve a la tabla antigua.
   - Durante el intercambio de nombres falla el doctor: las tareas siguen pausadas hasta que la tabla nueva tenga su nombre y el doctor pase.

### Bloque B. Por aplicación

En el orden de la decisión 24: Lanzadera, Expedientes, Condor, Seguridad, HPS y Solicitudes HPS, Gestión de Riesgos y No Conformidades, Registros, Brass y APGEDYS. Dentro de cada aplicación:

1. Las propuestas del grupo (a) que la comprobación sobre producción dé por aplicables.
2. El front adaptado, publicado y observado.
3. Las del grupo (b) que apruebe el propietario.
4. Al final, la cuarentena (renombrar a `zzBAJA_`, decisión 23). El borrado, tras un cierre mensual sin incidencias y antes de migrar ese backend ([protocolo de huérfanas](../epica-coexistencia-postgresql.md#protocolo-de-huérfanas)).

Van en la misma tanda las aplicaciones que escriben en tablas de la otra ([orden de paso a la web](../fase-0/orden-cortes-web.md)): HPS con Solicitudes HPS (cuatro tablas de HPST), Gestión de Riesgos con No Conformidades (`TbRiesgosNC` y `TbNoConformidades`) y AGE-E1 con Registros (`TbSalidas_RES`).

### No se aplican hasta resolverlas

| Propuesta | Motivo |
|---|---|
| EXP-E2, GRI-R2, AGT-A1, BFA-S5, BR-B3 | Sin clave candidata o con duplicados ([comprobaciones del grupo (a)](../fase-0/comprobaciones-grupo-a.md)) |
| LAN-A5 | La sustituye [Lanzadera#25](https://github.com/DysTelefonica/Lanzadera/issues/25) |

### Propuestas que tocan tablas de Automatización

Antes de aplicar cualquiera de estas propuestas se avisa a Automatización. Qué lee o escribe cada script y la condición de cada una están en el [catálogo](../fase-0/catalogo-accionables.md#consumidores-en-automatización-1808).

- COR-C1 ([#1741](https://github.com/DysTelefonica/access2web-blueprint/issues/1741)): bloque A.
- NOC-N4 ([No_conformidades#189](https://github.com/DysTelefonica/No_conformidades/issues/189)) y GRI-R6 ([GESTION_RIESGOS#343](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/343)): si se recrean las tablas, mismos nombres de tabla y de columna.
- GRI-R5 ([GESTION_RIESGOS#341](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/341)): `TbExpedientes1` sigue existiendo en `Gestion_Riesgos_Datos`.
- EXP-E8 ([Expedientes#289](https://github.com/DysTelefonica/Expedientes/issues/289)): no incluye `OrdinalE2E`. Son dos ordinales globales, no identificadores de fila: `TbExpedientes.OrdinalE2E` se asigna al enlazar y `TbE2ECola.OrdinalE2E` es el número de envío; las dos tablas ya tienen su autonumérico ([respuesta de Expedientes](https://github.com/DysTelefonica/Expedientes/issues/289#issuecomment-7192427256)). En Access sigue siendo `Long` con índice UNIQUE y los generadores reintentan ante el error 3022 en vez de saltar la fila; en PostgreSQL se asigna con una secuencia.
- BFA-S3 ([#1754](https://github.com/DysTelefonica/access2web-blueprint/issues/1754)): el cambio de Automatización, fusionado antes.
- [Expedientes#294](https://github.com/DysTelefonica/Expedientes/issues/294): solo los objetos vínculo de `Expedientes_datos`, nunca las tablas físicas de `AGEDYS_DATOS`.

## 4. Decisiones pendientes del propietario

1. **Ramas RC1 con UAT pendiente.** En Expedientes, Gestión de Riesgos, Condor, No Conformidades y Brass, el cambio de `IDCorreo` sale sobre su RC1 o en una versión posterior. Condiciona la duración del bloque A. **Resuelta (propietario, 2026-11-25):** en las aplicaciones con una RC1 pendiente de UAT, el cambio de `IDCorreo` sale en una versión posterior.
2. **Cola local de Solicitudes HPS** (decisión 26): bloquea la fase 1 de Solicitudes HPS.
3. **Tablas de AGEDO** que usan Gestión de Riesgos y AGEDYS Técnico, y el botón `ComandoAGEDO` de la Lanzadera ([cierre de la fase 0](../fase-0/cierre.md)).
4. **Propuestas del grupo (b)** que se aprueban.
