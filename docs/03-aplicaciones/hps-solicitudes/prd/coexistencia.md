# Coexistencia Access y web — HPS Solicitudes

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Issue:** #1077 · **Fase:** 4 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Qué debe adaptarse, preservarse o decidirse para que HPS Solicitudes opere sobre tablas PostgreSQL mientras la web la sustituye. Todo es lectura estática de `src` (excluidos los módulos `Test*`): no se ha ejecutado nada ni comprobado compatibilidad real con PostgreSQL. Las propuestas no son requisitos hasta que se aprueben. La arquitectura de datos de la coexistencia (P-EXP-3) está abierta y en estudio en el PRD de Expedientes (`docs/03-aplicaciones/expedientes/prd/coexistencia.md`, apartados 11 y 12, en la cadena de #1073); este documento mide lo mismo para HPS Solicitudes sin recomendar ninguna vía.

## 1. Cómo accede hoy HPS Solicitudes a los datos

| Hecho | Evidencia |
|---|---|
| El acceso a datos pasa por `constructor.getdb()` (heredado de `modules/constructor.bas`, replicado de Lanzadera): muchas llamadas en los repositorios y operaciones. | `modules/constructor.bas:getdb`, `getSolicitud`, `getExpediente`, `getExpedientesBusqueda`, `getConfiguracion`, `getUsuario`; `modules/CorreoRepositorio.bas:1-50`; `modules/JustificacionRepositorio.bas:1-50`; `modules/AutomatizacionRepositorio.bas:1-50` |
| La ruta y la clave del backend salen de la variable global `m_ObjEntorno.URLBackendHPS` (o análoga en `Variables Globales.bas`) leída por `constructor.getdb()` y por `Form_FormConfiguracion`. | `Form_FormConfiguracion.cls:325-369`; `Variables Globales.bas` (no leído directamente en este PRD) |
| No hay cadenas ODBC ni `.Connect` ni consultas paso-a-través en HPS Solicitudes. | 0 apariciones en `src` (búsqueda exhaustiva con `grep -E "ODBC|\.Connect|pt (pass)" src/`) |
| La linked-table `TbExpedientes` (de `Expedientes_datos.accdb`) y `TbHPS`/`TbUsuariosHPS` (de `HPST.accdb`) acceden a backends ajenos sin ODBC explícito: Access los vincula automáticamente. | `Form_FormExpedientesBusqueda.cls:233-238`; `Form_FormSolicitudRenovacionDatos.cls:640-643` |

## 2. Reglas del propietario, verificadas en HPS Solicitudes

- **Una sola tabla local propia:** el snapshot `ERD/schema.sql` lista 11 tablas locales (Copia de TbExpedientes, TbConfiguracion, TbCorreosEnviados, TbHPSGrado, TbJustificaciones, TbLogs, TbLogsGeneral, TbResponsables, TbSolicitudes, TbSolicitudesFechas, TbUltimoCambio). De estas, `TbConfiguracion` se queda en Access (decisión del propietario en HPS Solicitudes) por su rol de singleton de configuración global (`Rows: 1`); las otras 10 migran.
- **Tablas temporales o de configuración por puesto** (`TbUsuarioConfiguracion` y similares): 0 apariciones en `src/` de HPS Solicitudes. La regla del propietario (viven en el backend, no en el frontend) no se aplica directamente aquí, aunque `TbResponsables` y `TbConfiguracion` juegan un rol análogo de «datos de la aplicación» que se conserva en backend.
- **Formularios desconectados: confirmado.** 0 `RecordSource` y 0 `ControlSource` en los 34 `.form.txt`; los listados y combos con tipo de origen declarado son `Value List` o se cargan en `Form_Open` con `getdb().OpenRecordset` (Form_FormExpedientesBusqueda, Form_FormTareasTramitadorPendientes, Form_FormCorreoPruebas).
- **`constructor.getdb()`** se apoya en una variable global del módulo `Variables Globales.bas` para resolver la ruta del backend, leída con `DBEngine.Workspaces(0).OpenDatabase(ruta, False, False, "MS Access;PWD=…")`; la contraseña y el path se protegen en `EVE` (`modules/Variables Globales.bas:103-108`).

## 3. Lo específico de HPS Solicitudes: identidad y linked tables

Además del acceso a datos, HPS Solicitudes abre **ficheros `.accdb` linked**:

- **`Expedientes_datos.accdb`** (vía `TbExpedientes`, `TbExpedientesLugaresEjecucion`): 9 referencias en `src/` a la primera, 0 a la segunda. La local `Copia de TbExpedientes` es sentinel de copia previa a un cambio masivo; ver D-HPS-5.
- **`HPST.accdb`** (vía `TbHPS`, `TbHPSEquivalencia`, `TbHPSGrado1`, `TbMotivoHPS`, `TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`): 0 referencias explícitas en `src/`. La integración con la web de HPS se hace a través del catálogo de motivos (`Form_FormMotivoHPS`) y del identificador `IDUsuarioHPS` en `TbSolicitudes`; la canónica de la identidad de HPS vive en la linked `TbUsuarios`.

A diferencia de Lanzadera (que gestiona el despliegue de otras aplicaciones), HPS Solicitudes no altera binarios de otros `.accdb`; su rol de integración es de identidad y catálogo, no de despliegue.

## 4. Qué se adapta y qué se preserva

| Construcción | Ocurrencias | Primer ejemplo | Riesgo con PostgreSQL |
|---|---|---|---|
| `getdb().OpenRecordset` y `OpenRecordset` directo | 50+ líneas en 8 ficheros | `Form_FormExpedientesBusqueda.cls:233-238`; `modules/constructor.bas:2306-2310` | Recordsets actualizables exigen clave única declarada en ODBC; el `schema.sql` no la documenta. |
| `AddNew`/`Edit`/`Update` | 20+ líneas en 7 ficheros | `Form_FormCorreo.cls:103-119`; `MotivoHPS.cls:41-44, 180-182, 208-213` | Migrar a `INSERT`/`UPDATE` parametrizados; la última rama de `MotivoHPS.EditarMotivo` hace `UPDATE` en 4 tablas sin transacción (D-HPS-9) |
| `Delete *` o `.Delete` | 5 líneas en 2 ficheros | `Form_FormCorreosGestion.cls:344-350` | Sustituir por `DELETE FROM … WHERE …` |
| `OpenRecordset` con `dbOpenSnapshot` | (no en `src/`) | — | Los snapshots Access son read-only; PostgreSQL `READ COMMITTED` requiere configurar la transacción o usar un SELECT explícito |
| `RecordsAffected` y `AbsolutePosition` | (no en `src/`) | — | El modelo web no tiene equivalente directo; la API usa `result.rowCount` o `result.fetchone()` |
| `Me.RecordSource =` y `Me.ControlSource =` | 0 (formularios todos desconectados) | — | Cero riesgo en UI; todo se carga explícitamente en `Form_Open` |
| `Set m_DB = CurrentDb` o `CurrentDb.OpenRecordset` | 0 en HPS Solicitudes (Lanzadera sí lo hace para `TbConfiguracionBackends`) | — | HPS Solicitudes no escribe en una tabla local propia; toda escritura pasa por `getdb()` |
| `Yes`/`No` en `TEXT(2)` | 7+ columnas (4 en `TbConfiguracion`, 2 en `TbSolicitudes`, 1 en `TbLogsGeneral`) | `Form_FormConfiguracion.cls:115-130`; `Form_FormAltaDatosSolicitante.cls:218-220` | Conservar como `BOOLEAN` con CHECK `IN ('Sí','No')` o pasar a dominio web nativo (`BOOLEAN`) |
| `Me.OpenArgs` como `String` libre | 2 (`"MODO_GESTION"`, `CLng(tipoEnvio)`) | `Form_FormCorreo.cls:60-65`; `Form_FormEnvioCorreo.cls:363-381` | Validar contra tipos enumerados en la web; el `OpenArgs` se sustituye por parámetro en la URL |
| `dbOpenDynamic` implícito | (todos los recordsets) | `Form_FormTareasTramitadorPendientes.cls:312-553` | En web, cursor implícito; la API devuelve el conjunto directamente sin cursor |
| `Me.Recordset.FindFirst` / `.Seek` / `.Bookmark` | 0 | — | Cero riesgo |
| `dbPessimistic`/`dbOptimistic` | 0 | — | Concurrencia optimista en la web: `If-Match` + `ExpectedVersion` |
| `Like '*texto*'` | 2 (`Form_FormExpedientesBusqueda.cls`, `constructor.bas:3888-3905`) | `constructor.bas:3894` | `ILIKE '%texto%'` con índice trigram; el `*` Access es `*` o `?` en el shell, hay que escaparlo al traducir |
| `SELECT TOP n` | 1 (`constructor.bas:3873` carga `TbExpedientes` entera) | `Form_FormExpedientesBusqueda.cls:236-250` | `LIMIT n`; pero el código actual carga la tabla entera, no se beneficia del TOP — candidato a reescritura |
| `Format(…, "mm/dd/yyyy")` y `#…#` | 0 apariciones | — | Sin riesgo directo |
| `Date()` en SQL | 0 | — | Sin riesgo |
| `DMax`, `DLookup` | 0 | — | Sin riesgo |
| `ID` calculado con `Max+1` | 0 (constructor centraliza) | `modules/constructor.bas:getSolicitud`, `getExpediente` (revisión pendiente) | PostgreSQL `SERIAL` o `BIGSERIAL`; las referencias `DMax(…)` se sustituyen por `RETURNING id` |
| `Me.Recordset.AddNew` | 0 (constructor) | — | Mismo punto: sustituir por el helper que devuelve el `id` |
| `dbFailOnError` | (no explícito) | `constructor.bas:open` (a verificar) | `RaiseError` en la web; los `Resume Next` se traducen a logs de error |
| Concurrencia optimista (`LastModified`, version) | 0 | — | Sin versión de fila: dos escrituras simultáneas se pisan sin aviso (D-HPS-10) |
| Transacciones (`BeginTrans`/`CommitTrans`/`Rollback`) | 0 | — | Múltiples `UPDATE` (MotivoHPS, SolicitudOperaciones) sin transacción atómica; la web requiere `BEGIN/COMMIT/ROLLBACK` por operación multi-tabla (D-HPS-11) |
| Plantillas HTML con `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` | 4 referencias | `Form_FormCorreo.cls:296-302`; `Form_FormEnvioCorreo.cls:177-180`; `Form_FormCorreoPruebas.cls:170-179` | Web: URL firmada por tiempo limitado, sin acceso a UNC |
| `CDO.Message` para envío de correo | 1 (en `CorreoServicio.Enviar` o equivalente) | (no en `src/` de HPS Solicitudes; vive en `classes/CorreoServicio.cls`/módulo central) | SMTP autenticado con `SMTP_PASSWORD` server-side; sin TLS plano (D-HPS-12) |

Se preserva el comportamiento observable: los `Sí`/`No` se conservan como `BOOLEAN` con dominio cerrado, los `Long(4)` pasan a `INTEGER` con `SERIAL`/secuencia, los `MEMO` a `TEXT`, los `TEXT(n)` a `VARCHAR(n)`, los `SHORT_DATE_TIME` a `TIMESTAMP` con zona horaria.

## 5. Recordsets, identidad, transacciones y tipos

- **Recordsets:** `OpenRecordset` aparece 50+ veces en 8 ficheros. Sin `dbOpenSnapshot`, `dbOpenTable`, `.Seek`, `.Index`, `.FindFirst`, `.Bookmark`, `.Filter` ni `.Sort`. La coexistencia asume `ReadCommitted` por defecto en PostgreSQL; los `Recordset` se sustituyen por el resultado de la consulta (`rows.fetchall()`), sin cursor persistente.
- **Identidad:** `Me.RecordsetClone`, `.AddNew` directo o `DMax(…)` no se usan; la identidad se gestiona por la capa `constructor` (revisión pendiente en `modules/constructor.bas`). En la web: `INTEGER IDENTITY` o `UUID` con `RETURNING id`; el modelo de clave compuesta para tablas `N-M` (linked `TbUsuariosEntidades`, `TbHPSEquivalencia`) requiere PK explícita.
- **Transacciones:** 0 `BeginTrans`, `CommitTrans` o `Rollback`. Las escrituras multi-tabla (`MotivoHPS.EditarMotivo` en 4 tablas, `SolicitudOperaciones.Registro*` que toca `TbSolicitudesFechas` y `TbSolicitudes`) no son atómicas hoy; caída intermedia deja inconsistencia detectable (D-HPS-11).
- **Tipos:** 0 columnas `BOOLEAN` declaradas en `ERD/schema.sql` (las 4 de `TbConfiguracion` están como `TEXT(2)` con valor `Sí`/`No`). Fechas: 100+ columnas `SHORT_DATE_TIME` requieren decisión zona horaria. `LONG(4)` (autoincremental con `Max+1`) requiere transición a `SERIAL`; las 4 linked no requieren clave propia.

## 6. Identidad, permisos y credenciales

- **Identidad compartida con Lanzadera:** las linked `TbUsuarios`, `TbUsuariosEntidades` y `TbUsuariosHistoricos` viven en `HPST.accdb`; Lanzadera gestiona la identidad de las otras aplicaciones en `Lanzadera_Datos.accdb`. La web consolidada debe decidir entre (a) un solo `users.id` y migrar las linked a un endpoint de identidad único, o (b) dos identidades federadas vía `IdP` y `OpenID Connect`. D-HPS-8.
- **Credenciales de correo:** la contraseña del SMTP para `CDO.Message` vive como literal en `Variables Globales.bas` (a verificar) o como variable de entorno; la web la inyecta desde el gestor de secretos del cluster.
- **Sin TLS plano al SMTP:** `CDO.Message` se invoca sin `ConfigurationFields` para forzar `smtps://` o `STARTTLS`; a verificar si la configuración del servidor es TLS o no. D-HPS-12.
- **Auditoría con datos personales:** `TbLogsGeneral.DNI` y `TbLogsGeneral.Usuario` se escriben en plano. La web debe seudonimizar el `DNI` (hash con sal) y reemplazar el `Usuario` por FK a `users.id`.

## 7. Concurrencia y auditoría

Sin bloqueo explícito (0 usos de `dbPessimistic` u `dbOptimistic`) y sin versión de fila. La coexistencia asume `READ COMMITTED` por defecto en PostgreSQL; los `UPDATE` sin `WHERE` por clave primaria compuesta o por `Version` se pisan sin aviso. La web requiere una columna `version BIGINT DEFAULT 0` por tabla, incrementada en cada `UPDATE`, y validación `IF @current_version = :expected_version THEN … ELSE raise conflict` (D-HPS-10). El log de auditoría vive en `TbLogs` (vacía) y `TbLogsGeneral` (2058 filas); la web lo cubre con `TbLogsGeneral` y partición por `Fecha` (mensual).

## 8. Pruebas de aceptación bidireccionales (propuestas)

A ejecutar después, en entorno aislado y nunca sobre producción:

1. Un usuario creado en la web inicia sesión en HPS Solicitudes de Access, y al revés, con el mismo resultado de permisos.
2. Una solicitud creada en Access aparece en la web, y viceversa, con la misma `IDSolicitud` y los mismos hitos.
3. Un `UPDATE` simultáneo desde dos ventanas sobre la misma `IDSolicitud` se rechaza con `409 Conflict`; la última escritura gana solo si su `version` coincide.
4. Un cambio de `MotivoHPS` desde una ventana actualiza las 4 tablas implicadas (`TbMotivoHPS`, `TbSolicitudes`, `TbUsuarios`, `TbUsuariosHistoricos`) en una sola transacción; cualquier fallo intermedio hace `ROLLBACK`.
5. Un correo programado desde la web entra en la cola de HPS Solicitudes y se envía con la plantilla `VersionPlantillasHTML` activa.
6. El filtro por `PalabraClave` en `Form_FormExpedientesBusqueda` funciona con mayúsculas/minúsculas indistintas (`ILIKE '%…%'`) y devuelve menos de 100 ms para 1 000 expedientes.
7. Una solicitud con `Tipo = "Alta"` se renueva: se actualiza `FechaCreacion` y se generan los hitos de recordatorio según `TbConfiguracion.DiasParaRecordatorioExcel1` y `DiasParaRecordatorioExcel2`.
8. Una solicitud con `Tipo = "Traspaso"` se cierra: `URLAdjuntoEnvioONS` se rellena y `FechaEnvioTraspasoONS` se pinta.

## 9. Corte de la aplicación

HPS Solicitudes es una de las pantallas de Lanzadera. Al activar la web de HPS Solicitudes se deshabilita **su** botón en Lanzadera; el corte de Lanzadera depende de que las otras aplicaciones (Expedientes, HPS) hayan migrado. Deshabilitar el botón no borra binarios ni datos, ni cierra sesiones abiertas o accesos directos. La reversión está por definir.

## Pendientes

- D-HPS-1: extracción de solo lectura del backend `Solicitudes_HPS_datos.accdb` y de los `.accdb` linked desde Windows con `dysflow-erd` (#1144) para confirmar claves primarias, índices, nulabilidad y FKs de las 16 tablas (11 locales + 5 linked con contenido en `src/`).
- D-HPS-8: consolidar la identidad de las linked `TbUsuarios`/`TbUsuariosEntidades`/`TbUsuariosHistoricos` con la de Lanzadera (decisión abierta del bloque P-EXP-3 de la coexistencia global).
- D-HPS-9: `MotivoHPS.EditarMotivo` ejecuta `UPDATE` en 4 tablas sin transacción; D-HPS-11 propone `BEGIN/COMMIT/ROLLBACK` por operación multi-tabla.
- D-HPS-10: concurrencia optimista por columna `version` en cada tabla migrada.
- D-HPS-12: SMTP sin TLS explícito: pasar a `smtps://` o `STARTTLS` server-side.
- D-HPS-3: seudonimización de `DNI` y `Usuario` en `TbLogsGeneral` antes del corte.
- D-HPS-5: descartar la local `Copia de TbExpedientes` (sentinel) si ningún consumidor la lee; las linked `TbExpedientesLugaresEjecucion`, `TbHPSGrado1` y `TbHPSEquivalencia` también son candidatas a descarte.
- Política de retención de `TbLogsGeneral` (2058 filas en el snapshot): partición por `Fecha` o archivo en frío.
- Verificación de `dbOpenSnapshot` y `dbFailOnError` en `constructor.open` (no leídos en este PRD).
- Sentido de `m_ObjEntorno.URLBackendHPS` en el contexto actual: ruta de red, UNC o local; migrar a URL de servicio web.
- Cobertura de `dbOpenSnapshot`, `dbOpenDynamic`, `dbPessimistic`, `dbOptimistic`: lectura estática no exhaustiva, a complementar con `dysflow`.
