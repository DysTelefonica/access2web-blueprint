# Issues propuestas del PRD de HPS Solicitudes

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Issue:** #1077 · **Fase:** 6 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Issues implementables que cierran el PRD de HPS Solicitudes, con la etiqueta `app/hps-solicitudes` y la numeración `P-HPS-S-N`. No se abre ninguna issue en este lote: las publica el propietario tras la aprobación del merge.

## Contexto y vínculos

- **Issue que se cierra con este PR:** #1077 (PRD completo cuando la punta se fusiona con las 16 PR que la preceden: inventario + 4 fichas + 2b + 3 + 4 + correos-a + correos-b + matriz-1 + matriz-2 + coexistencia + capacidades + esta punta).
- **PRD completo:** [`README.md`](README.md) con `Totales: 34/0/40` (34 formularios analizados, 0 no aplicables, 40 pendientes = 14 módulos + 26 clases). Faltan 9 linked (D-HPS-1, extracción Windows con `dysflow-erd`).
- **Hallazgos consolidados:** 1..58 cubren las 6 fichas mías; 59..N los completan apap2 y los pendientes de la matriz.
- **Coexistencia y matriz:** [`coexistencia.md`](coexistencia.md) y [`matriz-campos.md`](matriz-campos.md), [`matriz-campos-2.md`](matriz-campos-2.md) definen D-HPS-1 a D-HPS-12; cada issue propuesta hereda las decisiones abiertas del bloque que la origina.

## Issues propuestas (10)

### P-HPS-S-1 · Extraer el ERD de origen y resolver claves, índices, nulabilidad y FKs de las 16 tablas (11 locales + 5 linked con contenido en `src/`)

- **Origen:** D-HPS-1, recurrente en la coexistencia y en la matriz.
- **Descripción:** ejecutar el workflow `dysflow-erd` (#1144) sobre `Solicitudes_HPS_datos.accdb` y los `.accdb` linked (`Expedientes_datos.accdb`, `HPST.accdb`) desde la máquina Windows del propietario, en solo lectura. Cerrar las 11 tablas locales y las 5 linked con contenido en `src/` (16 en total).
- **Criterio de aceptación:** documento con la lista de claves primarias, índices, relaciones, nulabilidad y FKs físicas; al fusionar, las D-HPS-2 a D-HPS-12 (decisiones abiertas) pueden cerrarse con evidencia.
- **Tamaño:** ~150 líneas de `dysflow read-only`.

### P-HPS-S-2 · Implementar el modelo de datos web (`lanzadera.hps_solicitudes.*`) y los `INSERT`/`UPDATE` parametrizados

- **Origen:** D-HPS-2 (decisión de 20 columnas `SHORT_DATE_TIME` → 2 columnas o 20), D-HPS-10 (concurrencia optimista con columna `version`).
- **Descripción:** crear el esquema web para HPS Solicitudes: 11 tablas locales + 9 linked consolidadas; columna `version BIGINT DEFAULT 0` por tabla; FKs numéricas reales (`responsables.id` en vez de `email`); validación `IF @current_version = :expected_version THEN … ELSE raise conflict` en cada `UPDATE`. Migración inicial con corte explícito.
- **Criterio de aceptación:** esquema desplegado, migraciones Alembic (o equivalente) ejecutadas, FKs y `version` activas; las operaciones concurrentes sobre la misma fila se rechazan con `409 Conflict`.
- **Tamaño:** ~250 líneas entre modelo, migraciones y tests de concurrencia.

### P-HPS-S-3 · Sustituir la lectura de linked tables por endpoints federados

- **Origen:** D-HPS-8 (identidad compartida con Lanzadera), `Expedientes_datos.accdb`, `HPST.accdb` linked.
- **Descripción:** decidir entre (a) un único `users.id` con migración de las linked a un endpoint interno, o (b) dos identidades federadas vía `OpenID Connect`. Migrar `Expedientes_Datos.accdb` y `HPST.accdb` a tablas de la web consolidada con corte controlado; eliminar la dependencia de `DoCmd.OpenRecordset` contra `.accdb` linked.
- **Criterio de aceptación:** cero referencias a linked tables en `src/`; los formularios que consultaban linked ahora leen de un endpoint (`GET /expedientes?…`, `GET /usuarios-hps?…`); los tests de aceptación bidireccional pasan.
- **Tamaño:** ~300 líneas entre `auth`, `hps_client`, endpoints y migración.

### P-HPS-S-4 · Transacciones atómicas para `MotivoHPS.EditarMotivo` y `SolicitudOperaciones.Registro*`

- **Origen:** D-HPS-9 (`MotivoHPS.EditarMotivo` en 4 tablas), D-HPS-11 (operaciones multi-tabla).
- **Descripción:** envolver `MotivoHPS.EditarMotivo` y cada `SolicitudOperaciones.Registro*` (los 7 hitos: `RegistroEnvioExcel`, `RegistroEnvioRecordatorioExcel1/2`, `RegistroEnvioAutocancelacionPREMARGA`, `RegistroEnvioRecordatorioMarga1/2`, `RegistroEnvioAutocancelacionMARGA`) en `BEGIN/COMMIT/ROLLBACK` (Access) o `@transactional`/`savepoints` (web). La web exige `transaction.atomic` en el endpoint POST correspondiente.
- **Criterio de aceptación:** un fallo de cualquier `UPDATE` intermedio hace `ROLLBACK`; no quedan filas en `TbMotivoHPS`, `TbSolicitudes`, `TbUsuarios` o `TbUsuariosHistoricos` con estado inconsistente.
- **Tamaño:** ~100 líneas de wrapper transaccional + tests de rollback.

### P-HPS-S-5 · SMTP con TLS explícito y seudonimización de credenciales

- **Origen:** D-HPS-12 (`CDO.Message` sin TLS explícito), credenciales en plano.
- **Descripción:** migrar el envío de correo a un servicio SMTP autenticado con `STARTTLS` o `SMTP over TLS`; las credenciales se inyectan desde el gestor de secretos del cluster (Vault/SOPS); se elimina cualquier literal en `Variables Globales.bas` o `constructor.bas`. La cola de envío (`TbCorreosEnviados`) se sustituye por un endpoint `POST /correos` server-side; la firma HTML de la plantilla se sella con hash (D-HPS-13).
- **Criterio de aceptación:** cero secretos en claro en el código; el endpoint de envío exige `SMTP_PASSWORD` server-side; los tests de envío registran el hash de la plantilla.
- **Tamaño:** ~150 líneas entre servicio SMTP, inyección de secretos y tests de envío.

### P-HPS-S-6 · Anonimización y seudonimización de datos personales

- **Origen:** D-HPS-3 (`DNI`, `email`, `emailResponsable`, `Gestor`, `Usuario` en `TbLogsGeneral`), H-HPS-S-26 (generador de DNI/NIE sintético), H-HPS-S-49 (datos personales en `RowSource` de `Form_FormTareasTramitadorPendientes`).
- **Descripción:** sustituir `DNI` por `dni_hash` (HMAC-SHA-256 con sal por usuario) en `TbLogsGeneral.DNI` y `TbSolicitudes.DNI`; sustituir `Gestor`, `emailResponsable` y `Usuario` por FKs a `users.id`; eliminar el generador de DNI/NIE sintético en producción (mantenerlo solo bajo `EnDesarrollo = "Sí"`); purgar los `RowSource` de los `.form.txt` que cargan datos personales.
- **Criterio de aceptación:** `dni_hash` reversible solo para admin auditado; las pantallas de la web muestran `users.nombre_completo` en lugar de `Gestor`/`emailResponsable`; los tests de regresión no muestran nombres reales en fixtures.
- **Tamaño:** ~120 líneas entre hashing helper, FKs, fixtures y tests.

### P-HPS-S-7 · Eliminar los 4 formularios sin cablear y la rampa de correos desactivada

- **Origen:** H-HPS-S-23 (`Form_FormAltaDatosSolicitante`), H-HPS-S-29 (`Form_FormMotivosHPS`), H-HPS-S-26 (`ComandoGenerarDNIValido`), H-HPS-S-38 (`Form_FormCambioTipo`; H-HPS-S-38 lo lista), la rampa de correos del menú principal (H-HPS-S-2).
- **Descripción:** confirmar con el propietario que estos cuatro formularios son residuo de un refactor previo; si lo son, eliminar su `.cls`, `.form.txt` y referencias en el menú principal. Eliminar la rampa comentada en `Form_Form0BDOpciones.cls:230-256` (`ActivadoCorreoAutomatico = "No"` en producción) y traducir el flag a un endpoint `GET /configuracion` server-side.
- **Criterio de aceptación:** los 4 formularios no aparecen en el menú, no se compilan y no se cargan; la rampa de correos queda como un único endpoint configurable; los tests E2E confirman que las opciones de menú visibles son las aprobadas.
- **Tamaño:** ~80 líneas entre borrado de `.cls`/`.form.txt`, simplificación de menú y tests.

### P-HPS-S-8 · Refactor de los `Err.Numbr` y la `MsgBox` de error copy-paste

- **Origen:** H-HPS-S-25, S-32, S-37, S-48 (literal copy-paste en `ComandoAyuda_Click`, `ComandoRegistrar_Click`, `ComandoUrlPlantillas*_Click`, `ComandoEliminar_Click`).
- **Descripción:** sustituir el literal de error copy-paste por un helper único (`RaiseHpsError` o equivalente) que centralice la construcción del mensaje, la notificación a `CorreoAlAdministrador` y el `MsgBox`; corregir el typo `Err.Numbr` en las cinco formas afectadas. Reemplazar `MsgBox` críticos por toasts accesibles desde la consola de la web.
- **Criterio de aceptación:** cero apariciones de `Err.Numbr` o de literales de error copy-paste en `src/`; los tests de error verifican el helper; los toasts reemplazan los `MsgBox` en los flujos críticos.
- **Tamaño:** ~60 líneas entre el helper, las llamadas y los tests.

### P-HPS-S-9 · Sustituir los `RowSource` literales con endpoints

- **Origen:** H-HPS-S-7, S-8, S-49 (`RowSource` de `Form_FormEntidades`, `Form_FormTareasTramitadorPendientes` con nombres y empresas reales).
- **Descripción:** eliminar los `RowSource` literales con datos reales de los `.form.txt`; en la web, los combos se cargan desde `GET /solicitudes?…`, `GET /responsables?…` con `Value List` o `Recordset`-equivalente.
- **Criterio de aceptación:** cero `RowSource` con literales de datos personales; los combos se hidratan desde la API; los tests de regresión confirman los datos esperados.
- **Tamaño:** ~70 líneas entre carga de combos y tests.

### P-HPS-S-10 · Eliminar `Copia de TbExpedientes` y tablas linked huérfanas

- **Origen:** D-HPS-5 (`Copia de TbExpedientes`), D-HPS-6 (`TbLogs` vacía), D-HPS-7 (`TbUltimoCambio` con FK a linked).
- **Descripción:** descartar la local `Copia de TbExpedientes` (sentinel) si ningún consumidor la lee; eliminar `TbExpedientesLugaresEjecucion`, `TbHPSGrado1` y `TbHPSEquivalencia` si están huérfanas; consolidar `TbLogs` (vacía) en `TbLogsGeneral` o eliminarla; traducir `TbUltimoCambio.IDUsuarioCambio` a FK a `users.id` (no a `TbUsuarios` linked).
- **Criterio de aceptación:** cero referencias a las tablas eliminadas en `src/`; las migraciones reflejan la decisión; los tests de integridad pasan.
- **Tamaño:** ~50 líneas entre migración y verificación de integridad.

## Pendientes

- Decidir el orden de merge: esta punta va después de `capacidades` (#1305 cadena: matriz-1, #1308 cadena: matriz-2, coexistencia, capacidades, esta punta).
- Confirmar que el mantenedor quiere abrir las 10 issues tras el merge final.
- D-HPS-1: todavía pendiente de extracción Windows con `dysflow-erd`; las 10 issues dependen en distinto grado de esa extracción.
- Política de retención de `TbLogsGeneral` (2058 filas) y `TbSolicitudesFechas` (245 filas × 20 columnas): decidir partición por `Fecha` antes de migrar.
- Anexo del binario: el `Solicitudes_HPS.accdb` y el `Solicitudes_HPS_datos.accdb` no se migran con `dysflow` directamente; requieren decisión de corte (HPS Solicitudes es una de las 8 aplicaciones y el orden del roadmap global lo decide el propietario, no este PRD).
- Decisión final sobre `C-HPS-S-10` (tareas pendientes) y la consolidación con la cola de Lanzadera.
- Aprobación de las decisiones D-HPS-2 (20 fechas vs 2), D-HPS-3 (seudonimización), D-HPS-8 (identidad consolidada) y D-HPS-12 (SMTP/TLS) por el propietario antes de implementar P-HPS-S-2, P-HPS-S-5, P-HPS-S-6.
