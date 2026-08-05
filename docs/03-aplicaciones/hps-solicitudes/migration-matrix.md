# HPS_Solicitudes — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. Disposiciones específicas de HPS_Solicitudes en § D98-D99.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbSolicitudes` (245 filas) | solicitud HPS | separar identidad, fechas, adjuntos, responsables, ONS | mapeado preliminar | unicidad por `IDSolicitud`; **FK por email ⚠️** | volumen real |
| `TbResponsables` (27 filas) | responsables de solicitudes | entidad de catálogo con email | mapeado preliminar | unicidad; **FK conceptual por email** | ⚠️ D99 migrar a FK numérica |
| `TbJustificaciones` (7 filas) | justificaciones de solicitud | entidad de justificación | mapeado preliminar | unicidad; ⚠️ **dirección de FK confusa** | ⚠️ D99 normalizar FK |
| `TbLogs` (0 filas) | log específico (en desuso) | descartar o archivar | mapeado preliminar | `TbLogsGeneral` la reemplazó | descartar |
| `TbLogsGeneral` (2058 filas) | log general activo | evento con timestamp + actor + contexto | mapeado preliminar | trazabilidad; **traducir a logs estructurados canónicos (D27)** | volumen real en producción |
| `TbCorreosEnviados` (9 filas) | registro de correos | evento de correo | mapeado preliminar | trazabilidad; volumen bajo | traducir a logs canónicos |
| `TbHPSGrado` (13 filas) | catálogo de grados HPS | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbSolicitudesFechas` (245 filas) | fechas de solicitud (1:1 con TbSolicitudes) | FK a `TbSolicitudes` + campos de fecha | mapeado preliminar | cardinalidad 1:1 con TbSolicitudes | cardinalidad |
| `TbConfiguracion` (1 fila presumido) | configuración de la app | config key-value | mapeado preliminar | unicidad; **migrar a config del módulo** | formato de almacenamiento |
| `TbUltimoCambio` (volumen presumido 245) | último cambio en solicitud | tracking de cambio | mapeado preliminar | cardinalidad 1:1 con TbSolicitudes | cardinalidad |
| **Copia de TbExpedientes** (legacy copy) | copia legacy de TbExpedientes | conservar en zona `legacy` | legacy copy | comparar recuentos y uso en código | archivo vs dominio |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera`) | identidad / permisos / entorno | identidad vía adaptador unificado (D9-D10) | mapeado preliminar | coherencia | acoplamiento directo a romper (D86/D87) |
| integraciones externas (Lanzadera, HPS, ONS) | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera de solicitudes, responsables, justificaciones, fechas, configuración, logs.
- **Necesita decisión:** FKs por email (D99), FKs conceptuales cross-app, política de datos personales (D98).
- **Legacy copy:** `Copia de TbExpedientes`.
- **External reference:** adjuntos (URLAdjunto), adjuntos ONS (URLAdjuntoEnvioONS), identidad, Lanzadera, HPS.
- **No hay campos/filas declarados obsoletos** (salvo los marcados en D98/D99).

## D98 · Datos personales en `TbSolicitudes` (245 filas)

**Estado**: PROPUESTO. Riesgo de seguridad análogo a HPS (D92).

### Hallazgo

`TbSolicitudes` contiene **245 filas con datos personales completos**: `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono`. Mismo patrón que HPS (ver D92).

Recomendaciones:

1. **Preservar en PostgreSQL sin transformaciones** (mantener paridad funcional).
2. **Enmascarar en logs y observabilidad** (no loguear valores completos).
3. **Documentar en matriz de capabilities** quién puede ver cada campo (D45).
4. **Evaluar encriptación en reposo** como follow-up de seguridad.
5. **Auditar el `.gitignore` del repo** para asegurar que `Solicitudes_HPS.accdb` no está siendo commiteado (regla D92 cross-cutting).

## D99 · FKs conceptuales con data integrity gaps

**Estado**: PROPUESTO.

### Hallazgos

1. **`TbResponsables.Correo → TbSolicitudes.emailResponsable`**: FK por **texto email**, no por ID. Riesgo de inconsistencia si el email cambia.

2. **`TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`**: dirección de FK confusa. La FK va de `TbJustificaciones` a `TbSolicitudes`, no al revés.

### Recomendaciones

1. **FK numérica para responsables**: agregar columna `idResponsable` (BIGINT) en `TbSolicitudes` y poblar con la conversión actual `emailResponsable → id`. Mantener `emailResponsable` como campo independiente (denormalización para búsqueda).

2. **Normalizar FK de justificaciones**: formalizar la FK con la dirección `TbSolicitudes.idjustificacion → TbJustificaciones.id` (numérica).

3. **FKs conceptuales cross-app** (`idExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`): mantener como referencia conceptual (D86/D87). Documentar la convención.

## D100 · `TbLogs` vacía — presumible desuso

**Estado**: PROPUESTO.

`TbLogs` tiene **0 filas** mientras que `TbLogsGeneral` tiene **2058 filas**. Esto sugiere que `TbLogs` está en desuso y fue reemplazada por `TbLogsGeneral`. Recomendación: **migrar `TbLogsGeneral` como tabla de logs principal** y archivar/migrar `TbLogs` solo si se decide reactivar.

## D101 · Integración con ONS (Organismo Notificador de Seguridad)

**Estado**: PROPUESTO.

HPS_Solicitudes tiene un sistema de **traspasos a ONS** externo:
- `Form_FormAdjuntaTraspasoONS.cls` para adjuntar traspasos.
- `URLAdjuntoEnvioONS` (Memo en `TbSolicitudes`) — ruta a fichero para ONS.
- Múltiples forms de traspaso: `Form_FormSolicitudAltaTraspaso.cls`, `Form_FormSolicitudesAltaTraspasoDatos.cls`, `Form_FormSolicitudesTraspasoFechas.cls`.

Recomendaciones:

1. **Migrar el sistema de traspasos a ONS como integración con servicio externo** vía adaptador (D16).
2. **Los adjuntos (`URLAdjuntoEnvioONS`) se migran al object storage** detrás del puerto (D16).
3. **Disponer de un endpoint de "traspaso a ONS"** que serialice la solicitud + adjuntos y los envíe al servicio externo ONS vía API o cola.
4. **Auditar el estado actual de la integración con ONS** — el campo `ExpedienteUnificado = "No"` indica que la unificación está desactivada.