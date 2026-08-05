# NoConformidades — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\NoConformidades_Datos.accdb` (27 MB). **Inventario real obtenido vía Dysflow MCP el 2026-08-05** después de resolver la ambigüedad multi-worktree (`register_worktree` + `migrate_project_config` + `accessPath` absoluto explícito). La diagnosis original D89 ("fallo de `list_objects`") estaba basada en análisis estático sin contactar el runtime y queda **invalidada** (ver [Seguridad § D89 invalidado](security-rules.md#d89--diagnóstico-del-fallo-de-list_objects-de-dysflow--invalidado)).

## Inventario real Dysflow (2026-08-05)

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **42** en `NoConformidades_Datos.accdb` |
| **NCs de Proyecto** | **438** filas en `TbNoConformidades` |
| **NCs de Auditoría** | **55** filas en `TbNoConformidadesAuditoria` |
| **FK relationships** (user tables) | **14** |
| **Columnas en `TbNoConformidades`** | **44** |

### Lista completa de las 42 tablas

```
Copia de TbNCARAvisos (legacy copy)
TbAnexos
TbAnexosAuditoria
TbAnexosNCAuditorias
TbAuditoriaLog
TbAuditorias
TbAuxPuntoNorma
TbCacheIndicadoresAuditoriaDetalle
TbCacheIndicadoresAuditoriaHeader
TbCacheIndicadoresConfig
TbCacheIndicadoresProyectoDetalle
TbCacheIndicadoresProyectoHeader
TbCacheListadoNC
TbCacheListadoNCAuditoria
TbCacheNCProyecto
TbConexiones
TbConfiguracion
TbConsultasPorFechas
TbDocumentosAuditorias
TbEstadoCatalogo
TbHerramientaDocAyuda
TbLog
TbLogAuditoria
TbLogCache
TbNCAccionCorrectivas
TbNCAccionesRealizadas
TbNCARAvisos
TbNCAuditoriaAccionCorrectivas
TbNCAuditoriaAccionesRealizadas
TbNCDocumentos
TbNCInformacionRAC
TbNCInformacionRACAnexos
TbNoConformidades
TbNoConformidadesAuditoria
TbNoConformidadesIngresoPorLotesDetalle
TbNoConformidadesIngresoPorLotesPrincipal
TbNoConformidadesIngresoPorLotesTemporal
TbReplanificacionesAuditoria
TbReplanificacionesProyecto
TbTareasExplicaciones
TbTipologia
TbTiposNCProyectos
```

### Esquema real `TbNoConformidades` (44 columnas, tipos reales)

| Columna | Tipo DAO | Size | Required | Notas para PostgreSQL |
|---|---|---|---|---|
| `IDNoConformidad` | 4 (LongInteger) | 4 | true | `BIGSERIAL` PK |
| `Juridica` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CodigoNoConformidad` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `EsNoConformidad` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT TRUE` |
| `EXPEDIENTE` | 10 (Text) | 255 | true | `VARCHAR(255) NOT NULL` |
| `PROYECTO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `VEHICULO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `DESCRIPCION` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `CAUSA` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `ENTIDADRESPONSABLE` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `RESPONSABLETELEFONICA` | 10 (Text) | 50 | false | `VARCHAR(50) NULL` |
| `FECHAAPERTURA` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `FECHACIERRE` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `FPREVCIERRE` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `TIPO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `NOTAS` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Borrado` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` (soft delete) |
| `RequiereACR` | 1 (YesNo) | 1 | false | `BOOLEAN DEFAULT FALSE` |
| `ACR` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `MotivoBorrado` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `RequiereControlEficacia` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` — texto, no bool |
| `ControlEficacia` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `FechaControlEficacia` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `FechaPrevistaControlEficacia` | 8 (DateTime) | 8 | false | `DATE NULL` |
| `ResultadoControlEficacia` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `ConformeControlEficacia` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ inconsistente: debería ser BOOLEAN pero está como texto 'Sí'/'No' |
| `RESPONSABLECALIDAD` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `IDExpediente` | 4 (LongInteger) | 4 | false | `BIGINT NULL` (sin FK física a TbExpedientes) |
| `CodExp` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `Nemotecnico` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `JuridicaExp` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `RESPONSABLECALIDADExp` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CausaYAnalisRaiz` | 12 (Memo) | 0 | false | `TEXT NULL` |
| `Tipologia` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `IDProyecto` | 4 (LongInteger) | 4 | false | `BIGINT NULL` (sin FK física) |
| `CodigoRiesgo` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `DetectadoPor` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `ResponsableEjecucion` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `ESTADO` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `IDTipo` | 4 (LongInteger) | 4 | false | `BIGINT NULL` |
| `Cerrada` | 10 (Text) | 2 | false | `VARCHAR(2) NULL` — ⚠️ inconsistente: debería ser BOOLEAN |
| `IDNCAsociada` | 4 (LongInteger) | 4 | false | `BIGINT NULL` |
| `CodigoNoConformidadAsociada` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `CodConcesionAsociada` | 10 (Text) | 255 | false | `VARCHAR(255) NULL` |
| `MotivoNoRequiereControlEficacia` | 12 (Memo) | 0 | false | `TEXT NULL` |

### Relaciones físicas reales (14 FK entre user tables)

| Origen | Columna FK | Destino | Columna FK |
|---|---|---|---|
| `TbAuditorias` | `IDAuditoria` | `TbDocumentosAuditorias` | `IDAuditoria` |
| `TbAuditorias` | `IDAuditoria` | `TbNoConformidadesAuditoria` | `IDAuditoria` |
| `TbNCAccionCorrectivas` | `IDAccionCorrectiva` | `TbNCAccionesRealizadas` | `IDAccionCorrectiva` |
| `TbNCAccionesRealizadas` | `IDAccionRealizada` | `TbReplanificacionesProyecto` | `IDAccionRealizada` |
| `TbNCAuditoriaAccionCorrectivas` | `IDAccionCorrectiva` | `TbNCAuditoriaAccionesRealizadas` | `IDAccionCorrectiva` |
| `TbNCAuditoriaAccionesRealizadas` | `IDAccionRealizada` | `TbDocumentosAuditorias` | `IDAccionRealizada` |
| `TbNCInformacionRAC` | `IDInformacionRAC` | `TbNCInformacionRACAnexos` | `IDInformacionRAC` |
| `TbNoConformidadesAuditoria` | `ID` | `TbDocumentosAuditorias` | `IDNoConformidad` ⚠️ |
| `TbNoConformidadesAuditoria` | `ID` | `TbNCAuditoriaAccionCorrectivas` | `ID` ⚠️ |
| `TbNoConformidades` | `IDNoConformidad` | `TbAnexos` | `IDNoConformidad` |
| `TbNoConformidades` | `IDNoConformidad` | `TbNCAccionCorrectivas` | `IDNoConformidad` |
| `TbNoConformidades` | `IDNoConformidad` | `TbNCDocumentos` | `IDNoConformidad` |
| `TbNoConformidades` | `IDNoConformidad` | `TbNCInformacionRAC` | `IDNoConformidad` |

⚠️ **Inconsistencias detectadas en FKs**:

1. `TbNoConformidadesAuditoria.ID` (PK genérico) → `TbDocumentosAuditorias.IDNoConformidad` (no es FK por `ID`, es por `IDNoConformidad` semántico). Parece un data integrity gap: la FK está mal declarada en Access; en PostgreSQL conviene redefinir como `TbNoConformidadesAuditoria.ID → TbDocumentosAuditorias.IDNoConformidadAuditoria` o equivalente.
2. `TbNoConformidadesAuditoria.ID` → `TbNCAuditoriaAccionCorrectivas.ID` (PK-to-PK). Similar gap.
3. **NO hay FK física** entre `TbNoConformidades` y `TbAuditorias`. La relación es a través de `TbNoConformidadesAuditoria`. Esta indirección es intencional o un gap histórico — verificar con negocio.
4. **NO hay FK física** entre `TbNoConformidades` y `TbExpedientes` pese a tener `IDExpediente` y `CodExp`. La relación es conceptual/por código.

## Diccionario completo

El diccionario fuente completo, con **cada tabla y cada campo**, se obtendrá al ejecutar Dysflow `list_tables`, `get_schema`, `get_relationships` y `count_rows` sobre `NoConformidades_Datos.accdb` una vez resuelto D89. Esta página añade la **estructura por clases** y los **perfiles agregados** declarados en código; no se han copiado valores de filas.

### Entidades de dominio principales (extraídas del código)

| Clase | Tabla backend probable | Rol |
|---|---|---|
| `NCAuditoria.cls` / `NCaUDITORIAOperaciones.cls` | `TbNCAuditoria` / `TbNCAuditoriaOperaciones` (presumidos) | NC de Auditoría + operaciones |
| `NCProyecto.cls` / `NCProyectoOperaciones.cls` | `TbNCProyecto` / `TbNCProyectoOperaciones` (presumidos) | NC de Proyecto + operaciones |
| `NCProyectoListItemVM.cls` | viewmodel, sin tabla propia | vista flat para UI |
| `ACAuditoria.cls` / `ACAuditoriaOperaciones.cls` | `TbACAuditoria` (presumido) | Acciones Correctivas de Auditoría |
| `ACProyecto.cls` / `ACProyectoOperaciones.cls` | `TbACProyecto` (presumido) | Acciones Correctivas de Proyecto |
| `ARAuditoria.cls` / `ARAuditoriaOperaciones.cls` | `TbARAuditoria` (presumido) | Acciones Preventivas de Auditoría |
| `ARProyecto.cls` / `ARProyectoOperaciones.cls` | `TbARProyecto` (presumido) | Acciones Preventivas de Proyecto |
| `DocumentoAuditoria.cls` / `DocumentoAuditoriaOperaciones.cls` | `TbDocumentoAuditoria` (presumido) | Documentos de NC de Auditoría |
| `DocumentoProyecto.cls` / `DocumentoProyectoOperaciones.cls` | `TbDocumentoProyecto` (presumido) | Documentos de NC de Proyecto |
| `ReplanificacionesAuditoria.cls` / `ReplanificacionesAuditoriaOperaciones.cls` | `TbReplanificacionesAuditoria` (presumido) | Replanificaciones de Auditoría |
| `ReplanificacionesProyecto.cls` / `ReplanificacionesProyectoOperaciones.cls` | `TbReplanificacionesProyecto` (presumido) | Replanificaciones de Proyecto |
| `Auditoria.cls` / `AuditoriaOperaciones.cls` | `TbAuditoria` (presumido) | Cabecera de Auditoría |
| `Auditoria.cls` (no, `AuditoriaOperaciones`) | `TbAuditoriaOperaciones` (presumido) | Operaciones de Auditoría |
| `TipologiaNCProyectos.cls` | `TbTipologiaNCProyectos` (presumido) | Tipologías NC de proyectos |
| `Riesgo.cls` / `RiesgoServicio.cls` | tablas de Gestion_Riesgos (vinculación externa) | Vinculación NC ↔ Riesgo |
| `LogNCAuditoria.cls` / `LogNCProyecto.cls` | `TbLogNCAuditoria` / `TbLogNCProyecto` (presumidos) | Logs por NC |
| `SegNCAuditoria.cls` / `SegNCProyecto.cls` | `TbSegNCAuditoria` / `TbSegNCProyecto` (presumidos) | Seguridad por NC |
| `SegTareasAuditoria.cls` / `SegTareasProyecto.cls` | `TbSegTareasAuditoria` / `TbSegTareasProyecto` (presumidos) | Tareas de seguimiento |
| `Juridica.cls` | `TbJuridicas` (presumido) | Catálogo de jurídicas |
| `IndicadorServicio.cls` / `IndicadorRepositorio.cls` | `TbIndicadores` (presumido) | Indicadores |
| `Informe.cls` / `InformeNCAuditorias.cls` | sin tabla; genera artefacto | Informes |
| `Usuario.cls` / `UsuarioAplicacionPermisos.cls` / `Entorno.cls` | `TbUsuariosAplicaciones*` (Lanzadera, vía `getdbLanzadera` o tablas compartidas) | Identidad / permisos / entorno |
| `Expediente.cls` / `ExpedienteResponsable.cls` | tablas de Expedientes (compartidas con Lanzadera) | Vinculación con Expedientes |
| `CacheNCCacheRepositorio.cls` | `TbCacheNCProyecto` + `TbLogCache` | Capa de caché |
| `Correo.cls` | `TbCorreosEnviados` (presumido) | Correos |

### La capa de caché (referencia principal del blueprint)

NoConformidades tiene la implementación de caché más rica del ecosistema. Es la **referencia de diseño** para el puerto de caché de la nueva plataforma (ver D91).

**Tabla `TbCacheNCProyecto`** (columnas inferidas del código en `InicializadorCache.bas`):

| Columna | Tipo inferido | Rol |
|---|---|---|
| `IDCache` | Long (PK) | Identificador interno de la fila de caché |
| `IDNoConformidad` | Long (FK conceptual) | NC cacheada |
| `CacheValida` | Boolean (`True`/`False`) | Si el caché es válido o fue invalidado |
| `FechaCache` | Date | Cuándo se generó/regeneró |
| `FechaUltimoUso` | Date | Última vez que se leyó (alimenta `InvalidarCachesObsoletos`) |
| `HitsConsultas` | Long | Número de hits (alimenta `MostrarEstadisticasUso`) |
| `TamanioBytes` | Long | Tamaño en bytes (alimenta `MostrarRendimiento` por rango) |
| `Version` | String | Versión del caché (cambia cuando se regenera) |
| `DatosNC` | Memo (JSON) | Snapshot del NC |
| `DatosACs` | Memo (JSON) | Snapshot de las ACs de la NC |
| `DatosARs` | Memo (JSON) | Snapshot de las ARs de la NC |

**Tabla `TbLogCache`** (inferida):

| Columna | Tipo inferido | Rol |
|---|---|---|
| (PK) | Long | Identificador |
| `FechaOperacion` | Date | Cuándo ocurrió la operación |
| `TipoOperacion` | String (`Error` u otro) | Tipo |
| (Detalle) | String/Memo | Mensaje |

**Comandos de mantenimiento** (en `InicializadorCache.bas`):

| Comando | Efecto | Default |
|---|---|---|
| `InvalidarCachesObsoletos(dias)` | Marca `CacheValida=False` si `FechaUltimoUso` es más antigua que `dias` | 30 |
| `EliminarCachesInvalidos(dias)` | Borra físicamente cachés con `CacheValida=False` y más de `dias` días | 60 |
| `MostrarEstadisticasUso()` | Top 10 más consultados, top 10 más grandes, totales | n/a |
| `DiagnosticarIntegridad()` | Huérfanos (cachés sin NC), JSONs vacíos, NCs con ACs sin ARs | n/a |
| `RegenerarCachesInvalidos()` | Regenera solo los marcados como inválidos | n/a |
| `PoblarCacheMasivo(soloFaltantes)` | Puebla la caché para todas las NCs activas | `True` |
| `LimpiarLogsAntiguos(diasLogs, diasErrores)` | Borra logs antiguos (90 días) y errores (30 días) | 90, 30 |
| `MostrarRendimiento()` | Distribución de tamaño en rangos | n/a |
| `AyudaCache()` | Lista los comandos disponibles | n/a |

**Kill switch** (`Test_KillSwitch.bas`):

- `Test_KillSwitch_IsCacheEnabled_Atomic` — test atómico para activar/desactivar la caché en runtime.
- Usa `BuildJsonFail`/`BuildJsonOk`/`AddLog`/`EndTestSession` — el patrón de tests VBA ya está presente.
- El nombre "atomic" sugiere que el toggle debe ser transaccional (sin race conditions).

### Relaciones físicas confirmadas (en código)

- `TbCacheNCProyecto.IDNoConformidad → TbNoConformidades.IDNoConformidad` (FK conceptual, validada por `DiagnosticarIntegridad`).
- `TbLogCache` (sin FK directa observable en código; probablemente log por operación sin relación FK explícita).
- `TbConfiguracionBackends` ↔ `TbConfiguracion.CacheHabilitada` (consultada por `ResolveCacheHabilitadaFromConfig`).
- **La relación física entre `TbNCProyecto` y `TbRiesgos` no aparece como FK explícita en el código Access**; se modela por código de NC ↔ código de riesgo.

### Semántica Access que debe conservarse

- Tipos 1/3/4/7/8/10/12 observados: Boolean, Integer/Long, Currency, DateTime, Text y Memo según columna; confirmar mapeo final por campo en la segunda pasada Dysflow (post-D89).
- `Sí/No` se almacena frecuentemente como texto de longitud 2, no como Boolean; **no convertirlo silenciosamente**.
- `Null` y cadena vacía se distinguen en formularios, DTO y `Registrar`.
- IDs calculados con `DameID("TbXxx", "IDXxx", getdb(), m_Error)` antes de `.AddNew` — la nueva plataforma debe reproducir el patrón (idempotente, atómico).
- `IDAplicacion = "8"` (producción) / `"81"` (pruebas) se asigna por la TempVar `EnPruebas`; en la nueva plataforma viene de la configuración (D9).
- `EnPruebas`/`EnDesarrollo` debe ser texto 'Sí' o 'No' en `TbConfiguracionBackends` (verificación explícita en `LeeConfiguracionLocal:258-261`).

### Perfil agregado (privacidad segura)

El perfilado de filas por tabla queda **pendiente para la segunda pasada Dysflow** (post-D89). En esta primera pasada se cosecharon:

- 44 clases de dominio.
- **344 callers de `getdb()`** (el más alto de las 8 aplicaciones).
- 7 archivos `Test_*.bas` con cobertura significativa.
- Una capa de caché con kill switch, diagnóstico, métricas y logs.

Una vez completada la segunda pasada, este apartado se completará con `count_rows` y comprobaciones de unicidad, nulos, huérfanos, rangos de fechas y ratios de hits de caché.