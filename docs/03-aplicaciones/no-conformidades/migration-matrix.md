# NoConformidades — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. La capa de caché maduro se preserva como referencia (D91). El diagnóstico del fallo de inventario Dysflow se aborda en [D89](security-rules.md#d89-diagnóstico-del-fallo-de-list_objects-de-dysflow) y el hallazgo de seguridad en [D90](security-rules.md#d90-riesgo-de-seguridad--backendsjson-con-contraseña-en-claro).

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbNCAuditoria` / `TbNCProyecto` (presumidos) | NC de Auditoría / NC de Proyecto | separar identidad, fechas, estado,Flags de AC/AR,Flags de replanificación | mapeado preliminar | recuento 1:1, IDs reconciliados | catálogo de estados final |
| `IDNC` (autonumérico) | identificador de NC | conservar; normalizar a BIGINT en PostgreSQL | mapeado preliminar | unicidad | sin FK física explícita con `TbExpedientes` (vinculación por código) |
| `CodigoNC` / `Nemotecnico` | identificadores funcionales | conservar columnas originales y normalizadas | necesita decisión | unicidad y vacíos | claves canónicas |
| Flags AC/AR y estados | clasificación y progreso | tabla de valores + valor legacy | mapeado preliminar | coherencia estado ↔ AC/AR | fuente de verdad |
| `TbACAuditoria` / `TbARAuditoria` (presumidos) | AC/AR de NC de Auditoría | FK a NC + responsable + fechas + motivo | mapeado preliminar | cardinalidad | versionado de AC/AR |
| `TbACProyecto` / `TbARProyecto` (presumidos) | AC/AR de NC de Proyecto | análogo | mapeado preliminar | análogo | análogo |
| `TbControlEficacia*` (presumidos) | veredicto del control | FK + estado + motivo | mapeado preliminar | secuencialidad | retención |
| `TbReplanificaciones*` (presumidos) | replanificación con motivo | FK + motivo + fecha + autor | mapeado preliminar | secuencialidad; vínculo con AC/AR | retención |
| `TbDocumento*` (presumidos) | documentos vinculados | metadatos + referencia externa | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino (D16) |
| `TbAuditoria` (presumido) | cabecera de Auditoría | identidad + responsable + fechas | mapeado preliminar | unicidad | versionado |
| `TbTipologiaNCProyectos` (presumido) | tipología NC | catálogo | mapeado preliminar | unicidad | versionado |
| `TbLogNCAuditoria` / `TbLogNCProyecto` (presumidos) | log de auditoría | evento + timestamp + actor | mapeado preliminar | trazabilidad | retención/legal |
| `TbSegNCAuditoria` / `TbSegNCProyecto` (presumidos) | seguridad por NC | reglas de visibilidad/edición | mapeado preliminar | coherencia con `UsuarioAplicacionPermisos` | modelo de capabilities (D45-D46) |
| `TbSegTareasAuditoria` / `TbSegTareasProyecto` (presumidos) | worklists | entidad de tarea + estado | mapeado preliminar | cardinalidad con NC | catálogo de tareas |
| `TbIndicadores` (presumido) | métricas | entidad de indicador + valor + periodo | mapeado preliminar | unicidad por periodo | versionado |
| `TbJuridicas` (presumido) | catálogo de jurídicas | entidad de catálogo | mapeado preliminar | unicidad; coherencia con Lanzadera | versionado |
| **`TbCacheNCProyecto`** (presumido) | caché selectivo de NC | **preservar como referencia del puerto de caché de la nueva plataforma** (D91) | mapeado preliminar | ratios de hits; diagnóstico de integridad; logs | traducción a adapter: invalidación por TTL + métricas Prometheus + kill switch operativo |
| **`TbLogCache`** (presumido) | log de operaciones de caché | evento + tipo + timestamp | mapeado preliminar | trazabilidad | traducción a logs estructurados canónicos (D27) |
| `TbConfiguracion` / `TbConfiguracionBackends` (presumidos) | configuración de backend | **eliminar acoplamiento directo**; sustituir por adaptador de configuración (D9-D10) | mapeado preliminar | coherencia | migración a config del puerto de persistencia |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera` o tablas compartidas) | identidad / permisos / entorno | identidad vía adaptador unificado (D9-D10) | mapeado preliminar | coherencia | acoplamiento directo a romper (D86/D87) |
| tablas E2E/temp/aux (si las hay) | soporte técnico | migrar solo si se requiere trazabilidad; conservar dump si no | technical-only candidate | documentar antes de excluir | decisión explícita, no obsolescencia |
| tablas copia/históricas/auxiliares | histórico o soporte | conservar en zona legacy hasta clasificación | needs business decision | comparar recuentos y uso en código | archivo vs dominio |
| integraciones externas (Lanzadera, Expedientes, Gestion_Riesgos, HPS, NC, AGEDYS, correos) | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera, AC/AR, replanificaciones, documentos, logs, seguimientos, caché, catálogos.
- **Necesita decisión:** flags frente a tipo, versionado de AC/AR, retención, control de eficacia.
- **Candidato técnico:** temporales y helpers de exportación; no retirado.
- **Referencia externa:** documentos, identidad, Lanzadera, Expedientes, Gestion_Riesgos, HPS, NC y correo.
- **No hay campos/filas declarados obsoletos.**

## D89 · Diagnóstico del fallo de `list_objects` de Dysflow — INVALIDADO

**Estado**: PROPUESTO. **INVALIDADO post-verificación con Dysflow live (2026-08-05).**

NoConformidades tiene **42 tablas** con **438 NCs de Proyecto y 55 NCs de Auditoría**. El inventario Dysflow funciona perfectamente. El "fallo de inventario" del que hablaba la versión original de este documento era una diagnosis errónea basada en análisis estático sin haber contactado el runtime. La secuencia correcta de invocación está documentada en [Seguridad § D89](security-rules.md#d89--diagnóstico-del-fallo-de-list_objects-de-dysflow--invalidado) y replicada abajo para referencia.

### Resultado real del inventario (2026-08-05)

| Categoría | Resultado |
|---|---|
| Tablas en `NoConformidades_Datos.accdb` | **42** |
| Filas en `TbNoConformidades` | **438** |
| Filas en `TbNoConformidadesAuditoria` | **55** |
| FK relationships (user tables) | **14** |
| Columnas en `TbNoConformidades` | **44** |

### Secuencia correcta de invocación Dysflow (referencia)

```js
// 1. Cargar skills primero (regla cross-project)
load_skill("dysflow-usage");
load_skill("dysflow-arnes");

// 2. Bootstrap
dysflow.get_capabilities({});
dysflow.register_worktree({cwd: "<worktree-root>"});
// o dysflow.resolve_project({cwd, projectId, projectChoiceReason, recoveryToken}) si hay ambigüedad

// 3. Migrar config si tiene allowWrites top-level
dysflow.migrate_project_config({cwd, apply: true});

// 4. Inventario (siempre con accessPath absoluto)
dysflow.list_objects({cwd, accessPath: "<abs/path/frontend.accdb>", outputMode: "summary"});
dysflow.list_tables({cwd, accessPath: "<abs/path/frontend.accdb>", outputMode: "full"});
dysflow.get_relationships({cwd, accessPath: "<abs/path/frontend.accdb>"});
dysflow.count_rows({cwd, accessPath: "<abs/path/frontend.accdb>", table: "<tableName>"});
dysflow.get_schema({cwd, accessPath, table: "<tableName>"});
```

### Por qué fallaba la invocación inicial

La invocación inicial usaba solo `cwd` sin `accessPath` absoluto. Dysflow retornaba `CONFIG_MISSING_ACCESS_PATH` no porque el config estuviera mal, sino porque algunas herramientas (`list_tables`, `get_relationships`) requieren el path absoluto del frontend como parámetro explícito. `list_objects` aceptaba `cwd` solo; `list_tables` no.

### Lecciones aprendidas (issue filed: DysTelefonica/dysflow#1403)

Ver el maintainer-prompt archivado en `docs/prompts/prompt-ia-mantenedora-dysflow-round-1-2026-08-05.md` y la issue filed en https://github.com/DysTelefonica/dysflow/issues/1403.

## D90 · Riesgo de seguridad: `backends.json` con contraseña en claro

**Estado**: PROPUESTO. El archivo `00_main/backends.json` contiene una entrada `ACCESS_VBA_PASSWORD` con valor en claro. Este archivo está versionado en git.

Riesgos concretos:

- **Exposición de credenciales en el historial de git**: cualquier push al remoto expone la contraseña. Si el repo es público o semi-público, el daño es inmediato.
- **Riesgo de rotación nunca ejecutada**: si el equipo cambia la contraseña real del VBA, `backends.json` queda desfasado y el binario falla al abrirse con esa contraseña (paradoja: la versión vieja en git puede no coincidir con la actual).
- **Uso indebido por terceros**: si la contraseña es la real, cualquiera con acceso al repo puede abrir el binario y leer datos de NC.

Recomendación operativa:

1. **Rotar la contraseña VBA del binario** antes de cualquier remediación.
2. **Eliminar `backends.json` del repo** (git rm + .gitignore + commit de limpieza).
3. **Reemplazar por `backends.example.json`** sin valor, con comentarios que documenten las claves esperadas (`ACCESS_VBA_PASSWORD`).
4. **Forzar el uso de env vars**: el código VBA ya lee `Environ$("ACCESS_VBA_PASSWORD")` (`Variables Globales.bas:250`); basta con que el runner de Dysflow y los entornos de desarrollo/staging/producción exporten la variable sin que esté en el repo.
5. **Documentar D90 en el blueprint** como hallazgo transversal: cualquier credencial versionada en `backends.json` o equivalente es un riesgo a corregir lote por lote.

## D91 · Caché selectivo maduro como referencia del puerto de caché

**Estado**: PROPUESTO. La capa de caché de NoConformidades (`TbCacheNCProyecto`, `InicializadorCache.bas`, kill switch, diagnóstico, métricas, logs) se preserva **íntegramente** como referencia de diseño para el puerto de caché de la nueva plataforma (D70-D71).

Mapeo a la nueva plataforma:

| Concepto legacy | Concepto nueva plataforma | Notas |
|---|---|---|
| `TbCacheNCProyecto` | Tabla `cache_nc_proyecto` en esquema del módulo (o del puerto de caché si se centraliza) | Preservar columnas; añadir `tenant_id` si multi-tenant |
| `TbLogCache` | Logs estructurados canónicos (D27) con `correlation_id` | Vía el logger canónico de la plataforma |
| Kill switch (`Test_KillSwitch_IsCacheEnabled_Atomic`) | Toggle runtime vía variable de entorno o endpoint admin | Mantener capacidad de desactivación atómica |
| Comandos de mantenimiento (`InvalidarCachesObsoletos`, `DiagnosticarIntegridad`, etc.) | Jobs del scheduler unificado (D59) o endpoints admin | Traducir a casos de uso Python; mantener las mismas semánticas |
| Métricas (`HitsConsultas`, `TamanioBytes`, `FechaUltimoUso`) | Métricas Prometheus / OpenTelemetry | Exponer vía puerto de observabilidad (D27) |
| `CacheValida` (Boolean) | `EstadoCache` (enum: Válido, Inválido, Regenerándose) | Más expresivo en el modelo nuevo |
| `CacheHabilitada` (de `TbConfiguracion`) | Config del módulo / feature flag | Alineado con D45 (capabilities) y la decisión de caché selectiva (D70) |

## D103 · Unificación de projectIds Dysflow en los 4 worktrees

**Estado**: APLICADO el 2026-08-05.

### Hallazgo original

Los 4 worktrees de `00_NO_CONFORMIDADES` (`00_main`, `hotfix-replanificadas`, `slice10`, `staging`) tenían el **mismo `projectId: "00-no-conformidades-staging-clean"`** en sus `.dysflow/project.json`. Esto causaba `FRONTEND_TARGET_AMBIGUOUS` en `resolve_project` y violaba HR-11 del arnés dysflow (un `projectId` único por worktree).

### Fix aplicado

Se aplicó `setup_project` con `apply: true` y `projectId` distintos para unificar los IDs:

| Worktree | `projectId` (antes) | `projectId` (después) | `name` (después) |
|---|---|---|---|
| `00_main` | `00-no-conformidades-staging-clean` | `00-no-conformidades-00-main-clean` | (presumido) "NoConformidades 00_main clean" |
| `hotfix-replanificadas` | `00-no-conformidades-staging-clean` | `00-no-conformidades-hotfix-replanificadas-clean` | (presumido) "NoConformidades hotfix-replanificadas clean" |
| `slice10` | `00-no-conformidades-staging-clean` | `00-no-conformidades-slice10-clean` | (presumido) "NoConformidades slice10 clean" |
| `staging` | `00-no-conformidades-staging-clean` | `00-no-conformidades-staging-clean` ⚠️ (mantenido) | "NoConformidades staging clean" |

**Resultado**: `resolve_project` ahora lista los 4 projectIds distintos y el agente debe elegir uno por llamada (HR-11 comportamiento correcto). La ambigüedad persiste por diseño (múltiples proyectos visibles) pero cada proyecto es único.

### Recomendación cross-cutting

Aplicar el mismo patrón a cualquier repo que tenga múltiples worktrees hermanos con el mismo `.dysflow/project.json`. La regla HR-11 del arnés dysflow lo exige: **un `projectId` único por worktree**.

⚠️ **Mismo principio que en Lanzadera, Expedientes, Gestion_Riesgos, Condor, HPS, HPS_Solicitudes, Brass**: auditar si tienen múltiples worktrees con mismo projectId y unificarlos antes de operar.

Cada fila futura debe expandirse a `source table.field` para los 49 esquemas (cuando estén cosechados post-D89), con transformación, regla de reconciliación, rechazo y disposición aprobada.