# Épica — Condor (migración a web)

> **Estado:** DRAFT — pendiente revisión.
> **Versión:** v0.1 (2026-08-06).
> **Autor:** placeholder.
> **Lote de discovery:** 5 (junto al resto de las 7 apps restantes).
> **Cross-refs:** `docs/03-aplicaciones/condor/{capabilities,data-model,forms,integrations-automation,migration-matrix,security-rules,README}.md` · engram obs #24085 (audit) · engram obs #24086 (corrección `tbHistorialRechazos`).

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_CONDOR` · frontend `CONDOR.accdb` (44 MB) + backend `condor_datos.accdb` (5 MB) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (15 tablas, 8 features de negocio, 52+ clases, 4 tipos de Solicitud) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, vía `getdbLanzadera`); Expedientes (FK conceptual `idExpediente`); NoConformidades (FK conceptual `idNCAsociada`); AGEDYS y Gestion_Riesgos/HPS presumiblemente por código compartido (no inspeccionado en detalle) |
| **Riesgo dominante** | Seguridad D93 — contraseña `"dpddpd"` hardcodeada como fallback en `GetPasswordDB` |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ obs #24085 (15/15 ACTIVE, 0 ZOMBIE, 0 UNCLEAR) |

## Forma del documento

1. Scope (en / fuera)
2. Estado del descubrimiento
3. Hallazgos críticos
4. Decisiones aplicadas
5. Criterios de aceptación
6. Pendientes operacionales
7. Tickets derivables (preview)

---

## 1. Scope

### 1.1 En scope

**8 features de negocio** con paridad funcional en la nueva plataforma:

| # | Feature | Respaldo en docs |
|---|---|---|
| F1 | Solicitudes: alta, edición, baja, cambio de estado, snapshot, búsqueda y filtros avanzados | `capabilities.md` · `forms.md` · `migration-matrix.md` |
| F2 | 4 tipos de Solicitud con datos específicos: `PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB` | `data-model.md` (tbDatosPC/PCSUB, tbDatosCDCA/CDCASUB) · `forms.md` · `capabilities.md` |
| F3 | Workflow declarativo con transiciones y roles requeridos | `migration-matrix.md` § D96 · `data-model.md` § tbTransiciones |
| F4 | Recepción de rechazo e historial de resoluciones | `capabilities.md` · `data-model.md` § tbRechazos + tbHistorialRechazos · `forms.md` |
| F5 | Validación de calidad con hash de integridad | `data-model.md` § tbValidacionRevision (`HashDatos`) |
| F6 | Adjuntos asociados a etapa del workflow | `data-model.md` § tbAdjuntos (`etapaWF`, `TipoAccion`) |
| F7 | Vinculación con Lanzadera/Expedientes vía adaptadores unificados (D9-D10) | `forms.md` · `integrations-automation.md` · `migration-matrix.md` § D94 |
| F8 | Vinculación con NoConformidades vía adaptador (orden de migración estricto) | `forms.md` · `migration-matrix.md` § D95 |

**3 catálogos config seed-only** (datos de runtime, sin CRUD de usuario):

| Catálogo | Volumen | Notas |
|---|---|---|
| `estados` (era `tbEstados`) | 9 filas | Estados del workflow + flags `esEstadoInicial`/`esEstadoFinal` |
| `mapeo_campos` (era `tbMapeoCampos`) | **183 filas** | Reglas de merge legacy → moderno; service `MapeoServicio` ya existe (D97) |
| `transiciones` (era `tbTransiciones`) | TBD | Definidas por origen + destino + `rolRequerido`; motor de workflow data-driven (D96) |

**Mapeo de campos legacy → moderno**: 183 reglas en `tbMapeoCampos` se migran como `INSERT` en la inicialización; el servicio de merge se traduce a backend hexagonal.

**Vinculaciones externas**: la vinculación a Expedientes y NoConformidades se media por los **adaptadores unificados** definidos por D9-D10 (autenticación, identidad, permisos). No acoplamiento directo entre servicios.

**Snapshots web equivalentes al Edge WebView legacy**: `SnapshotServicio` + `WebVisorCacheServicio` se traducen a endpoints de snapshot server-side con autorización (D27, D45) y fragment caching con ETag (D72).

**Búsqueda y filtros**: `FiltrosSolicitud` con 3 callers (`Form_frm0PpalTecnico`, `Form_frmBuscarSolicitudes`, `Form_frmFiltrosAvanzadosSolicitudes`) se traduce a queries parametrizadas más UI HTMX.

### 1.2 Fuera de scope (REPLACE)

> Las 3 tablas puras de log **NO migran como tablas PostgreSQL**. Se reemplazan por stack de observabilidad web-native (Sentry / OpenTelemetry / structured logs a Loki o CloudWatch). Sus llamadas VBA se traducen a eventos web.

- `tbLogCambios` (5 filas en staging) — auditoría general con campo `suplantadoPor` (impersonación).
- `tbLogErrores` (3 filas en staging) — errores con `suplantadoPor`.
- `tbLogEstados` (0 filas en staging) — transiciones de estado (subsumido por el motor de workflow declarativo que registra su propio historial en stack de logs).

**Excluido del REPLACE**: `tbHistorialRechazos` es **tabla de negocio**, no log (corrección obs #24086). Mantiene esquema completo en PostgreSQL con CRUD en web.

**Persistencia**: engram topic_key `condor/log-strategy-2026-08-05`.

### 1.3 Fuera de scope (no documentado)

Si aparece algo que no está en los 7 docs de Condor, se marca como **no documentado** y se acumula en pendientes de discovery para iteración posterior. Ejemplos conocidos:

- Volúmenes reales de producción (no staging).
- `tbConfiguracion` y `tbConfiguracionBackends` presumidos en `00_main` pero no inspeccionados en staging.
- Mapeo semántico profundo entre columnas legacy y modernas más allá de `tbMapeoCampos`.
- Integraciones con AGEDYS, Gestion_Riesgos y HPS (vía código compartido, no inspeccionado).

---

## 2. Estado del descubrimiento

### 2.1 Inventario

| Categoría | Resultado |
|---|---|
| **Tablas totales** | **15** en `condor_datos.accdb` |
| **Filas totales (auditadas)** | ≈**305** (volumen real, incluye estimaciones de staging y referencias en producción) |
| **Schemas documentados** | **15/15** tablas en `data-model.md` (commits `d23d256`, `85cfc2a`, `68cc8aa`) |
| **Uso (audit obs #24085)** | **15/15 ACTIVE** · **0 ZOMBIE** · **0 UNCLEAR** |
| **Subcategorías** | 8 business feature · 4 audit/history · 3 config |
| **FKs físicas** | 5 desde `tbSolicitudes` (todas intra-app, parciales) |
| **Clases inventariadas** | 52+ (15 domain, 4 domain-data, 7 ViewModels, 21 Servicios, 3 Sandbox, mocks, errores) |
| **Forms** | ~30-40 archivos `Form_*.cls` (inventario completo en iteración posterior) |
| **Tests VBA** | `tests/` (cobertura probable significativa, confirmado por D87 — referencia para pytest) |

### 2.2 Herramientas usadas

- **Dysflow MCP** (read-only) sobre `CONDOR.accdb` + `condor_datos.accdb`: inventario real, conteo de filas, FKs.
- **codegraph-vba** sobre `C:\00repos\codigo\00_CONDOR\staging\.codegraph-vba` (14 MB): audit de uso (obs #24085) + extracción de call paths.
- **codegraph_explore**: blast radius de símbolos clave (`EVE`, `getdb`, `getUsuario`, `FiltrosSolicitud`, `WorkflowServicio`, `SnapshotServicio`, `NoConformidadServicio.getNoConformidadPorCodigoCondor`).

### 2.3 Pendientes menores de discovery

| Pendiente | Impacto | Iteración |
|---|---|---|
| Volúmenes reales de producción (no staging) | Ajustar criterios de aceptabilidad de migración de datos | Cuando se acceda a producción |
| Inventario completo de forms/queries/macros embebidos | Codegraph no captura macros ni QueryDefs no exportados | Próxima iteración |
| Cadena de uso de `tbHistorialRechazos` por saved queries / macros | Audit es lower bound; obs #24086 ya marcó como NEGOCIO | Inspección manual si negocio requiere |

---

## 3. Hallazgos críticos

Cada hallazgo: descripción corta + impacto + acción + referencia documental.

<a id="hallazgo-D93"></a>
### H1 · D93 — Contraseña `"dpddpd"` hardcodeada como fallback en `GetPasswordDB`  ⚠️⚠️⚠️ CRÍTICO

- **Descripción**: `FUNCIONES UTILES.bas:150` contiene `GetPasswordDB = "dpddpd"` como fallback si la lectura del INI falla. La contraseña está expuesta en código fuente; es la misma para todos los entornos (PROD, SANDBOX, TEST); si fue rotada, el fallback conserva la versión vieja.
- **Impacto**: bypass de rotación, exposición en repo, incumplimiento de HR-3 (cero secretos en código). El mismo patrón puede existir en Lanzadera, Gestion_Riesgos, NoConformidades y HPS — D104 complementa a D93.
- **Acción**: eliminar fallback; fallar explícitamente con `Err.Raise`; rotar contraseña real del binario; mover a secret manager (D9-D10); auditar git history por commits con la cadena `"dpddpd"`. Acción operativa documentada en § 6.
- **Detalle completo**: [`migration-matrix.md § D93`](migration-matrix.md#d93--password-hardcodeado-como-fallback-en-getpassworddb) · [`security-rules.md § D93`](security-rules.md#d93--password-hardcodeado-como-fallback-en-getpassworddb).

<a id="hallazgo-D94"></a>
### H2 · D94 — FKs conceptuales sin constraint

- **Descripción**: `tbSolicitudes.idEstadoInterno`, `tbSolicitudes.idExpediente`, `tbSolicitudes.idNCAsociada`, `tbTransiciones.idEstadoOrigen`/`idEstadoDestino`, `tbLogEstados.idEstadoAnterior`/`idEstadoNuevo` y las FKs de `tbRechazos`, `tbAdjuntos`, `tbValidacionRevision`, `tbHistorialRechazos` a `tbSolicitudes` **no tienen constraint físico**. Data integrity gap.
- **Impacto**: integridad referencial no garantizada en legacy; en PostgreSQL hay que formalizar las FKs intra-app y mantener las FKs cross-app como referencias conceptuales (D86/D87).
- **Acción**: en PostgreSQL, formalizar todas las FKs intra-app con `FOREIGN KEY` reales; las FKs cross-app (`idExpediente`, `idNCAsociada`) se mantienen como referencia conceptual mediate adaptadores.
- **Detalle completo**: [`migration-matrix.md § D94`](migration-matrix.md#d94--fks-conceptuales-sin-constraint).

<a id="hallazgo-D95"></a>
### H3 · D95 — Vinculación con NoConformidades vía código (no por FK directa)

- **Descripción**: `Form_frmGestionSolicitud.cls:1908` verifica la NC antes de eliminar vía `ncServ.getNoConformidadPorCodigoCondor(codigoSolicitud)`. La relación es por código de Solicitud, no por `idNCAsociada` directo. La verificación depende de Lanzadera/Expedientes.
- **Impacto**: orden de migración estricto — NoConformidades debe migrarse **antes** que Condor para que la verificación tenga contraparte. Si hay NCs vinculadas, deben migrarse antes que las Solicitudes que las referencian.
- **Acción**: migración en dos olas (NoConformidades primero, Condor después); mantener FK conceptual mediate adaptador; documentar regla "idNCAsociada solo cuando hay NC explícitamente vinculada".
- **Detalle completo**: [`migration-matrix.md § D95`](migration-matrix.md#d95--vinculación-con-noconformidades).

<a id="hallazgo-D96"></a>
### H4 · D96 — Workflow declarativo vía `tbTransiciones`

- **Descripción**: `tbTransiciones` (4 columnas: `idTransicion`, `idEstadoOrigen`, `idEstadoDestino`, `rolRequerido`) define el workflow **como datos**, no como código. `WorkflowServicio.cls` lo lee en runtime.
- **Impacto**: el motor de workflow es data-driven y soporta evolución del workflow sin redeploy de código; en PostgreSQL hay que preservar este patrón como `transiciones` con FKs explícitas a `estados(origen)` y `estados(destino)`, y traducir `rolRequerido` a verificación de capabilities (D45-D46).
- **Acción**: seed de `transiciones`; endpoint admin para evolucionar el workflow; endpoint "transiciones disponibles" para un usuario/estado dado; verificación de capabilities server-side.
- **Detalle completo**: [`migration-matrix.md § D96`](migration-matrix.md#d96--workflow-declarativo-via-tbtransiciones).

<a id="hallazgo-D97"></a>
### H5 · D97 — `tbMapeoCampos` con 183 filas de config legacy → moderno

- **Descripción**: `tbMapeoCampos` define 183 reglas que mapean columnas legacy a campos en documentos Word (sistema de merge de plantillas). Es **config que se preserva como datos**, no código.
- **Impacto**: las 183 reglas son un activo del sistema que debe sobrevivir la migración como datos de runtime. Servicio `MapeoServicio` ya existe; la nueva plataforma debe exponer UI admin para mantenerlo sin deploys (bajo control de capabilities D45).
- **Acción**: `INSERT INTO tb_mapeo_campos VALUES (...)` en migración inicial; endpoint admin CRUD; decisión abierta con negocio sobre cómo renderizar (Word vs Jinja2/PDF).
- **Detalle completo**: [`migration-matrix.md § D97`](migration-matrix.md#d97--tbmapeocampos-con-183-filas--config-que-se-preserva-como-datos).

### Hallazgos cross-cutting que aplican a Condor

<a id="hallazgo-D102"></a>
### H6 · D102 — Booleanos como `Text(2)` 'Sí/No' (cross-cutting)

- **Descripción**: varias apps usan `Text(2)` con valores `'Si'`/`'No'` en columnas booleanas; en PostgreSQL debe estandarizarse a `BOOLEAN`. En Condor las columnas booleanas explícitamente documentadas (`esEstadoInicial`, `esEstadoFinal`, `esActivo`, `EstaResuelto`) son `YesNo` real, así que D102 **no es bloqueante para Condor**; aplica como revisión preventiva en columnas no inspeccionadas.
- **Impacto**: inconsistencias residuales si alguna columna descubierta en `00_main` usa el patrón. Migración debe cubrir la regla general.
- **Acción**: al descubrir columnas booleanas adicionales, validar su tipo DAO; si es `Text(2)`, traducir a `BOOLEAN` con regla explícita (`'Si' → TRUE`, `'No' → FALSE`, otros → NULL).

<a id="hallazgo-D104"></a>
### H7 · D104 — Patrón de secret manager (cross-cutting)

- **Descripción**: el patrón de contraseña hardcodeada puede existir también en otras apps (Brass D104 ya documentado como contraseña **REAL** directa, no fallback). Aunque D93/D104 no son la misma instancia, comparten la misma mitigación: secret manager.
- **Impacto**: descubrimiento paralelo necesario; la regla HR-3 aplica a todos los repos consumer.
- **Acción**: la migración de Condor no debe esperar al audit cross-cutting; D93 se cierra con rotación de la contraseña real y migración a secret manager (D9-D10) en la ola de Condor.

---

## 4. Decisiones aplicadas

> Las decisiones que ya están tomadas y que esta épica respeta sin复议ar.

<a id="dec-logs-web-native"></a>
### D-Logs · 3 tablas puras → REPLACE web-native

- **Alcance**: `tbLogCambios`, `tbLogErrores`, `tbLogEstados` NO migran como tablas PostgreSQL. Se reemplazan por stack de observabilidad canónico (Sentry / OpenTelemetry / structured logs).
- **Motivación**: tooling moderno de observabilidad con rotación/retención resuelta por el stack; queries de logs vía UI/SQL del backend de logs en vez de PostgreSQL.
- **Trade-off explícito del usuario (opción B)**: hay que traducir las llamadas VBA existentes a eventos web; se pierde la posibilidad de JOINs SQL directos sobre logs.
- **Persistencia**: engram topic_key `condor/log-strategy-2026-08-05`.
- **Ref**: [`migration-matrix.md`](migration-matrix.md#decisión-d-new-2026-08-05--logs-a-stack-web-native).

### D-Historial · `tbHistorialRechazos` = tabla de NEGOCIO (no log)

- **Alcance**: `tbHistorialRechazos` mantiene esquema completo en PostgreSQL con CRUD en web. Es histórico de resoluciones de rechazo.
- **Por qué**: corrección del audit (obs #24086) — el usuario confirmó que es contenido de dominio, no log. Acepta el lower bound del codegraph.
- **Diferencia con `tbRechazos`**: `tbRechazos` es cabecera activa del rechazo; `tbHistorialRechazos` es log histórico de resoluciones.

### D-FK-intra · FKs conceptuales intra-app → formales en PostgreSQL

- **Aplicación de D94**: `tbSolicitudes.idEstadoInterno → tbEstados`, `tbTransiciones.idEstadoOrigen/Destino → tbEstados` se formalizan con `FOREIGN KEY` real. `idExpediente` e `idNCAsociada` se mantienen conceptuales (D86/D87, mediate adaptadores).

### D-Identidad · Vía adaptador unificado (D9-D10)

- **Aplicación**: la identidad se resuelve por el adaptador unificado de autenticación, no por acceso directo a la BD de Lanzadera (`getdbLanzadera`). `IDAplicacion = "23"` queda como config del módulo.
- **Implicación**: el módulo Condor no conoce la BD de Lanzadera; el adaptador de identidad expone `getUsuarioConPermisos(usuario, aplicacion)` server-side.

### D-Backend-config · `TbConfiguracionBackends` → config del módulo hexagonal

- **Aplicación**: `BackendActivo`, `EnPruebas`, `EnDesarrollo` y similares dejan de ser TempVars o flags de runtime; pasan a **config del módulo hexagonal** (variables de entorno del runner). `PasswordBackend` desaparece (D93).
- **Detalle**: [`integrations-automation.md`](integrations-automation.md#configuración-y-flags-de-operación).

### D-Testing · `m_TestingMode` + `m_BackendSandboxURL` como referencia

- **Aplicación**: el patrón maduro de Condor se preserva como **referencia del puerto de testing** del puerto de caché y del runner de pytest. Cache safety (Spec-008) y fail-fast "TESTS BLOCKED" se traducen al adapter de testing.
- **Detalle**: [`integrations-automation.md`](integrations-automation.md#testing-sandbox-seguro-patrón-único-de-condor).

### D-MVVM-traductor · ViewModels/Servicios/Repositorios → Pydantic/use-cases/adaptadores

- **Aplicación**: la separación MVVM en VBA se traduce uno a uno — `*ViewModel` → DTOs Pydantic, `*Servicio` → use cases Python, `*Repositorio` → adaptadores de persistencia detrás del puerto.

### D-Workflow · Motor de workflow server-side + transiciones como datos

- **Aplicación**: `WorkflowServicio` se traduce a motor de workflow server-side; `tbTransiciones` se carga desde la tabla `transiciones` con FKs explícitas; las capabilities (D45-D46) validan `rol_requerido` antes de ejecutar la transición.

### D-WebView · Snapshots web equivalentes al Edge WebView legacy

- **Aplicación**: `SnapshotServicio` + `WebVisorCacheServicio` se traducen a endpoints de snapshot server-side con autorización (D27, D45); HTML se sirve autenticado; se elimina el patrón de escribir `.html` en disco del cliente. Fragment caching con ETag (D72).

### D-Mapeo · `tbMapeoCampos` con 183 filas como datos de runtime

- **Aplicación**: las 183 filas son datos; se migran como `INSERT` y se exponen vía servicio + endpoint admin. Decisión abierta con negocio sobre el render (Word / Jinja2 + HTML / WeasyPrint).

---

## 5. Criterios de aceptación

Lista verificable de qué define "épica de Condor cerrada".

### Funcionales (paridad con legacy)

- [ ] Las 8 features de negocio (F1-F8) tienen CRUD en web con paridad funcional mínima (ver § 1.1).
- [ ] Los 4 tipos de Solicitud (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`) tienen formularios diferenciados (vista única por defecto, especialidades justificadas si difieren materialmente — D46).
- [ ] El workflow declarativo se carga desde `transiciones` (no hardcoded); las transiciones devueltas por el endpoint matchean el ciclo esperado.
- [ ] Las capabilities de Condor están declaradas y verifican `rol_requerido` antes de cada transición (D45-D46).
- [ ] `tbRechazos` y `tbHistorialRechazos` migran como tablas de negocio con CRUD completo.
- [ ] `tbValidacionRevision` migra con `HashDatos` preservado.
- [ ] `tbAdjuntos` migra con metadatos (`etapaWF`, `TipoAccion`); el contenido se externaliza a object storage (D16); papelera 30 días (D19).
- [ ] Las vinculaciones con Expedientes y NoConformidades operan via adaptadores unificados (D9-D10); verifican existencia antes de acciones destructivas.

### Config / catálogos

- [ ] Los 3 catálogos (`estados`, `transiciones`, `mapeo_campos`) están migrados como seed inicial con sus volúmenes originales (9, TBD, 183).
- [ ] Endpoint admin CRUD sobre `mapeo_campos` con control de capabilities.
- [ ] Endpoint admin para evolucionar `transiciones` sin deploy.

### Forma y plataforma

- [ ] Módulo hexagonal dentro del monolito modular (D68); puertos por capacidad.
- [ ] Adaptadores driving: web (HTMX) + CLI admin global (D24).
- [ ] Adaptadores driven: PostgreSQL, object storage S3-compatible (D16), secret manager (D9-D10), notificación unificada (D11).
- [ ] Versión semántica del módulo: `condor/v0.1.0-rc.1` en primer RC, `condor/v1.0.0` en release (D78).
- [ ] Migraciones de BD backward-compatibles con Expand and Contract (D82).

### Seguridad y operación

- [ ] Logs en stack web-native (Sentry/OTel); **NO** existen `tb_log_cambios`/`tb_log_errores`/`tb_log_estados` en PostgreSQL (verificación binaria).
- [ ] Password `"dpddpd"` **eliminada** del código; secret en secret manager; rotación documentada (D93 cerrado).
- [ ] Suplantación acotada al administrador global con doble identidad visible (D44); audit completa (D27).
- [ ] Capabilities declaradas en código; UI consulta endpoint de capabilities; servidor rechaza operaciones no autorizadas (D45).

### Datos

- [ ] D94 formalizado: **todas** las FKs intra-app tienen `FOREIGN KEY` en PostgreSQL.
- [ ] D95 resuelto: orden de migración NoConformidades → Condor verificado; `idNCAsociada` mantiene coherencia referencial mediate adaptador.
- [ ] Migración de datos validada con backfill contra staging + conteos contra el backend autoritativo (~305 filas; staging vacío o casi vacío según auditoría).
- [ ] `tbMapeoCampos` con las 183 filas cargado **antes** del go-live; ausencia causa comportamiento incorrecto.

### Testing y calidad

- [ ] Tests E2E que cubren los 4 tipos de Solicitud + ciclo de workflow completo.
- [ ] Tests de capabilities verifican rechazo de operaciones no autorizadas.
- [ ] Tests de fallo explícito: ausencia de `m_BackendSandboxURL` en pytest → error claro (Spec-008).
- [ ] Smoke contra staging antes de promover a producción.
- [ ] UAT por admins de aplicación + usuarios de negocio (D47-D48); resultado y excepciones documentadas (D49-D50).

---

## 6. Pendientes operacionales

Acciones manuales que el equipo debe ejecutar antes, durante o después de la migración.

### Antes de empezar

- [ ] **Auditar `GetPasswordDB` en `00_CONDOR`** y rotar la contraseña real del backend si `"dpddpd"` era la de producción.
- [ ] **Mover la contraseña a secret manager** (D9-D10). Actualizar manifests de despliegue y dotenv si aplica.
- [ ] **Auditar git history de `00_CONDOR`** por commits que contengan la cadena `"dpddpd"`; documentar si fue commiteada. Si sí, **rotación obligatoria** + nota de incidente.
- [ ] **Respaldar `condor_datos.accdb` de producción** antes de la migración de datos (snapshot inmutable).
- [ ] **Verificar `TbConfiguracionBackends`** en `00_main` (no listado en staging) y consolidar config del módulo hexagonal.

### Durante la migración

- [ ] Validar que `tb_mapeo_campos` carga las **183 filas** antes del go-live; ausencia = comportamiento incorrecto del merge de plantillas.
- [ ] Validar FKs intra-app (D94) con pruebas de integridad; cualquier violación debe corregirse en origen antes de promover.
- [ ] Coordinar migración con NoConformidades (D95 — orden estricto).
- [ ] Probar cache safety (Spec-008) en el adaptador de testing equivalente.

### Después del go-live

- [ ] Confirmar que los logs llegan a Sentry/OTel antes de retirar las tablas VBA.
- [ ] Confirmar que la papelera de adjuntos (D19) funciona end-to-end.
- [ ] Verificar que la impersonación (D44) solo la inicia el administrador global.
- [ ] Cerrar el lower bound del audit (obs #24085) cuando se complete `00_main`: ¿alguna columna de `tbTransiciones`/`tbValidacionRevision`/`tbHistorialRechazos` con `Text(2)` que deba normalizarse a `BOOLEAN` (D102)?

---

## 7. Tickets derivables (preview — NO crear issues todavía)

> Lista de issues que nacerían de esta épica. **NO** se crean todavía; se trata de un preview para alinear con el equipo y empezar el desglose SDD.

### Núcleo de funcionalidad

- [ ] **TK-CONDOR-1**: Implementar CRUD Solicitud con 4 tipos (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`) en backend web — clases Pydantic + use cases + adaptadores de persistencia + UI HTMX por tipo.
- [ ] **TK-CONDOR-2**: Seed de `tbEstados` + `tbMapeoCampos` + `tbTransiciones` (9 + 183 + TBD filas) en migración inicial — scripts Alembic idempotentes.
- [ ] **TK-CONDOR-3**: Implementar workflow declarativo desde `tbTransiciones` — motor server-side que lee la tabla en runtime + endpoint "transiciones disponibles" + validación de capabilities (D45-D46).
- [ ] **TK-CONDOR-4**: Historia de Rechazo (`tbRechazos` + `tbHistorialRechazos`) en PostgreSQL con CRUD completo — tabla de NEGOCIO per obs #24086, no REPLACE.
- [ ] **TK-CONDOR-5**: Validación de calidad (`tbValidacionRevision`) con preservación de `HashDatos` para integridad de la revisión.
- [ ] **TK-CONDOR-6**: Adjuntos con referencia externa a object storage S3-compatible (D16) — metadatos en `tbAdjuntos`, contenido en S3, papelera 30 días (D19).
- [ ] **TK-CONDOR-7**: Vinculaciones con Lanzadera/Expedientes y con NoConformidades vía adaptadores unificados (D9-D10) — resolver `idExpediente` e `idNCAsociada` como referencias conceptuales mediate servicios remotos.

### Forma y plataforma

- [ ] **TK-CONDOR-8**: Snapshots web equivalentes al Edge WebView legacy — endpoints server-side de snapshot con autorización (D27, D45) + ETag (D72); eliminar `SnapshotServicio`/`WebVisorCacheServicio` cliente.
- [ ] **TK-CONDOR-9**: Sustituir 3 tablas de log VBA por Sentry/OTel — traducir llamadas VBA a eventos web con correlación (D27); verificar ausencia de tablas en PostgreSQL.
- [ ] **TK-CONDOR-10**: Mapeo de plantillas (`tbMapeoCampos`) — decisión con negocio (Word / Jinja2 / WeasyPrint) + UI admin + endpoint admin CRUD.
- [ ] **TK-CONDOR-11**: Capacidad de búsqueda y filtros avanzados (`FiltrosSolicitud`) — queries parametrizadas + UI HTMX con refresh manual (D69).
- [ ] **TK-CONDOR-12**: Notificaciones transaccionales (`NotificacionServicio`) — traducir a servicio unificado de notificaciones (D11, D12) canal email v1.
- [ ] **TK-CONDOR-13**: Suministradores + documentos anexos — Suministrador como entidad vinculada a `ExpedienteSuministrador` cross-app (F7); `DocumentoServicio` como servicio de documentos anexos.

### Seguridad y cumplimiento

- [ ] **TK-CONDOR-14**: Rotar `"dpddpd"` + mover a secret manager + audit git history (cierre D93 / D104) — sin este ticket, la épica no puede avanzar.
- [ ] **TK-CONDOR-15**: Formalizar todas las FKs intra-app en PostgreSQL (D94) — script Alembic con `FOREIGN KEY` para `tbSolicitudes.idEstadoInterno → tbEstados`, `tbTransiciones.idEstadoOrigen/Destino → tbEstados`, FKs a `tbSolicitudes` desde `tbRechazos`/`tbAdjuntos`/`tbValidacionRevision`/`tbHistorialRechazos`.
- [ ] **TK-CONDOR-16**: Suplantación solo por administrador global (D44) — doble identidad visible + audit completa (compatible con `suplantadoPor` que actualmente se loguea en `tbLogCambios`/`tbLogErrores`).

### Testing y calidad

- [ ] **TK-CONDOR-17**: Tests E2E para los 4 tipos de Solicitud + ciclo de workflow completo — pytest + httpx + Playwright para flujos críticos de UI (D68).
- [ ] **TK-CONDOR-18**: Adaptador de testing equivalente a `m_TestingMode`/`m_BackendSandboxURL` — puerto de testing con cache safety (Spec-008) + fail-fast "TESTS BLOCKED" si no hay sandbox.
- [ ] **TK-CONDOR-19**: Tests de capabilities — verificación que el servidor rechaza operaciones no autorizadas aunque la UI las muestre (D45).
- [ ] **TK-CONDOR-20**: Migración de datos + smoke contra staging antes de promover a producción — backfill validado con conteos contra `condor_datos.accdb` autoritativo (~305 filas).

---

## Anexo · Tabla de decisiones referenciadas

| Decisión | Tema | Estado | Aplica a Condor |
|---|---|---|---|
| D8 | Hexagonal global | APROBADO | Toda la épica |
| D9-D10 | Autenticación via adaptador | APROBADO | Identidad/permisos (D-Identidad) |
| D11-D13 | Notificación unificada + cola email | APROBADO | `NotificacionServicio` → email v1 |
| D14 | PostgreSQL compartido por esquemas | APROBADO | Persistencia del módulo |
| D16-D20 | Adjuntos object storage + papelera 30 días | APROBADO | `tbAdjuntos` (F6) |
| D27-D29 | Logs estructurados + retención | APROBADO | Sustituye 3 tablas de log |
| D36-D41 | Credenciales, lockout, caducidad | APROBADO | Integración con Lanzadera (no del módulo) |
| D42-D44 | Autorización + suplantación | APROBADO | Capa de auth + TK-CONDOR-16 |
| D45-D46 | Capabilities declaradas | APROBADO | Validación de `rolRequerido` |
| D66-D67 | Stack backend + frontend | APROBADO | Toda la épica |
| D68 | Monolito modular | APROBADO | Estructura del módulo |
| D78-D82 | Versionado + trunk-based + Expand and Contract | APROBADO | Releases y migraciones |
| D86-D87 | Forma hexagonal + tests VBA | APROBADO | Referencia de mapeo del puerto |
| D92 | Datos personales (HPS) — política | APROBADO | No bloqueante para Condor, revisión preventiva |
| D93 | Password hardcoded fallback | PROPUESTO (cierre en esta épica) | TK-CONDOR-14 |
| D94 | FKs conceptuales intra-app | PROPUESTO (formalización) | TK-CONDOR-15 |
| D95 | Vinculación NoConformidades | PROPUESTO | TK-CONDOR-7 |
| D96 | Workflow declarativo `tbTransiciones` | PROPUESTO | TK-CONDOR-3 |
| D97 | `tbMapeoCampos` 183 filas | PROPUESTO | TK-CONDOR-10 |
| D102 | Booleanos `Text(2)` cross-cutting | PROPUESTO | H6 — no bloqueante en Condor (ya son YesNo) |
| D104 | Secret manager cross-cutting | PROPUESTO | H7 — se cierra con TK-CONDOR-14 |

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| `docs/03-aplicaciones/condor/capabilities.md` | Inventario de features F1-F8, ViewModels/Servicios/Repositorios, edge WebView |
| `docs/03-aplicaciones/condor/data-model.md` | 15 schemas, volúmenes, convenciones, hallazgos D93-D97 ampliados |
| `docs/03-aplicaciones/condor/forms.md` | Navegación, call paths, patrones MVVM, formularios críticos |
| `docs/03-aplicaciones/condor/integrations-automation.md` | Cross-app, Edge WebView, testing sandbox, flags de operación |
| `docs/03-aplicaciones/condor/migration-matrix.md` | D93-D97 detallados + decisión D-new logs web-native |
| `docs/03-aplicaciones/condor/security-rules.md` | Autorización, D93-D95, riesgos de privacidad, XSS de WebView |
| `docs/03-aplicaciones/condor/README.md` | Estado del lote, hallazgos críticos, checklist |
| `docs/09-arquitectura-objetivo-y-principios.md` | Decisiones D5-D82 referenciadas |
| `docs/08-decisiones-y-preguntas-abiertas.md` | Decisiones D83-D97 cross-cutting con detalle |
| engram obs #24085 | Audit de uso de las 15 tablas |
| engram obs #24086 | Corrección: `tbHistorialRechazos` es de NEGOCIO |
| engram topic_key `condor/log-strategy-2026-08-05` | Decisión B — logs web-native |

---

## Checklist del documento

- [x] Estado del descubrimiento sincronizado con el audit obs #24085.
- [x] Hallazgos D93-D97 con anchor links.
- [x] Decisión B de logs (engram) citada textualmente con topic_key.
- [x] Corrección obs #24086 sobre `tbHistorialRechazos` aplicada.
- [x] Criterios de aceptación verificables y agrupados por dimensión.
- [x] Tickets cubren las 8 features + seguridad + testing (TK-CONDOR-1..20).
- [x] Tabla de decisiones referenciadas (D5-D104).
- [x] Tabla de fuentes.
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.

## Siguiente paso

Revisión con el equipo. Una vez validada, abrir SDD (`sdd-propose` + `sdd-spec` + `sdd-design` + `sdd-tasks`) para arrancar la implementación por ticket, comenzando por **TK-CONDOR-14** (rotación de contraseña) antes que cualquier otro cambio de código.
