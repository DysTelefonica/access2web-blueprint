# Épica — HPS (migración a web)

> **Estado:** DRAFT — pendiente revisión.
> **Versión:** v0.1 (2026-08-06).
> **Autor:** placeholder.
> **Lote de discovery:** 4 (junto al resto de las 7 apps restantes).
> **Cross-refs:** `docs/03-aplicaciones/hps/{capabilities,data-model,forms,integrations-automation,migration-matrix,security-rules,README}.md` · engram obs #24084 (consolidado de descubrimiento) · D92 dispositions específicas de HPS.

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_HPS` · frontend `HPS.accdb` (30 MB) + backend `HPST.accdb` (7 MB, autoritativo en `C:\00repos\datos\`) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (22 tablas, 16+ features de negocio, 29 clases, 30 módulos, 4 fuentes de anexos) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, vía `getdbLanzadera` o tablas compartidas); Expedientes (FK conceptual `IDExpediente`); HPS_Solicitudes (FK conceptual `IDSolicitud` — **app independiente** ID 22, D83); SICA (sistema externo vía `TbUsuariosSICA`); AGEDYS presumible (vía código compartido) |
| **Riesgo dominante** | Seguridad D92 — caché local en frontend `HPS.accdb` con **344 filas de PII** (DNI, nombres, fechas de nacimiento, correos, observaciones) en `TbDatosLocal` |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ **committed** (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`): **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE** de 22 tablas backend. Detalle en § 1.4 y § 2.1; hallazgo D92-HPS-AUDIT en § 3. Limitación declarada: lower bound del codegraph no incluye macros embebidas ni queries no exportadas; gap extractor en patrón `m_SQL='TableName'`. |

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

**16+ features de negocio** con paridad funcional en la nueva plataforma:

| # | Feature | Respaldo en docs |
|---|---|---|
| F1 | Arranque e identidad (`EVE` + `getUsuario`) | `capabilities.md` · `forms.md` · `data-model.md` § IDAplicacion |
| F2 | Configuración de backend (`TbConfiguracionBackends`) → config del módulo hexagonal | `capabilities.md` · `integrations-automation.md` § Configuración y flags |
| F3 | CRUD Usuario HPS (`TbUsuarios`, 345 filas) — alta, baja, renovación, modificación de datos personales, cálculo de `APuntoDeCaducar`/`Caducado`/`Solicitado`/`PendienteRenovacion` | `capabilities.md` · `forms.md` · `data-model.md` |
| F4 | Histórico de usuarios (`TbUsuariosHistoricos`, 242 filas) con `TbUsuariosHistoricosLocal` (235 filas en caché) | `capabilities.md` · `data-model.md` · `security-rules.md` § Histórico con anexos |
| F5 | HPS propiamente (`TbHPS`, 1280 filas) + equivalencias (`TbHPSEquivalencia`) + grado (`TbHPSGrado`) | `capabilities.md` · `data-model.md` · `forms.md` |
| F6 | Motivo HPS (`TbMotivoHPS`) — motivos de alta/baja | `capabilities.md` · `migration-matrix.md` |
| F7 | Cursos obligatorios (`TbAuxCursos`) | `capabilities.md` · `migration-matrix.md` |
| F8 | Observaciones (`TbObservaciones`, 330 filas) + históricas (`TbObservacionesHistoricas`) | `capabilities.md` · `forms.md` · `migration-matrix.md` |
| F9 | Anexos 4 fuentes — HPS / Histórico / SICA / Usuario (`TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos`) | `capabilities.md` · `forms.md` · `migration-matrix.md` |
| F10 | Integración SICA externa (`TbUsuariosSICA`) — referencia externa, sin propiedad | `capabilities.md` · `integrations-automation.md` |
| F11 | Indicadores en tiempo real con kill switch atómico (`Indicador.cls`, `clsIndicadoresBus.cls`, `Test_RealTimeIndicatorCoherence.bas`) | `capabilities.md` · `forms.md` · `security-rules.md` § Indicadores |
| F12 | Transacciones de selección de anexos (`AnexoSelectionTransactionCoordinator`) + lifecycle de usuarios (`UsuarioLifecycleTransactionCoordinator`) | `capabilities.md` · `forms.md` · `security-rules.md` § Transaccionalidad |
| F13 | Consultas pre-armadas — 8 formularios `Form_FormInicial03Consultas<PorCampo>` (Contratistas, Datos, EmpresaUsuario, Estado, Grado, JuridicaTramitacion, MotivoHPS, Nombre, ProyectoAsignado) | `capabilities.md` · `forms.md` § Navegación principal |
| F14 | Consultas SQL externas (`TbConsultas` + `Consulta.cls`) — query builder parametrizable | `capabilities.md` · `migration-matrix.md` |
| F15 | Sincronización histórico ↔ anexos (`Mod_Sincronizacion_Historico.bas` + `Test_HistoricoAdjuntosTransactionWrapper.bas`) | `capabilities.md` · `security-rules.md` § Sincronización |
| F16 | Catálogo de jurídicas (`TbJuridicasContratacion`) | `capabilities.md` · `migration-matrix.md` |

**5 catálogos config seed-only** (datos de runtime, sin CRUD de usuario):

| Catálogo | Volumen | Notas |
|---|---|---|
| `motivos_hps` (era `TbMotivoHPS`) | TBD | Motivos de alta/baja HPS |
| `cursos` (era `TbAuxCursos`) | TBD | Cursos obligatorios |
| `hps_equivalencia` (era `TbHPSEquivalencia`) | TBD | Equivalencias entre cursos HPS |
| `hps_grado` (era `TbHPSGrado`) | TBD | Grados del HPS |
| `juridicas_contratacion` (era `TbJuridicasContratacion`) | TBD | Jurídicas de contratación |

**Volumen real autoritativo** (backend `HPST.accdb` en `C:\00repos\datos\`):

| Tabla | Filas |
|---|---|
| `TbUsuarios` | 345 (usuarios activos) |
| `TbUsuariosHistoricos` | 242 (usuarios históricos) |
| `TbHPS` | 1280 (relación usuario-curso) |
| `TbObservaciones` | 330 (observaciones) |
| **Total principales** | **≈2197** (más catálogos y anexos sin auditar) |

**Read-heavy**: `getdb()` con **107 callers** (el más bajo de las 8 aplicaciones). La mayoría de las operaciones son consultas sobre el catálogo de usuarios; escrituras limitadas a actualizaciones de HPS, observaciones y carga inicial.

**Caché selectivo maduro** (confirma D91 cross-app): 6 módulos (`cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas`) controlados vía `TbConfiguracion.CacheHabilitada` (kill switch). Patrón idéntico al de NoConformidades — se preserva como **referencia del puerto de caché** (D70-D71).

**Tests VBA como referencia** (D87): 9 archivos `Test_*.bas` (`Test_AnexoSelectionTransaction`, `Test_CacheConsistencyAudit`, `Test_HistoricoAdjuntosTransactionWrapper`, `Test_HPSConfig`, `Test_HPSEntorno`, `Test_LocalReadAccessAuthorization`, `Test_PerAnexoMove`, `Test_RealTimeIndicatorCoherence`, `Test_StartupCacheInitialization`). Cobertura significativa; **test crítico a portar**: `Test_LocalReadAccessAuthorization.bas` valida que el acceso local a datos sigue reglas de autorización.

**TDD maduro**: 5 archivos `clsTestDouble*` (`clsTestDoubleForm`, `clsTestDoubleHpsEditListener`, `clsTestDoubleHpsEditPublisher`, `clsTestDoubleIndicador`, `clsTestDoubleIndicadorConsumer`) + 6 fixtures (`DummyBig`, `DummyBig2`, `DummyHatw`, `DummyHatw2`, `DummySmall`, `DummySpecial`, `Módulo1`). Indica disciplina TDD madura en el legacy.

**Vinculaciones externas**: la identidad se resuelve vía **adaptador unificado** (D9-D10); las FKs conceptuales a Expedientes/HPS_Solicitudes se mantienen como referencia conceptual mediate adaptadores (D86/D87). **HPS_Solicitudes es app independiente** (ID 22 vs ID 17 de HPS) — no comparten tablas (D83).

### 1.2 Fuera de scope (REPLACE)

> Las **12 tablas locales del frontend** NO migran como caché en cliente. Se reemplazan por stack de caché server-side detrás del puerto (D70-D71). Sus equivalentes de negocio viven en el backend autoritativo; el frontend `.accdb` deja de existir como runtime.

**⚠️ CRÍTICO — caché local con PII** (D92):

- `TbDatosLocal` (**344 filas**) — contiene TODOS los datos personales de usuarios HPS activos: DNI, Nombre, Apellidos, Teléfono, Correo_e, F_Nacimiento, LugarNacimiento, Observaciones + estado HPS desnormalizado por organismo (NAC, OTAN, ESA, UE).
- `TbUsuariosHistoricosLocal` (235 filas) — datos personales del histórico de usuarios (DNI, nombres, fechas, correo).

**Caché auxiliar** (sensibilidad media-baja):

- `TbCursosLocal` (caché de cursos)
- `TbSuministradoresLocal` (caché de suministradores)
- `TbDatosLocalParaIndicadores` (caché para indicadores)
- `TbUsuariosSICALocalParaIndicadores` (caché SICA para indicadores)
- `TbConfiguracionHPS` (clave-valor local: Clave, Valor, Activo, FechaModificacion, UsuarioModificacion)

**Metadatos locales** (config, no datos):

- `TbVinculosTablas`, `tblInfo`, `tblSettings` — metadatos locales del frontend → migrar como **config del módulo** (no tablas, sino variables de entorno o config centralizada).

**Descartar**:

- `TbUsuariosSICALocal` (0 filas, no se usa) — descartar.

**Excluido del REPLACE**: las **22 tablas del backend** (`HPST.accdb`) migran como tablas PostgreSQL con sus claves y datos preservados. Solo `TbConsultas` requiere decisión sobre query builder parametrizable.

**Persistencia**: engram obs #24084 (consolidado); D92 dispositions específicas de HPS en [`migration-matrix.md`](migration-matrix.md#d92--disposiciones-específicas-de-hps).

### 1.3 Fuera de scope (no documentado)

Si aparece algo que no está en los 7 docs de HPS, se marca como **no documentado** y se acumula en pendientes de discovery para iteración posterior. Ejemplos conocidos:

- Volúmenes reales de producción (no staging): `HPST.accdb` en `C:\00repos\datos\` no fue inspeccionado en esta pasada.
- Conteos de filas de los catálogos seed-only (`TbAuxCursos`, `TbHPSEquivalencia`, `TbHPSGrado`, `TbMotivoHPS`, `TbJuridicasContratacion`) — solo `TbUsuarios` y `TbObservaciones` principales auditadas.
- Auditoría de uso `codegraph-vba` específica para HPS (committed — engram obs #24097, topic_key `hps/usage-audit-2026-08-06`): 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE. La auditoría equivalente de Condor sigue en obs #24085 como referencia del patrón.
- `TbConsultas` contenido (queries SQL pre-armadas) — sensible a SQL injection si no se parametriza; decisión pendiente.
- Inventario completo de queries exportadas y macros embebidas — codegraph no captura macros embebidas ni QueryDefs no exportados.
- Plantillas de correo en `Correo.cls` — pueden contener datos sensibles (D92); revisión pendiente antes de portar.

### 1.4 Fuera de scope (no-migrate — zombies confirmados)

> 9 tablas de las 22 backend son zombies confirmados (audit obs #24097). No migran como features a la nueva plataforma. Tratamiento: ver § 6 Pendientes operacionales.

**4× copias legacy "Copia de..."** (herencia de mass-change pre-consolidación; retención indefinida, sin endpoints, sin UI):

- `Copia de TbExpedienteLugares`
- `Copia de TbExpedientes`
- `Copia de TbUsuarios`
- `Copia de TbUsuariosEntidades`

**Sentinel de errores**:

- `Errores de pegado` — sentinel de errores en pegado masivo. Transformar en log estructurado canónico (D27) o preservar con retención limitada.

**3 catálogos sin uso** (confirmar con negocio antes de descartar formalmente; ver § 6):

- `TbHPSEquivalencia` — catálogo sin uso (0 callers en audit).
- `TbAuxCursos` — catálogo sin uso (0 callers en audit).
- `TbJuridicasContratacion` — catálogo sin uso (0 callers en audit).

**Solapamiento con tabla activa**:

- `TbUsuarioAnexos` — 0 callers; sospecha de solape con `TbAnexosUsuariosHPS`. Confirmar con negocio antes de descartar.

---

## 2. Estado del descubrimiento

### 2.1 Inventario

| Categoría | Resultado |
|---|---|
| **Tablas totales (backend)** | **22** en `HPST.accdb` (autoritativo en `C:\00repos\datos\`) |
| **Tablas totales (frontend local)** | **12** en `HPS.accdb` (caché + metadatos) — **NO migran como tales** |
| **Filas totales (backend, auditadas)** | **≈2197** (4 tablas principales: 345 + 242 + 1280 + 330) |
| **Filas totales (frontend local)** | **≈579** (344 + 235 + catálogos auxiliares) — **NO migran como tales** |
| **Schemas documentados** | **22/22** tablas backend en `data-model.md` (commits HPS del lote 4, ver obs #24084) — **esquema detallado completo de `TbUsuarios` (27 columnas) committed**; resto en nivel inventario |
| **Uso (audit)** | ✅ **committed** (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`): **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE** de 22 tablas backend. Desglose en § 1.4 (ZOMBIE) y § 2.1.1 (ACTIVE / UNCLEAR). Limitación: lower bound del codegraph no incluye macros embebidas ni queries no exportadas; gap extractor en patrón `m_SQL='TableName'` (afecta a `TbConsultas`, marcada UNCLEAR). |
| **Subcategorías** | 6 dominio (usuarios, histórico, HPS, observaciones, anexos, indicadores) · 5 catálogos · 4 legacy copies · 1 sentinel · 4 anexos cross-source · 2 config |
| **FKs físicas (backend)** | **6** desde `TbUsuarios` / `TbUsuariosHistoricos` / `TbUsuariosSICA` (3 con PK-to-PK genérico a corregir, ver D94) |
| **FKs conceptuales (sin constraint)** | **5** — `IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud` |
| **Clases inventariadas** | 29 (16 dominio, 5 Test Doubles, 2 coordinadores transaccionales, 3 anexos, 2 compartidas con Lanzadera, 2 indicadores) |
| **Módulos** | 30 (4 bootstrap/factory/DAO, 6 caché, 1 indicadores, 1 configuración, 1 filtros, 1 JSON, 7 fixtures, 9 tests VBA) |
| **Forms** | ~30 archivos `Form_*.cls` cada uno con su `.form.txt` compañero |
| **Tests VBA** | 9 archivos `Test_*.bas` (cobertura significativa, confirma D87) |
| **Duplicación frontend/backend** | `TbHPS` y `TbUsuariosHistoricos` existen en ambos `.accdb` (Dysflow emite `ACCESS_TABLE_AMBIGUOUS` al consultarlas sin `target`) |

### 2.1.1 Veredicto del schema usage audit (obs #24097)

> Audit ejecutado 2026-08-06 vía `codegraph-vba` sobre `00_HPS` READ-ONLY. Repo `00_HPS` no modificado. Veredicto sobre las **22 tablas backend** de `HPST.accdb`: **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE**.

**12 ACTIVE** (callers encontrados en `codegraph-vba`):

| Tabla | Callers | Notas |
|---|---|---|
| `TbUsuarios` | 5 | PII D92, 27 columnas, FKs conceptuales; columna central del dominio usuarios HPS. |
| `TbUsuariosHistoricos` | 5 | Histórico; FK PK-to-PK a corregir (D94). |
| `TbUsuariosSICA` | 3 | Integración externa SICA (sistema fuera de scope). |
| `TbUsuariosEntidades` | 5 | Relación usuario ↔ entidad N:N; consumida por caché y sync histórico. |
| `TbHPS` | 5 | 1280 filas; duplicación frontend/backend (D92). |
| `TbHPSGrado` | 1 | Catálogo lightweight — mantener. |
| `TbMotivoHPS` | 1 | Catálogo lightweight — mantener. |
| `TbObservaciones` | 5 | 330 filas; FK PK-to-PK a corregir (D94). |
| `TbObservacionesHistoricas` | 5 | Histórico; FK PK-to-PK a corregir (D94). |
| `TbAnexosUsuariosHPS` | 5 | Anexos de usuario HPS — external refs D16. |
| `TbAnexosUsuariosSICA` | 2 | Anexos SICA — referencia externa. |
| `TbAnexosUsuariosHistoricos` | 5 | Anexos histórico — external refs D16. |

**1 UNCLEAR con gap del extractor**:

| Tabla | Veredicto | Gap |
|---|---|---|
| `TbConsultas` | UNCLEAR | `codegraph-vba` no emite arista `vba-sql-table` para el patrón `m_SQL='TableName'`. Class-chain detectado via `Constructor.getConsultas:3192` → `Entorno.ColConsultas` pero sin blast a la tabla. **Gap del extractor a reportar al maintainer**; revisión humana pendiente antes de clasificar como ACTIVE o ZOMBIE. |

**9 ZOMBIE** (0 callers en audit; detalle por subcategoría en § 1.4):

| # | Tabla | Subcategoría |
|---|---|---|
| 1 | `Copia de TbExpedienteLugares` | Legacy copy pre-mass-change. |
| 2 | `Copia de TbExpedientes` | Legacy copy pre-mass-change. |
| 3 | `Copia de TbUsuarios` | Legacy copy pre-mass-change. |
| 4 | `Copia de TbUsuariosEntidades` | Legacy copy pre-mass-change. |
| 5 | `Errores de pegado` | Sentinel de errores en pegado masivo. |
| 6 | `TbHPSEquivalencia` | Catálogo sin uso (0 callers). |
| 7 | `TbAuxCursos` | Catálogo sin uso (0 callers). |
| 8 | `TbJuridicasContratacion` | Catálogo sin uso (0 callers). |
| 9 | `TbUsuarioAnexos` | Solapamiento presunto con `TbAnexosUsuariosHPS` (0 callers). |

**Limitaciones declaradas del audit**:

- Macros embebidos en `HPST.accdb` NO exportados → uso real puede ser mayor al lower bound.
- `codegraph_explore` limita el blast radius visible a 5 entries por símbolo.
- Forms sin `RecordSource`/`RowSource` exportados: app carga vía DAO + clases.
- `TbDatosLocal` (344 filas PII) está en el FRONTEND (`HPS.accdb`), fuera del scope del audit backend.

**Referencia canónica**: engram obs #24097, topic_key `hps/usage-audit-2026-08-06`, skill `codegraph-usage` + `vba-sql-impact`, 100% vía codegraph MCP (cero grep sobre `.bas`/`.cls`).

### 2.2 Hallazgos estructurales del inventario

**4 tablas "Copia de..."**: patrón legacy de copia antes de cambios masivos. Probablemente obsoletas. Requieren disposición explícita en migración:

- `Copia de TbExpedientes`
- `Copia de TbExpedienteLugares`
- `Copia de TbUsuarios`
- `Copia de TbUsuariosEntidades`

**Sentinel `Errores de pegado`**: tabla para capturar errores en operaciones de pegado masivo (imports batch). Migración debe decidir si preservar (con política de retención) o transformar en un log estructurado canónico (D27).

**Inconsistencias detectadas en FKs** (data-model.md § Relaciones físicas reales):

1. `TbUsuariosHistoricos.ID → TbObservacionesHistoricas.ID` (PK-to-PK genérico, debería ser por `IDUsuarioHistorico`).
2. `TbUsuarios.ID → TbObservaciones.ID` (PK-to-PK genérico, debería ser por `IDUsuario`).
3. NO hay FK física entre `TbUsuarios` y `TbExpedientes` pese a `IDExpediente`.
4. NO hay FK física entre `TbUsuarios` y `TbEmpresas` pese a `IDEmpresaUsuario`/`IDEmpresaHPS`.
5. NO hay FK física entre `TbUsuarios` y `TbJuridicasContratacion` pese a `IDJuridicaContrato`.
6. NO hay FK física entre `TbUsuarios` y `TbSolicitudes` pese a `IDSolicitud`.

**Booleanos como Text(2)** (D102): 3 columnas en `TbUsuarios` con `VARCHAR(2) NULL` y valores 'Sí'/'No' — `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`. Inconsistencia detectada también en NoConformidades; regla general cross-cutting (D102).

### 2.3 Herramientas usadas

- **Dysflow MCP** (read-only) sobre `HPS.accdb` + `HPST.accdb`: inventario real, conteo de filas, FKs, esquema de `TbUsuarios`.
- **codegraph-vba** sobre `C:\00repos\codigo\00_HPS\00_main\.codegraph-vba`: extracción de call paths dinámicos (31 símbolos en 1 archivo para la query inicial), blast radius de símbolos clave (`EVE`, `getdb`, `getUsuario`, `Indicador`, `AnexoSelectionTransactionCoordinator`, `UsuarioLifecycleTransactionCoordinator`).
- **No se ejecutó `codegraph_explore` de audit de uso** específico para HPS (análogo al obs #24085 de Condor). Pendiente antes de fase SDD.

### 2.4 Pendientes menores de discovery

| Pendiente | Impacto | Iteración |
|---|---|---|
| Audit de uso específico para HPS (codegraph-vba + cruzar con `src/`) | Determinar ACTIVE/ZOMBIE/UNCLEAR de las 22 tablas backend | ✅ **cerrado** (engram obs #24097: 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE). Revisión humana pendiente para `TbConsultas` (gap extractor `m_SQL='TableName'`) y `TbUsuarioAnexos` (solapamiento presunto con `TbAnexosUsuariosHPS`). |
| Conteos reales de catálogos seed-only (`TbAuxCursos`, `TbHPSEquivalencia`, `TbHPSGrado`, `TbMotivoHPS`, `TbJuridicasContratacion`) | Ajustar criterios de aceptabilidad de migración de catálogos | Cuando se acceda al backend autoritativo |
| Contenido de `TbConsultas` (queries SQL pre-armadas) | Evaluar SQL injection; decidir query builder vs deprecate | Próxima iteración |
| Cadena de uso de las 4 "Copia de..." por saved queries / macros | Decidir archivo vs dominio | Inspección manual si negocio requiere |
| Plantillas de correo en `Correo.cls` con datos sensibles (D92) | Política de privacidad antes de portar | Antes de TK-HPS-14 (D92 cleanup) |
| Auditoría operativa del frontend `HPS.accdb` en `00_HPS` (no desde aquí): `git log --all -- HPS.accdb`, añadir `*.accdb` al `.gitignore` | Riesgo de PII en historial git | TK-HPS-15 (operativo en repo `00_HPS`) |

---

## 3. Hallazgos críticos

Cada hallazgo: descripción corta + impacto + acción + referencia documental.

<a id="hallazgo-D92-cache"></a>
### H1 · D92 — Caché local en frontend con 344 filas de PII  ⚠️⚠️⚠️ CRÍTICO

- **Descripción**: el frontend `HPS.accdb` (30 MB) tiene **12 tablas locales** que funcionan como caché sincronizado con el backend. La más sensible es `TbDatosLocal` con **344 filas** que contiene TODOS los datos personales de usuarios HPS activos (DNI, Nombre, Apellidos, Teléfono, Correo_e, F_Nacimiento, LugarNacimiento, Observaciones) + estado HPS desnormalizado por organismo (NAC, OTAN, ESA, UE). `TbUsuariosHistoricosLocal` (235 filas) repite el patrón para histórico.
- **Impacto**:
  1. El `.gitignore` del repo `00_HPS` NO excluye `*.accdb` (solo `*.accde`, `*.mdb`, `*.mde`, `HPST.accdb`). Si `HPS.accdb` se versiona, **344 usuarios con datos personales quedan en el historial de git**.
  2. El frontend se mueve con los datos: copia a otra máquina, backup no cifrado, robo de portátil = exposición de PII.
  3. Sin política de retención: datos locales pueden persistir después de baja del usuario.
  4. Caché desactualizada: frontend puede mostrar datos obsoletos sin que el usuario lo sepa.
- **Acción** (combinada, operativa + migración):
  - **Operativo en `00_HPS`** (no desde aquí): añadir `*.accdb` al `.gitignore`; verificar `git log --all -- HPS.accdb`; si está versionado, **rotar repo + git-filter-repo**.
  - **Migración**: las 12 tablas locales **NO se reproducen como caché en cliente**. Se traducen a **vistas materializadas o queries server-side** con caché detrás del puerto (D70-D71). Detalle completo en D92 dispositions de [`migration-matrix.md`](migration-matrix.md#4--hallazgo-crítico-caché-local-en-frontend-con-datos-personales).
  - Considerar cifrar el frontend o extraer datos personales a un esquema separado que pueda rotarse independientemente.
- **Detalle completo**: [`migration-matrix.md § D92`](migration-matrix.md#d92--disposiciones-específicas-de-hps) · [`security-rules.md § D92`](security-rules.md#hallazgo-crítico-de-seguridad-caché-local-en-frontend-con-datos-personales) · engram obs #24084 (consolidado).

<a id="hallazgo-D102-hps"></a>
### H2 · D102 — Booleanos como `Text(2)` 'Sí/No' en HPS (cross-cutting)

- **Descripción**: 3 columnas en `TbUsuarios` son `VARCHAR(2) NULL` con valores 'Sí'/'No' — `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`. Inconsistencia detectada también en NoConformidades; regla general cross-cutting (D102).
- **Impacto**: queries booleanas se complican; PostgreSQL debe estandarizar a `BOOLEAN`. Cualquier valor fuera de 'Sí'/'No' en origen requiere default explícito.
- **Acción**: regla de migración explícita — `'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`. Aplicar en HPS, Lanzadera, Expedientes, Gestion_Riesgos, NoConformidades (cross-cutting).
- **Detalle completo**: [`migration-matrix.md § D92.5`](migration-matrix.md#5--booleanos-como-text2) · [`data-model.md`](data-model.md#esquema-real-tbusuarios-27-columnas-tipos-reales).

<a id="hallazgo-D94-hps"></a>
### H3 · D94 — FKs conceptuales sin constraint (HPS)

- **Descripción**: `IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud` en `TbUsuarios` son FKs por convención de código, no por constraint. Data integrity gap en legacy. Las 2 FKs físicas `TbUsuariosHistoricos.ID → TbObservacionesHistoricas.ID` y `TbUsuarios.ID → TbObservaciones.ID` son PK-to-PK genéricas y deberían ser por `IDUsuarioHistorico`/`IDUsuario`.
- **Impacto**: integridad referencial no garantizada en legacy. En PostgreSQL hay que formalizar las FKs intra-app (D94) y mantener las FKs cross-app (`IDExpediente`, `IDSolicitud`) como referencia conceptual mediate adaptadores (D86/D87).
- **Acción**:
  - **Intra-app**: formalizar con `FOREIGN KEY` real — `TbObservaciones.IDUsuario → TbUsuarios.ID`, `TbObservacionesHistoricas.IDUsuarioHistorico → TbUsuariosHistoricos.ID`, las 4 FKs `TbUsuarios.ID → Tb*` (`TbAnexosUsuariosHPS`, `TbHPS`, `TbObservaciones`, `TbUsuarioAnexos`), las 2 FKs desde `TbUsuariosSICA`.
  - **Cross-app**: mantener como referencia conceptual; HPS_Solicitudes es app independiente (ID 22) y **no comparte tablas** (D83).
- **Detalle completo**: [`migration-matrix.md § D92.6`](migration-matrix.md#6--fks-conceptuales-sin-fk-física) · [`data-model.md` § Relaciones físicas reales](data-model.md#relaciones-físicas-reales-6-fk-entre-user-tables).

<a id="hallazgo-D44-hps"></a>
### H4 · D44 — Suplantación (cross-cutting, debe respetarse)

- **Descripción**: el patrón `suplantadoPor` se loguea en `tbLogCambios`/`tbLogErrores` (no aplica directo a HPS — no tiene logs propios — pero el patrón existe en otras apps del ecosistema). Suplantación restringida al administrador global con doble identidad visible y auditoría completa (D44).
- **Impacto**: HPS no introduce suplantación propia; pero debe respetar la política cross-cutting de plataforma al integrar con adaptadores de identidad.
- **Acción**: capabilities (D45-D46) verifican que solo el administrador global puede iniciar suplantación. Audit completa en logs estructurados (D27).
- **Detalle completo**: D44 en `docs/08-decisiones-y-preguntas-abiertas.md`.

### Hallazgos cross-cutting que aplican a HPS

<a id="hallazgo-D91-hps"></a>
### H5 · D91 — Caché selectivo maduro cross-app (confirmado en HPS)

- **Descripción**: HPS implementa el mismo patrón de caché selectivo maduro que NoConformidades — 6 módulos propios (`cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas`) coordinados vía `TbConfiguracion.CacheHabilitada` (kill switch atómico).
- **Impacto**: el patrón legacy **se preserva como referencia** del puerto de caché de la nueva plataforma (D70-D71). La nueva plataforma usa caché server-side (Redis o equivalente detrás del puerto), no caché en cliente.
- **Acción**: la migración de HPS aporta dos referencias valiosas — el patrón de kill switch (`Mod_StartupCacheInitialization.bas` + `CacheConsistencyAudit.bas`) y el patrón de transacciones atómicas sobre anexos (`AnexoSelectionTransactionCoordinator` + `Test_AnexoSelectionTransaction.bas`).
- **Detalle completo**: [`integrations-automation.md § Automatización`](integrations-automation.md#automatización-y-tareas).

<a id="hallazgo-D86-hps"></a>
### H6 · D86/D87 — Forma hexagonal + tests VBA como evidencia (cross-cutting)

- **Descripción**: HPS tiene 29 clases (dominio + Test Doubles + coordinadores transaccionales) + 9 archivos `Test_*.bas` con cobertura significativa. Indica forma hexagonal previa al blueprint moderno (D86): dominio en clases, helpers por dominio, transaccionalidad en coordinadores. Los tests VBA existentes son **evidencia de comportamiento** que se preserva como referencia para los nuevos tests pytest (D87).
- **Impacto**: ningún `Test_*.bas` se descarta sin trazabilidad. Test crítico a portar: `Test_LocalReadAccessAuthorization.bas`.
- **Acción**: cada test VBA tiene un equivalente pytest en la migración. Tests de coordinadores transaccionales (`Test_AnexoSelectionTransaction`, `Test_HistoricoAdjuntosTransactionWrapper`) son la base del puerto de transacciones de la nueva plataforma.
- **Detalle completo**: D86/D87 en `docs/08-decisiones-y-preguntas-abiertas.md`.

<a id="hallazgo-D103-cache"></a>
### H7 · Cache local del frontend — patrón que NO se reproduce (D92)

- **Descripción**: HPS es la **única aplicación del ecosistema con caché en el frontend** (no presente en Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades/Condor). El sistema de caché local completo (`TbDatosLocal` + 11 tablas auxiliares) es un patrón legacy que **NO debe reproducirse** en la nueva plataforma.
- **Impacto**: la nueva plataforma usa caché server-side detrás del puerto (D70-D71). La eliminación del caché frontend es una mejora de seguridad (PII no se mueve con el cliente) y de coherencia (no hay datos obsoletos).
- **Acción**: la regla "no caché en cliente" aplica a HPS como principio fundador. Migración traduce las 12 tablas a vistas materializadas o queries parametrizadas server-side.

<a id="hallazgo-D92-hps-audit"></a>
### H8 · D92-HPS-AUDIT — 9 tablas ZOMBIE confirmadas (audit obs #24097)

- **Descripción**: el schema usage audit ejecutado vía `codegraph-vba` el 2026-08-06 (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`) clasificó las 22 tablas backend de HPS como **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE**. Las 9 ZOMBIE (0 callers en `codegraph_explore`) son: 4× "Copia de..." (`Copia de TbExpedienteLugares`, `Copia de TbExpedientes`, `Copia de TbUsuarios`, `Copia de TbUsuariosEntidades`), el sentinel `Errores de pegado`, los 3 catálogos `TbHPSEquivalencia` / `TbAuxCursos` / `TbJuridicasContratacion`, y `TbUsuarioAnexos` (solapamiento presunto con `TbAnexosUsuariosHPS`). Las 4 copias "Copia de..." son herencia de un mass-change que no se consolidó en su día.
- **Impacto**: 9 de 22 tablas backend **no migran como features** a la nueva plataforma. Reducción significativa del scope de migración (≈41% del backend en superficie). Las 4 "Copia de..." arrastran origen cross-app (`TbExpedientes`/`TbExpedienteLugares` son de la app Expedientes — ver `migration-matrix.md` § 2); su retención indefinida en zona `legacy` del esquema PostgreSQL requiere coordinar con el equipo de Expedientes antes de consolidar el veredicto.
- **Acción**:
  - Cerrar el veredicto de las 9 ZOMBIE con **TK-HPS-27** (confirmación con Natalia / negocio: mantener histórico, transformar en log, o purgar — ver § 6 Pendientes operacionales).
  - Documentar formalmente la disposición en `migration-matrix.md` como categoría "no-migrate".
  - Resolver el UNCLEAR `TbConsultas` (gap del extractor `m_SQL='TableName'`) antes de fase SDD — pendiente de revisión humana.
  - Mantener el patrón de audit previo a migrar (obs #24082) como gate para próximas aplicaciones del blueprint.
- **Detalle completo**: engram obs #24097 (veredicto completo); § 1.4 (fuera de scope ZOMBIE); § 2.1.1 (veredicto); § 6 (pendiente operacional TK-HPS-27).

---

## 4. Decisiones aplicadas

> Las decisiones que ya están tomadas y que esta épica respeta sin复议ar.

<a id="dec-d92-cache-frontend"></a>
### D-CacheFrontend · NO migrar la caché frontend (D92)

- **Alcance**: las 12 tablas locales del frontend (`TbDatosLocal` + 11 auxiliares) **NO migran como caché en cliente**. Se reemplazan por stack de caché server-side detrás del puerto (D70-D71).
- **Motivación**:
  - Riesgo de PII: el frontend se mueve con los datos personales (D92).
  - Caché desactualizada: el frontend puede mostrar datos obsoletos.
  - Patrón no presente en las otras 7 apps: HPS es la única con caché en cliente.
- **Trade-off explícito**: el rendimiento legacy dependía de caché local. La nueva plataforma debe demostrar que PostgreSQL bien indexado + caché server-side selectiva (D70-D71) elimina la percepción de lentitud sin replicar el patrón legacy.
- **Persistencia**: engram obs #24084 (consolidado).
- **Ref**: [`migration-matrix.md § D92.4`](migration-matrix.md#4--hallazgo-crítico-caché-local-en-frontend-con-datos-personales).

### D-Bool · Booleanos `Text(2)` → `BOOLEAN` con regla explícita (D102)

- **Alcance**: las columnas `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion` se traducen a `BOOLEAN` en PostgreSQL con regla de migración explícita (`'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`).
- **Motivación**: regla general cross-cutting (D102); queries booleanas se complican con texto. Aplicar misma regla en Lanzadera, Expedientes, Gestion_Riesgos, NoConformidades.
- **Trade-off**: si en producción aparecen valores fuera de 'Sí'/'No' se loguean como anomalía y se decide caso a caso.

### D-FK-intra · FKs conceptuales intra-app → formales en PostgreSQL

- **Aplicación de D94 a HPS**: formalizar las 6 FKs físicas reales (corregir las 2 PK-to-PK genéricas por `IDUsuario`/`IDUsuarioHistorico`). Las 5 FKs conceptuales cross-app (`IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud`) se mantienen como referencia conceptual mediate adaptadores.
- **Caso especial `IDSolicitud`**: HPS_Solicitudes es app independiente (ID 22, D83); **no se comparte tabla**, la referencia es solo por convención. La migración de HPS_Solicitudes puede ocurrir en orden distinto sin acoplamiento.

### D-Identidad · Vía adaptador unificado (D9-D10)

- **Aplicación**: la identidad se resuelve por el adaptador unificado de autenticación, no por acceso directo a la BD de Lanzadera (`getdbLanzadera`). `IDAplicacion = "17"` (producción) / `"51"` (pruebas) queda como config del módulo.
- **Implicación**: el módulo HPS no conoce la BD de Lanzadera; el adaptador de identidad expone `getUsuarioConPermisos(usuario, aplicacion)` server-side.

### D-Backend-config · `TbConfiguracionBackends` → config del módulo hexagonal

- **Aplicación**: `BackendActivo`, `EnPruebas`, `IDAplicacion`, `PasswordBackend`, `CacheHabilitada` dejan de ser TempVars o flags de runtime; pasan a **config del módulo hexagonal** (variables de entorno del runner). `PasswordBackend` desaparece (D93, secret manager).
- **Detalle**: [`integrations-automation.md`](integrations-automation.md#configuración-y-flags-de-operación).

### D-Cache · Caché selectiva server-side detrás del puerto (D70-D71)

- **Aplicación**: el patrón maduro de HPS (6 módulos de caché + kill switch) se preserva como **referencia del puerto de caché**. La nueva plataforma usa caché server-side (Redis u opción detrás del puerto), no caché en cliente. Caché selectiva y justificada por medición, no proactiva (D70).
- **Detalle**: [`integrations-automation.md`](integrations-automation.md#automatización-y-tareas).

### D-PII · Política de datos personales (D92)

- **Alcance**: las columnas `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e`, `F_Nacimiento`, `LugarNacimiento` en `TbUsuarios` (y equivalentes en histórico) requieren política explícita:
  - **Preservar en PostgreSQL** sin transformaciones para mantener paridad funcional.
  - **Enmascarar en logs y observabilidad** (no loguear valores completos).
  - **Documentar en matriz de capabilities** quién puede ver cada campo (D45).
  - **Evaluar encriptación en reposo** como follow-up de seguridad (no bloqueante para la migración inicial).
- **Persistencia**: engram obs #24084; [`migration-matrix.md § D92.1`](migration-matrix.md#1--campos-personales-sensibles-en-tbusuarios).

### D-Testing · Tests VBA como referencia (D87)

- **Aplicación**: los 9 archivos `Test_*.bas` de HPS son **evidencia de comportamiento** que se preserva como referencia para los nuevos tests pytest. Cada test tiene un equivalente pytest en la migración; `Test_LocalReadAccessAuthorization.bas` es **test crítico a portar** (autorización local).
- **Detalle**: [`capabilities.md` § Reglas de conservación](capabilities.md#reglas-de-conservación).

### D-LegacyCopies · 4 "Copia de..." → zona legacy, deprecate

- **Aplicación**: las 4 tablas `Copia de Tb*` (legacy copy antes de cambios masivos) migran a una **zona `legacy` del esquema PostgreSQL** con retención indefinida. No se crean endpoints ni UI. Se marcan como deprecated en la documentación del módulo.

### D-Sentinel · `Errores de pegado` → log estructurado canónico (D27)

- **Aplicación**: el sentinel `Errores de pegado` se evalúa caso a caso:
  - Si la lógica de captura es **replicable** → migrar a **logs estructurados canónicos** (D27).
  - Si es **específica de HPS** → preservar como tabla de auditoría con retención limitada.

### D-Historico-Anexos · Sincronización atómica preservada

- **Aplicación**: `Mod_Sincronizacion_Historico.bas` + `Test_HistoricoAdjuntosTransactionWrapper.bas` mantienen coherencia atómica entre `TbUsuariosHistoricos` y `TbAnexosUsuariosHistoricos`. El patrón se traduce a SQLAlchemy `AsyncSession.begin()` (D66, D82) en la nueva plataforma. Test VBA equivalente pytest.

### D-MVVM-traductor · Clases/Servicios/Coordinadores → use-cases/adaptadores

- **Aplicación**: la separación MVVM en VBA se traduce uno a uno — clases de dominio → entidades Pydantic, coordinadores transaccionales → use cases Python con `AsyncSession.begin()`, Test Doubles (`clsTestDouble*`) → doubles pytest (mocks ya en `Dummy*` fixtures).

---

## 5. Criterios de aceptación

Lista verificable de qué define "épica de HPS cerrada".

### Funcionales (paridad con legacy)

- [ ] Las 16+ features de negocio (F1-F16) tienen CRUD en web con paridad funcional mínima (ver § 1.1).
- [ ] El ciclo de usuario HPS (alta, baja, renovación, modificación) preserva los 4 estados calculados (`APuntoDeCaducar`, `Caducado`, `Solicitado`, `PendienteRenovacion`) con su fuente persistida.
- [ ] El histórico de usuarios (`TbUsuariosHistoricos`) preserva trazabilidad; retención ≥ auditoría (D29).
- [ ] Las observaciones (`TbObservaciones` + `TbObservacionesHistoricas`) preservan sincronización con histórico de usuarios (D-Historico-Anexos).
- [ ] Las 4 fuentes de anexos (`TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos`) migran con metadatos; contenido externalizado a object storage (D16); papelera 30 días (D19).
- [ ] Las vinculaciones con Lanzadera (identidad), Expedientes (`IDExpediente`), HPS_Solicitudes (`IDSolicitud`), SICA operan via adaptadores unificados (D9-D10); verifican existencia antes de acciones destructivas.
- [ ] `Test_LocalReadAccessAuthorization.bas` tiene equivalente pytest que valida la regla de autorización local.
- [ ] Los 8 formularios `Form_FormInicial03Consultas<PorCampo>` (Contratistas, Datos, EmpresaUsuario, Estado, Grado, JuridicaTramitacion, MotivoHPS, Nombre, ProyectoAsignado) tienen UI HTMX equivalente.
- [ ] `TbConsultas` migrado a query builder parametrizable, sin SQL injection.

### Config / catálogos

- [ ] Los 5 catálogos seed-only (`motivos_hps`, `cursos`, `hps_equivalencia`, `hps_grado`, `juridicas_contratacion`) están migrados como seed inicial con sus volúmenes originales.
- [ ] Endpoint admin CRUD sobre cada catálogo con control de capabilities (D45).
- [ ] `CacheHabilitada` migrado como config del módulo + kill switch runtime vía endpoint admin.

### Forma y plataforma

- [ ] Módulo hexagonal dentro del monolito modular (D68); puertos por capacidad.
- [ ] Adaptadores driving: web (HTMX) + CLI admin global (D24).
- [ ] Adaptadores driven: PostgreSQL, object storage S3-compatible (D16), secret manager (D9-D10), notificación unificada (D11).
- [ ] Versión semántica del módulo: `hps/v0.1.0-rc.1` en primer RC, `hps/v1.0.0` en release (D78).
- [ ] Migraciones de BD backward-compatibles con Expand and Contract (D82).

### Seguridad y operación

- [ ] **NO existe caché en cliente** (verificación binaria: ausencia de las 12 tablas locales en artefactos web).
- [ ] `PasswordBackend` **eliminado** del código; secret en secret manager (D93 cruzado con D104).
- [ ] Datos personales enmascarados en logs y observabilidad (D92); capabilities declaran visibilidad por campo (D45).
- [ ] Suplantación acotada al administrador global con doble identidad visible (D44); audit completa (D27).
- [ ] Capabilities declaradas en código; UI consulta endpoint de capabilities; servidor rechaza operaciones no autorizadas (D45).
- [ ] `Test_LocalReadAccessAuthorization` pytest equivalente valida autorización local en cada acción.

### Datos

- [ ] D94 formalizado: las 6 FKs físicas reales (corregidas las 2 PK-to-PK) tienen `FOREIGN KEY` en PostgreSQL.
- [ ] D102 aplicado: las 3 columnas booleanas en `TbUsuarios` migradas a `BOOLEAN` con regla explícita.
- [ ] FKs conceptuales cross-app (`IDExpediente`, `IDSolicitud`, `IDEmpresa*`, `IDJuridicaContrato`) se mantienen como referencia conceptual mediate adaptadores.
- [ ] HPS_Solicitudes (ID 22) **NO comparte tabla** con HPS (ID 17) — apps independientes (D83).
- [ ] Migración de datos validada con backfill contra staging + conteos contra `HPST.accdb` autoritativo (~2197 filas principales).
- [ ] Las 4 "Copia de..." residen en zona `legacy` del esquema; `Errores de pegado` evaluado (logs estructurados D27 o tabla de auditoría).
- [ ] **CA-ZOMBIE-Nuevo**: las **9 tablas ZOMBIE** confirmadas en el audit (obs #24097) están formalmente descartadas para migración — no aparecen en el esquema PostgreSQL, no tienen endpoints, no tienen UI. Tratamiento documentado en `migration-matrix.md` con la categoría "no-migrate". Las 4 "Copia de..." y `Errores de pegado` siguen el lineamiento de § 1.4; los 3 catálogos sin uso (`TbHPSEquivalencia`, `TbAuxCursos`, `TbJuridicasContratacion`) y `TbUsuarioAnexos` requieren confirmación con negocio antes de descartar formalmente.

### Testing y calidad

- [ ] Los 9 archivos `Test_*.bas` tienen equivalente pytest que valida el mismo comportamiento (D87).
- [ ] Tests E2E para ciclo HPS completo (alta, baja, renovación, observación, anexo, indicador).
- [ ] Tests de capabilities verifican rechazo de operaciones no autorizadas.
- [ ] Tests de transacciones atómicas (AnexoSelectionTransactionCoordinator + UsuarioLifecycleTransactionCoordinator) verifican rollback ante fallo.
- [ ] Tests de kill switch de indicadores verifican coherencia con caché deshabilitada.
- [ ] Smoke contra staging antes de promover a producción.
- [ ] UAT por admins de aplicación + usuarios de negocio (D47-D48); resultado y excepciones documentadas (D49-D50).

---

## 6. Pendientes operacionales

Acciones manuales que el equipo debe ejecutar antes, durante o después de la migración.

### Antes de empezar

- [ ] **Operativo en `00_HPS`** (no desde aquí): añadir `*.accdb` al `.gitignore` del repo.
- [ ] **Operativo en `00_HPS`** (no desde aquí): verificar `git log --all -- HPS.accdb`; si está versionado, **rotar repo + aplicar git-filter-repo** para limpiar el historial.
- [ ] **Auditar `PasswordBackend` en `00_HPS`** y rotar la contraseña real del backend si era expuesta.
- [ ] **Mover la contraseña a secret manager** (D9-D10). Actualizar manifests de despliegue y dotenv si aplica.
- [ ] **Auditar git history de `00_HPS`** por commits que contengan `PasswordBackend`; documentar si fue commiteada. Si sí, **rotación obligatoria** + nota de incidente.
- [ ] **Respaldar `HPST.accdb` de producción** antes de la migración de datos (snapshot inmutable).
- [ ] ✅ Schema usage audit de HPS committed (obs #24097, topic_key `hps/usage-audit-2026-08-06`): 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE de 22 tablas backend. Cierre del gap del obs #24084.
- [ ] **Confirmar con Natalia** (negocio) la disposición de las **9 tablas ZOMBIE** (audit obs #24097):
  - [ ] 4× "Copia de..." (`Copia de TbExpedientes`, `Copia de TbExpedienteLugares`, `Copia de TbUsuarios`, `Copia de TbUsuariosEntidades`) → retención indefinida como histórico o purga con auditoría. Coordinar con equipo de Expedientes para las dos copias que afectan tablas de esa app (`TbExpedientes`/`TbExpedienteLugares`).
  - [ ] `Errores de pegado` (sentinel) → transformar en log estructurado canónico (D27) o purgar con retención limitada.
  - [ ] 3 catálogos sin uso (`TbHPSEquivalencia`, `TbAuxCursos`, `TbJuridicasContratacion`) → confirmar con negocio si son histórico o se descartan formalmente.
  - [ ] `TbUsuarioAnexos` → confirmar solapamiento con `TbAnexosUsuariosHPS` antes de descartar.
- [ ] **Auditar contenido de `TbConsultas`** y revisar políticas de SQL injection antes de migrar (TK-HPS-9). Adicional: revisar el gap del extractor `m_SQL='TableName'` con el maintainer de `codegraph-vba` antes de clasificar `TbConsultas` como ACTIVE/ZOMBIE.
- [ ] **Revisar plantillas en `Correo.cls`** con datos sensibles (D92) antes de portar al servicio unificado de notificaciones.

### Durante la migración

- [ ] Validar que las **344 filas de `TbDatosLocal`** no migran como caché en cliente (verificación binaria de ausencia de las 12 tablas locales en artefactos web).
- [ ] Validar FKs intra-app corregidas (D94) — específicamente las 2 PK-to-PK que pasan a `IDUsuario`/`IDUsuarioHistorico`. Cualquier inconsistencia debe corregirse en origen antes de promover.
- [ ] Validar booleanos (D102): las 3 columnas traducidas con regla explícita, sin valores fuera de 'Sí'/'No' sin default documentado.
- [ ] Coordinar con migración de HPS_Solicitudes (D83 — apps independientes, no comparten tabla; pero la FK conceptual `IDSolicitud` requiere convención documentada).
- [ ] Probar cache safety (D70-D71) en el puerto de caché — kill switch equivalente al `Mod_StartupCacheInitialization.bas` legacy.
- [ ] Portar `Test_LocalReadAccessAuthorization.bas` a pytest antes del go-live (test crítico de seguridad).

### Después del go-live

- [ ] Confirmar que las 12 tablas locales del frontend NO existen en artefactos web (D92 cerrado).
- [ ] Confirmar que la papelera de adjuntos (D19) funciona end-to-end para las 4 fuentes de anexos.
- [ ] Verificar que la impersonación (D44) solo la inicia el administrador global.
- [ ] Cerrar el gap del audit: ✅ emitido obs #24097 con el veredicto (12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE). Paridad con obs #24085 de Condor lograda.
- [ ] Evaluar encriptación en reposo para datos personales (D92 follow-up) — no bloqueante para v1.
- [ ] Verificar que el patrón de transacciones atómicas (`AnexoSelectionTransactionCoordinator` + `UsuarioLifecycleTransactionCoordinator`) preserva rollback end-to-end en producción.

---

## 7. Tickets derivables (preview — NO crear issues todavía)

> Lista de issues que nacerían de esta épica. **NO** se crean todavía; se trata de un preview para alinear con el equipo y empezar el desglose SDD.

### Núcleo de funcionalidad

- [ ] **TK-HPS-1**: Implementar CRUD Usuario HPS con ciclo completo (alta, baja, renovación, modificación) en backend web — entidades Pydantic + use cases + adaptadores de persistencia + UI HTMX. Preservar los 4 estados calculados (`APuntoDeCaducar`, `Caducado`, `Solicitado`, `PendienteRenovacion`) con su fuente persistida.
- [ ] **TK-HPS-2**: Histórico de usuarios (`TbUsuariosHistoricos`) + usuarios entidades (`TbUsuariosEntidades`) en PostgreSQL — snapshot conservado para auditoría, retención ≥ auditoría (D29).
- [ ] **TK-HPS-3**: Seed de los 5 catálogos (`motivos_hps`, `cursos`, `hps_equivalencia`, `hps_grado`, `juridicas_contratacion`) en migración inicial — scripts Alembic idempotentes con conteos del backend autoritativo.
- [ ] **TK-HPS-4**: Gestión HPS propiamente (`TbHPS`, 1280 filas) con equivalencias y grado — FK a `usuarios` + metadatos del curso + versión del HPS.
- [ ] **TK-HPS-5**: Observaciones (`TbObservaciones`, 330 filas) + históricas (`TbObservacionesHistoricas`) en PostgreSQL con CRUD completo — sincronización atómica con histórico de usuarios (D-Historico-Anexos).
- [ ] **TK-HPS-6**: Anexos con referencia externa a object storage S3-compatible (D16) — 4 fuentes (`TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos`) con metadatos en PostgreSQL y contenido en S3, papelera 30 días (D19).
- [ ] **TK-HPS-7**: Integración SICA externa (`TbUsuariosSICA` + `TbAnexosUsuariosSICA`) — referencia externa, reconciliación por identificador sin asumir propiedad; el sistema SICA queda fuera de scope de la migración.
- [ ] **TK-HPS-8**: Vinculaciones conceptuales cross-app mediate adaptadores unificados (D9-D10) — `IDExpediente` (Expedientes), `IDSolicitud` (HPS_Solicitudes, ID 22 — app independiente, D83), `IDEmpresa*`, `IDJuridicaContrato` como referencias conceptuales.
- [ ] **TK-HPS-9**: Consultas pre-armadas (8 formularios `Form_FormInicial03Consultas<PorCampo>`) + `TbConsultas` query builder — queries parametrizadas sin SQL injection; UI HTMX con refresh manual (D69).

### Forma y plataforma

- [ ] **TK-HPS-10**: Sustituir caché frontend (12 tablas locales) por server-side (D92) — vistas materializadas o queries parametrizadas con caché detrás del puerto (D70-D71); **NO persistir en cliente**.
- [ ] **TK-HPS-11**: Cache selectivo maduro preservado como referencia del puerto de caché (D91 cross-app) — kill switch equivalente al `Mod_StartupCacheInitialization.bas` legacy; Redis como opción detrás del puerto (D71), no dependencia inicial.
- [ ] **TK-HPS-12**: Configuración de backend → config del módulo hexagonal + secret manager (D9-D10) — `BackendActivo`, `EnPruebas`, `IDAplicacion`, `PasswordBackend`, `CacheHabilitada` como variables de entorno del runner; `PasswordBackend` desaparece.
- [ ] **TK-HPS-13**: Transacciones de selección + lifecycle (D-Historico-Anexos) — `AnexoSelectionTransactionCoordinator` + `UsuarioLifecycleTransactionCoordinator` como use cases Python con `AsyncSession.begin()` (D66, D82); equivalentes pytest de `Test_AnexoSelectionTransaction` y `Test_HistoricoAdjuntosTransactionWrapper`.
- [ ] **TK-HPS-14**: Notificaciones transaccionales (`Correo.cls` + `modIndicadores`) — traducir a servicio unificado de notificaciones (D11, D12) canal email v1; revisar plantillas con datos sensibles (D92) antes de portar.
- [ ] **TK-HPS-15**: Indicadores en tiempo real + kill switch atómico (puerto observabilidad) — `Indicador.cls` + `clsIndicadoresBus.cls` + equivalente pytest de `Test_RealTimeIndicatorCoherence`; kill switch preservado.
- [ ] **TK-HPS-16**: Consultas SQL pre-armadas (`TbConsultas`) — traducir a query builder parametrizable; evaluar deprecate vs portar el contenido.

### Seguridad y cumplimiento

- [ ] **TK-HPS-17**: Plan D92 — limpieza operativa del caché frontend con PII (operativo en `00_HPS`, no desde aquí) + **NO migrar la caché a cliente** (D-CacheFrontend) — `*.accdb` al `.gitignore`, `git log --all -- HPS.accdb`, git-filter-repo si aplica; cifrado del frontend o extracción de PII a esquema separado como follow-up.
- [ ] **TK-HPS-18**: Booleanos `Text(2)` → `BOOLEAN` con regla explícita (D102 cross-cutting) — las 3 columnas (`CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`) con regla `'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`.
- [ ] **TK-HPS-19**: Datos personales (DNI, nombres, fechas, correos, teléfonos) — enmascarar en logs y observabilidad (D92) + capabilities declaran visibilidad por campo (D45) + evaluar encriptación en reposo como follow-up.
- [ ] **TK-HPS-20**: FKs conceptuales sin constraint formal (D94) — formalizar las 6 FKs intra-app (corregir las 2 PK-to-PK genéricas por `IDUsuario`/`IDUsuarioHistorico`); mantener las 5 FKs cross-app como referencia conceptual mediate adaptadores; HPS_Solicitudes (ID 22) **NO comparte tabla** (D83).
- [ ] **TK-HPS-21**: Disposición de 4 "Copia de..." y sentinel "Errores de pegado" (D92 dispositions) — "Copia de..." a zona `legacy` del esquema con retención indefinida, sin endpoints; `Errores de pegado` evaluado caso a caso (logs estructurados D27 o tabla de auditoría).
- [ ] **TK-HPS-22**: Suplantación solo por administrador global (D44 cross-cutting) — doble identidad visible + audit completa; integración con adaptador de identidad (D9-D10).
- [ ] **TK-HPS-27**: Resolución de las 9 tablas ZOMBIE (audit obs #24097) — confirmar disposición con Natalia / negocio (mantener histórico, transformar en log, o purgar). Alcance: las 4 "Copia de..." (coordinación con equipo de Expedientes para `Copia de TbExpedientes` / `Copia de TbExpedienteLugares`), el sentinel `Errores de pegado`, los 3 catálogos sin uso (`TbHPSEquivalencia`, `TbAuxCursos`, `TbJuridicasContratacion`), y `TbUsuarioAnexos` (solapamiento presunto con `TbAnexosUsuariosHPS`). Documentar el veredicto en `migration-matrix.md` como categoría "no-migrate". Gate previo a fase SDD de HPS.

### Testing y calidad

- [ ] **TK-HPS-23**: Tests E2E para ciclo HPS completo (alta, baja, renovación, observación, anexo, indicador) — pytest + httpx + Playwright para flujos críticos de UI (D68).
- [ ] **TK-HPS-24**: Preservar cobertura de los 9 archivos `Test_*.bas` como referencia del puerto de testing (D87) — `Test_AnexoSelectionTransaction`, `Test_CacheConsistencyAudit`, `Test_HistoricoAdjuntosTransactionWrapper`, `Test_HPSConfig`, `Test_HPSEntorno`, `Test_LocalReadAccessAuthorization` (**test crítico a portar**), `Test_PerAnexoMove`, `Test_RealTimeIndicatorCoherence`, `Test_StartupCacheInitialization`.
- [ ] **TK-HPS-25**: Tests de capabilities verifican rechazo de operaciones no autorizadas (D45) — el servidor rechaza aunque la UI las muestre.
- [ ] **TK-HPS-26**: Migración de datos + smoke contra staging antes de promover a producción — backfill validado con conteos contra `HPST.accdb` autoritativo (~2197 filas principales); validación de ausencia de las 12 tablas locales en artefactos web (D92 verificación binaria).

---

## Anexo · Tabla de decisiones referenciadas

| Decisión | Tema | Estado | Aplica a HPS |
|---|---|---|---|
| D8 | Hexagonal global | APROBADO | Toda la épica |
| D9-D10 | Autenticación via adaptador | APROBADO | Identidad/permisos (D-Identidad) |
| D11-D13 | Notificación unificada + cola email | APROBADO | `Correo.cls` + `modIndicadores` → email v1 |
| D14 | PostgreSQL compartido por esquemas | APROBADO | Persistencia del módulo |
| D16-D20 | Adjuntos object storage + papelera 30 días | APROBADO | 4 fuentes de anexos (F9) |
| D27-D29 | Logs estructurados + retención | APROBADO | Sustituye sentinel `Errores de pegado` |
| D36-D41 | Credenciales, lockout, caducidad | APROBADO | Integración con Lanzadera (no del módulo) |
| D42-D44 | Autorización + suplantación | APROBADO | Capa de auth + TK-HPS-22 |
| D45-D46 | Capabilities declaradas | APROBADO | Visibilidad por campo (TK-HPS-19) |
| D66-D67 | Stack backend + frontend | APROBADO | Toda la épica |
| D68 | Monolito modular | APROBADO | Estructura del módulo |
| D70-D72 | Caché selectiva + ETag + polling | APROBADO | Sustituye caché frontend (TK-HPS-10, TK-HPS-11) |
| D78-D82 | Versionado + trunk-based + Expand and Contract | APROBADO | Releases y migraciones |
| D83 | HPS_Solicitudes app independiente (ID 22 vs HPS ID 17) | APROBADO | TK-HPS-8 (`IDSolicitud` referencia conceptual, NO tabla compartida) |
| D86-D87 | Forma hexagonal + tests VBA | APROBADO | Referencia de mapeo del puerto (TK-HPS-24) |
| D91 | Caché selectivo maduro cross-app | (confirmado por HPS) | TK-HPS-11 preserva patrón |
| D92 | Disposiciones específicas de HPS (PII frontend) | PROPUESTO (cierre en esta épica) | H1, D-CacheFrontend, TK-HPS-17, TK-HPS-19 |
| D94 | FKs conceptuales intra-app | PROPUESTO (formalización) | H3, TK-HPS-20 |
| D102 | Booleanos `Text(2)` cross-cutting | PROPUESTO | H2, TK-HPS-18 |

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| `docs/03-aplicaciones/hps/capabilities.md` | Inventario de features F1-F16, 29 clases, 30 módulos, 9 Test_*.bas, cache pattern |
| `docs/03-aplicaciones/hps/data-model.md` | 22 schemas, 12 tablas frontend, 6 FKs, esquema `TbUsuarios` 27 columnas, booleanos Text(2) |
| `docs/03-aplicaciones/hps/forms.md` | Navegación, 8 `Form_FormInicial03Consultas*`, call paths críticos, coordinadores transaccionales |
| `docs/03-aplicaciones/hps/integrations-automation.md` | Cross-app (Lanzadera, Expedientes, SICA, HPS_Solicitudes), 6 módulos de caché (D91), flags de operación |
| `docs/03-aplicaciones/hps/migration-matrix.md` | D92 disposiciones específicas (1-7): PII, legacy copies, sentinel, caché frontend (H1), booleanos, FKs conceptuales, histórico ↔ anexos |
| `docs/03-aplicaciones/hps/security-rules.md` | Autorización, PII crítica (D92), indicadores en tiempo real, `LocalReadAccessAuthorization` test |
| `docs/03-aplicaciones/hps/README.md` | Estado del lote 4, hallazgos críticos (H1 incluido), 12 hallazgos, APAP excluido |
| `docs/08-decisiones-y-preguntas-abiertas.md` | Decisiones D5-D87 cross-cutting con detalle |
| `docs/09-arquitectura-objetivo-y-principios.md` | Decisiones D66-D82 arquitectura consolidada |
| engram obs #24084 | Consolidado del descubrimiento: D92 (HPS), D102 cross-cutting, HPS read-heavy, 22 tablas backend |
| engram obs #24082 | Regla: audit de uso antes de migrar (cierre del gap pendiente en HPS) |
| engram obs #24085 | Audit de uso de Condor (15 tablas) — referencia del patrón aplicado a HPS |
| engram obs #24097 | Audit de uso de las 22 tablas backend de HPS (12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE) — referencia canónica del veredicto (topic_key `hps/usage-audit-2026-08-06`) |

---

## Checklist del documento

- [x] Estado del descubrimiento sincronizado con obs #24084.
- [x] Hallazgos D92, D94, D102, D44 con anchor links.
- [x] Decisión D-CacheFrontend (D92) citada con referencia a `migration-matrix.md § D92.4`.
- [x] 16+ features de negocio documentadas (F1-F16).
- [x] 12 tablas locales del frontend identificadas explícitamente con disposición (D92).
- [x] 5 catálogos seed-only identificados.
- [x] Criterios de aceptación verificables y agrupados por dimensión.
- [x] Tickets cubren 16+ features + seguridad + testing (TK-HPS-1..27).
- [x] Schema usage audit de HPS committed (obs #24097, topic_key `hps/usage-audit-2026-08-06`): 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE.
- [x] Las 9 ZOMBIE listadas en § 1.4 y referenciadas desde § 2.1.1, § 3 (H8 D92-HPS-AUDIT), § 5 (CA-ZOMBIE-Nuevo), § 6 (pendiente con Natalia) y § 7 (TK-HPS-27).
- [x] Tabla de decisiones referenciadas (D5-D102).
- [x] Tabla de fuentes.
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] HPS_Solicitudes (ID 22) marcada como app independiente (D83), no comparte tabla.

## Siguiente paso

Revisión con el equipo. El schema usage audit de HPS ya está cerrado (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`). Próximo paso natural de la cadena: ejecutar **TK-HPS-27** (confirmar con Natalia / negocio la disposición de las 9 tablas ZOMBIE — 4× "Copia de...", `Errores de pegado`, los 3 catálogos sin uso y `TbUsuarioAnexos`) y resolver el UNCLEAR de `TbConsultas` (gap del extractor `m_SQL='TableName'`). Una vez cerrada esa disposición, abrir SDD (`sdd-propose` + `sdd-spec` + `sdd-design` + `sdd-tasks`) para arrancar la implementación por ticket, comenzando por **TK-HPS-17** (limpieza operativa del caché frontend con PII en `00_HPS`) antes que cualquier otro cambio de código.