# Diseño: migración web de Expedientes

## Enfoque técnico

Expedientes será un bounded context hexagonal del monolito modular actual, bajo `app/src/modules/expedientes/`, con las capas ya usadas por Lanzadera: `domain`, `ports`, `application`, `adapters`, `delivery` y `di`. Conserva las 63 capacidades, no formularios Access. PostgreSQL es autoritativo tras cutover; durante la transición el legacy permanece intacto.

```text
HTTP SSR/HTMX ─→ casos de uso ─→ dominio ─→ puertos ─→ PostgreSQL
      │                 │                     ├→ Lanzadera
 Jinja2/Alpine      auditoría/outbox          ├→ storage/documentos
                                              └→ E2E e integraciones
Access snapshot ─→ extractor RO ─→ staging ─→ validación/reconciliación
```

## Registro de decisiones

| Decisión | Elección y razón | Alternativa descartada |
|---|---|---|
| D-EXP-1 Límites | `Expediente` es raíz de ciclo, jerarquía y relacionados; `E2EBatch/ExportSession` es agregado independiente; catálogos son maestros propios. Tareas son proyecciones calculadas. | Agregado único gigante o tabla de tareas inventada. |
| D-EXP-2 Datos | Schema `expedientes`, propiedad exclusiva del módulo; `expedientes_staging` solo para migración. Identidades externas son IDs opacos, sin FK/escritura cruzada. Alembic aplica Expand → backfill → switch → Contract, siempre reversible antes del cutover. | Schema compartido o DDL destructivo. |
| D-EXP-3 Identidad | `CurrentPrincipalPort` y `AuthorizationPort` consumen `effective_permissions`; ausencia/error deniega. Fakes deterministas solo en tests/dev y readiness falla si se enlazan en producción, permitiendo trabajo paralelo hasta el adapter Lanzadera. | Stub permisivo o dependencia directa del schema Lanzadera. |
| D-EXP-4 Escritura | Un comando abre Unit of Work; valida invariantes, compara `version`, registra idempotency key, mutación, auditoría y outbox en una transacción. Conflictos devuelven 409 sin commit parcial. | Last-write-wins o efectos externos dentro de la transacción. |
| D-EXP-5 Lectura/UI | Query services SQLAlchemy Core producen read-models reconstruibles. FastAPI renderiza páginas/fragmentos Jinja2; HTMX actualiza regiones y Alpine solo estado local/accesibilidad (`aria-busy`). Sin SPA ni paridad de pantalla. | ORM del agregado en plantillas o cliente rico. |
| D-EXP-6 Documentos | `DocumentStoragePort` guarda contenido por referencia opaca; metadata, autorización, tombstone y retención viven en el módulo. Upload se prepara, confirma tras commit y limpia huérfanos; proveedor S3/SharePoint queda en adapters. | Rutas locales o binarios en PostgreSQL. |
| D-EXP-7 E2E | Canonicalizador versionado emite `{meta,data}`, `apiVersion=1.0`, nueve colecciones, ISO/null explícito y orden `OrdinalE2E+ID`; expande familias con detección de ciclos. `FileExportAdapter` y futuro `RestExportAdapter` implementan puertos separados. Hash y canonicalización quedan versionados; FNV-1a requiere golden de continuidad antes de aprobarse. | Mezclar transporte y serialización o fijar FNV sin prueba. |
| D-EXP-8 Integraciones | Puertos separados para HPS, AGEDYS, Riesgos, NC y notificación; DTOs anticorrupción, timeouts, idempotencia, errores tipados y contract tests por adapter. | SDKs/HTTP dentro del dominio. |
| D-EXP-9 Migración | Extractor Access read-only sobre snapshot con watermark; carga staging, aplica mapping versionado, cuarentena con motivo, promoción idempotente y reconciliación por conteos/reglas/muestras/hash. Cutover: ensayo, delta final, freeze, reconciliación, UAT y switch; rollback devuelve tráfico al legacy, sin reverse-write automático. | Importación directa a tablas finales. |

## Archivos y entrega

Crear el árbol del módulo anterior, `tests/expedientes/`, templates en `delivery/http/templates/expedientes/`, migraciones `app/migrations/versions/*_expedientes_*.py` y runbooks/artefactos en `docs/03-aplicaciones/expedientes/migration/`.

Cada PR debe quedar por debajo de 400 líneas e incluir comportamiento, prueba y rollback. Cadena base necesaria: puertos/fakes → schema/UoW → repositorios. Después corren en paralelo: (a) ciclo/relacionados/catálogos, (b) read-models/SSR, (c) E2E canonical/batch/file, (d) cada integración, (e) extractor/staging. Dentro de cada carril se encadenan solo dependencias; REST E2E, cutover y retirada legacy esperan contratos y reconciliación.

## Verificación y UAT

Strict TDD: unitarios de invariantes/canonicalización; integración con PostgreSQL y storage reales; contratos para repositorios, Lanzadera e integraciones; HTTP/fragmentos y accesibilidad; Playwright en flujos críticos; golden E2E; migración idempotente, cuarentena y reconciliación. UAT recorre `EXP-CAP-001..063` por perfil, estados vacíos, errores y concurrencia; nunca compara pantallas.

## Trazabilidad

| Specs | Componentes |
|---|---|
| 001–024 | agregados, repositorios, catálogos |
| 025–032 | queries, read-models, SSR, UoW/idempotencia |
| 033–042 | canonicalizador, hash, batch, sinks |
| 043–050 | principal/authz, auditoría, DI/readiness/cache |
| 051–056 | adapters de integración y documentos |
| 057–063 | extractor, UAT, cutover y retirada verificable |

## Matriz de amenazas

| Frontera | Aplicabilidad |
|---|---|
| Paths tipo documentación | N/A: no se clasifican ni ejecutan ficheros. |
| Selección Git; estado commit; push; comandos PR | N/A: el producto no automatiza VCS. |

## Riesgos, artefactos pendientes y decisiones abiertas

| Riesgo/gap | Control o decisión necesaria |
|---|---|
| 49 tablas y campos no revalidados | Diccionario fuente, mapping campo-a-campo, ownership, transformaciones y claves; no diseñar columnas finales antes. |
| Pérdida o duplicado | Taxonomía de rechazos, manifest/watermarks, plan y reporte de reconciliación, runbooks de ensayo/cutover/rollback. |
| Contratos incompletos | Decidir dígito CPV, límite de anexos, atomicidad/ciclo batch, continuidad FNV, ownership/errores AGEDYS, REST E2E, retenciones/proveedor documental. |
| Aceptación insuficiente | Matriz UAT capability→perfil→caso→evidencia y criterio de go/no-go aprobados. |
