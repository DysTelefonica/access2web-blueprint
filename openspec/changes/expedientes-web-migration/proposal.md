# Propuesta: migración web de Expedientes

## Propósito y resultado

Migrar Expedientes al monorepo web sin perder resultados de negocio. El alcance se gobierna por **paridad de capacidades**, no por réplica de 46 formularios ni de mecanismos Access. La épica histórica `docs/03-aplicaciones/expedientes/epic.md` queda **superseded** como plan de destino; conserva valor únicamente como evidencia.

## Alcance

**Incluye:** las 63 entradas con disposición explícita del [catálogo canónico](../../../docs/03-aplicaciones/expedientes/capability-ledger.md), ciclo de vida, datos relacionados, catálogos, consultas/tareas, escritura resiliente, E2E, seguridad, runtime e integraciones.

**No incluye:** copiar PII o rutas locales; paridad de formularios; React/SPA; trasladar Win32, OLE, estado global, cambio interactivo de backend o popups; implementación, migración de datos, UAT o cutover en esta fase.

## Capacidades nuevas

- `expedientes-lifecycle`
- `expedientes-related-data`
- `expedientes-catalogs`
- `expedientes-query-and-tasks`
- `expedientes-write-resilience`
- `expedientes-e2e`
- `expedientes-access-control`
- `expedientes-runtime`
- `expedientes-integrations`
- `expedientes-legacy-retirement`

**Capacidades modificadas:** ninguna; `openspec/specs/` aún no contiene contratos de Expedientes.

## Enfoque y dependencias

Módulo hexagonal dentro del monolito modular (D68), PostgreSQL con schema propio, SSR HTMX/Jinja2/Alpine (D67) y adaptadores para sistemas externos. Autorización deny-by-default mediante el contrato de Lanzadera `CurrentPrincipal/effective_permissions`. La evidencia Dysflow es solo lectura sobre `expedientes-e2e-integration-rest`; su desajuste de configuración impide elevarla a verificación runtime.

## Entrega y áreas afectadas

| Área | Disposición |
|---|---|
| `app/src/modules/expedientes/` | Nueva, por slices de capacidad |
| `app/migrations/` | Expand-and-Contract, schema `expedientes` |
| `docs/03-aplicaciones/expedientes/` | Catálogo canónico y evidencia |

Cada PR tendrá ≤400 líneas cambiadas. Solo se encadenarán PRs por dependencia real; una cadena nunca sustituye la partición.

## Riesgos y rollback

- Contratos incompletos: cerrar field mapping, reconciliación, rechazos, topología, UAT y cutover antes de migrar.
- 49 tablas históricas sin revalidación runtime: no inferir uso ni descarte.
- Integraciones y permisos: contract tests y deny-by-default antes de escritura.

Rollback: despliegue paralelo, legacy intacto, migraciones reversibles y retorno operativo por slice; ninguna retirada se ejecuta antes de reconciliación y aprobación.

## Criterios de éxito y siguientes artefactos

- [ ] Cada entrada del catálogo tiene spec, prueba y disposición trazable.
- [ ] Cero capacidades retiradas por silencio; cero PII/rutas locales copiadas.
- [ ] Reconciliación y UAT demuestran paridad funcional antes del cutover.
- [ ] Siguientes: specs por capacidad, diseño, tasks/PRs, mapping de campos y plan de reconciliación/rollback/UAT/cutover.
