[← Back to capabilities-index](../capabilities-index.md) · [← Back to DOCS](../../../DOCS.md)

# Expedientes · runtime (D52 / EXP-CAP-047..050)

Resumen a nivel repo del slice D52 del task plan. Fuente de verdad: [`openspec/changes/expedientes-web-migration/specs/runtime.md`](../../../../openspec/changes/expedientes-web-migration/specs/runtime.md).

## Capacidades

| ID | Título | Escenario clave |
|---|---|---|
| [EXP-CAP-047](../../../../openspec/changes/expedientes-web-migration/specs/runtime.md#requirement-exp-cap-047--readiness) | Readiness | Camino feliz: sistema arrancado y dependencias (Lanzadera, Riesgos, NC) disponibles → contrato `/readyz` verde. |
| [EXP-CAP-048](../../../../openspec/changes/expedientes-web-migration/specs/runtime.md#requirement-exp-cap-048--configuración-tipada) | Configuración tipada | Camino feliz: env vars declaradas y validadas con tipos → runtime rechaza config inválida en arranque, no a mitad de request. |
| [EXP-CAP-049](../../../../openspec/changes/expedientes-web-migration/specs/runtime.md#requirement-exp-cap-049--caché-reconstruible) | Caché reconstruible | Camino feliz: caché con TTL y fingerprint de config → se reconstruye automáticamente al cambiar config sin reinicio. |
| [EXP-CAP-050](../../../../openspec/changes/expedientes-web-migration/specs/runtime.md#requirement-exp-cap-050--binding-de-backend) | Binding de backend | Camino feliz: secret + backend selector + test sandbox → binding explícito por entorno, sin fallback a credenciales hardcoded. |

Cada Requirement tiene además escenarios de validación y de concurrencia o fallo. Ver el spec.

## Cómo se aplica a access2web-blueprint

- **D77** (Docker desde día uno): `app/Dockerfile` (`python:3.12-slim-bookworm`) expone `/readyz` y `/healthz`; el readiness check no depende de las dependencias cross-app en producción pero sí en pre-prod.
- **D70+D71** (no caché de contadores; Redis detrás del `CachePort`): el `CachePort` con `TTLCache` in-process en MVP (DA-8); Redis queda como opción futura detrás del puerto, no como dependencia obligatoria.
- **D9-D10+D25** (SecretManagerPort): el `SecretManagerPort` con `EnvSecretManagerAdapter` en MVP; producción intercambiable por Vault o AWS Secrets Manager sin tocar el dominio.
- **D93+D104** (sin contraseñas hardcoded): Brass y Condor ya están en cleanup; Expedientes hereda la regla via `SecretManagerPort`.

## Core invariants

- **`/readyz` refleja estado real**: el check falla si alguna dependencia crítica (Lanzadera, DB) está caída; no usa timeouts generosos para aparentar verde.
- **Configuración validada en arranque**: tipos estrictos (Pydantic / dataclasses) en `app/src/config.py`; runtime aborta con error diagnosticable si la config no pasa, no espera a runtime errors.
- **Caché con fingerprint de config**: cualquier cambio de env var relevante invalida la caché; el fingerprint se calcula al construir el container.
- **Backend binding explícito**: el selector de backend (`BackendActivo` en `TbConfiguracionBackends`) está pinned por entorno; ningún fallback a `"production"` silencioso.

## Contributor checklist

- [ ] Si el PR añade readiness checks, cubre todas las dependencias críticas (Lanzadera, Riesgos, NC, DB).
- [ ] Si el PR añade config, los tipos son estrictos y la validación ocurre al arrancar (no por-request).
- [ ] Si el PR modifica la caché, el fingerprint de config invalida correctamente; tests pinean esto.
- [ ] Si el PR añade binding de backend, no introduce fallback silencioso a credenciales hardcoded.
- [ ] El PR respeta los 5 gates del workflow y es ≤ 400 líneas.

## Lista de comprobación final

- [ ] Las 4 capabilities referencian anchors al spec fuente.
- [ ] Las decisiones D77/D70/D71/D9-D10/D93/D104 citadas están vigentes en `docs/architecture.md`.
- [ ] Tono castellano peninsular formal con usted.
- [ ] Cross-references desde DOCS.md y `docs/03-aplicaciones/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [access-control](access-control.md) | Next: [integrations](integrations.md)
