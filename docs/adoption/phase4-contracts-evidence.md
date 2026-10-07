# Fase 4 — Contratos cableados y ejecutados contra datos reales

> **Qué es.** La evidencia de la fase 4: los cinco contratos del patrón declarados en `phases."4".gates`, cada uno ejecutado con datos reales del destino, más el equivalente de issue-spec con su procedencia.

> **Qué no es.** Un cambio de los gates de calidad del repositorio: los 14 `scripts/check_*.py`, el plugin de cobertura y sus umbrales quedan como están.

## Los cinco contratos, con sus entradas y sus veredictos

| Contrato | Entradas declaradas | Veredicto medido |
|---|---|---|
| `host-readback` | `--contract @.github/host-contract.json` y las cuatro instantáneas (`repo`, `labels`, `branch-protection`, `rulesets`) en `docs/adoption/` | `VERDICT: PASS`; única regla `documented-only`: `chain:partial` |
| `required-jobs` | `--policy @.github/required-jobs-policy.json --event pull_request --needs-file @.github/needs/needs.json` | `VERDICT: PASS (event: pull_request)` |
| `local-ci-parity` | `--repo @. --policy @.github/local-ci-parity-policy.json` | `LOCAL/CI PARITY OK: 6 gates comunes (exclusiones aplicadas: 6)` |
| `workflow-policy` | `--workflow @.github/workflows/ci.yml --base-branch main --required-jobs @.github/required-jobs-policy.json` | `WORKFLOW-POLICY OK: 1 workflow(s), 5 required job(s) checked; no findings` |
| `hr-matrix` | sin argumentos: el gate trabaja sobre la matriz del patrón | `HR-GATE-MATRIX OK: 56 reglas — gate 38, manual 0, process 18` |

Sin `runner-binding`: ese contrato solo se exige cuando el host declara camino runner y este repositorio tiene la protección de rama disponible, así que sus controles son `host-enforced` (decisión ratificada en la fase 0).

## La única lista de gates

`scripts/local-preflight.sh` es la lista: los mismos seis gates, en el mismo orden, que el job `quality` de `ci.yml`, y `make verify` delega en él. La política de paridad declara seis exclusiones, todas del paso de instalación del runner (sparse-checkout, reconstrucción del blob de la rama base, `pip install` desde el fichero con hashes y sus ramas de diagnóstico). El test que antes leía la receta del Makefile ahora fija la delegación, la igualdad de los gates y su orden, y que no se corra ninguno fuera del contrato; quitar un gate del script lo pone en rojo.

## Equivalente de issue-spec, con su procedencia

El patrón **no publica** un gate de issue-spec portable: su implementación vive en el repositorio de origen del patrón (APAP_WEB), ruta `scripts/check_issue_specs.py`, en estado de plantilla, y el propio patrón declara que la contraparte issue-side es documental, no de gate. Este repositorio declara su equivalente así:

- El formulario canónico (`.github/ISSUE_TEMPLATE/issue-canonical.yml`) fija las seis secciones con sus nombres exactos y en su orden.
- `tests/test_governance_templates.py` valida la forma del formulario y la de **una issue real**: el cuerpo de la issue #794, capturado por la API el 2026-10-07 y guardado en `tests/fixtures/issue-canonical-794.md`, tiene que llevar las seis secciones en el mismo orden.
- No se copia el script del origen: no está publicado y copiarlo sin su árbol verificado sería inventar el equivalente.

## La decisión de no instalar el binding evidencia↔deploy

Registrada en `phases."0".deploy_binding` con su evidencia: el repositorio no aloja ningún mecanismo de despliegue (`docs/11-releases.md` se declara «No es: una guía de despliegue en producción», `release.yml` publica la imagen y verifica su firma sin desplegar, y no hay etiquetas ni releases). Sin deploy no hay evidencia que registrar por SHA desplegado.

## Superficie declarada y pendientes

- El contrato `workflow-policy` se ejecuta sobre `ci.yml`, que es el workflow que define y publica los cinco nombres requeridos. `security.yml`, `codeql.yml`, `release.yml` y `security-deep.yml` quedan fuera con su motivo en el PR de la rebanada 2/3.
- Pendiente declarado, con issue propia: el lock está atrás en dos transitivos con arreglo disponible (`starlette 0.50.0`, `urllib3 2.7.0`), y el venv de auditoría de dependencias mide una resolución fresca mientras los gates instalan el lock. Issue #794, con las versiones y los avisos exactos.
