# Runners de GitHub Actions

Este documento define la frontera de confianza de los runners del repositorio.
La API de GitHub es la fuente de verdad para el inventario operativo.

## Regla del repositorio

Un repositorio público **no consume los runners propios del VPS**. Todos los jobs
corren en runners hospedados por GitHub, con la etiqueta que necesite cada uno:

| Necesidad del job | Etiqueta |
|---|---|
| Git, `gh`, `jq` o `cosign` contra la API y el registro | `ubuntu-24.04` |
| Imagen o escaneo de arquitectura ARM64 | `ubuntu-24.04-arm` |

El motivo: en un repositorio público los runners hospedados son gratuitos y ARM64
nativo está disponible con `ubuntu-24.04-arm` (GA para repos públicos desde el
2025-08-07), mientras que un job propio acopla el CI a una máquina que hay que
mantener, comparte carga con otros repositorios y, cuando se cae, deja el job en
cola sin runner.

## Frontera de confianza

Los pull requests públicos contienen código no confiable. Todos los jobs que
pueden ejecutarlos usan `ubuntu-24.04`, un runner efímero hospedado por GitHub.

| Origen | Infraestructura permitida |
|---|---|
| `pull_request` | Runner hospedado por GitHub con etiqueta literal |
| `push` a `main` | Runner hospedado por GitHub |
| `schedule` y `workflow_dispatch` | Runner hospedado por GitHub |
| Tag de release `v*` | Runner hospedado por GitHub |

El gate `scripts/check_workflows.py` rechaza **cualquier** job que use
`self-hosted`, alcance o no a los pull requests, y además rechaza —para los
alcanzables desde un PR— una matriz o una expresión dinámica.

La regla se mide, no se predica: hasta #800 la exención de los jobs excluidos de
PR dejaba fuera del gate al job `mutation`, pineado a
`[self-hosted, Linux, ARM64, oracle-vps, a2w]` —una etiqueta con **cero runners
registrados**—, que llevaba cuatro semanas en rojo mientras el agregador
`required` publicaba verde (#785).

Única excepción (#1074): el runner Windows del propietario, etiqueta
`dysflow-erd`, para extraer ERD de solo lectura de los `.accdb`. Solo se admite
en un workflow cuyo único disparador sea `workflow_dispatch`
(`.github/workflows/dysflow-erd.yml`); cualquier otro disparador en ese workflow
hace fallar el gate. El runner está en línea solo cuando el propietario lo
arranca: un job en cola no es un fallo.

Un job propio dentro de un workflow de PR debe excluir ese evento mediante una
condición estática sobre `github.event_name`.

## Inventario

Consulte el estado actual antes de operar sobre un runner:

```bash
gh api repos/DysTelefonica/access2web-blueprint/actions/runners \
  --jq '.runners[] | {name, status, busy, labels: [.labels[].name]}'
```

El 7 de octubre de 2026 esa consulta devolvía `total_count: 0`: el repositorio no
tiene runners propios registrados. El inventario es una instantánea. No copie
nombres, rutas ni servicios desde este documento para operar el host; vuelva a
consultar GitHub y el runbook de infraestructura autorizado.

## Cargas por workflow

| Workflow | Jobs | Etiqueta |
|---|---|---|
| `ci.yml` | `pr-gate`, `quality`, `integration`, `e2e-smoke`, `pip-audit-alert`, `docs-architecture`, `docs-codeguide`, `docs-links`, `docs-commands`, `docs-decisions`, `governance-seals`, `ci-images`, `required` | `ubuntu-24.04` |
| `pr-gate.yml` | `pr-gate` | `ubuntu-24.04` |
| `pr-gates.yml` | `canon`, `pr-gate` | `ubuntu-24.04` |
| `operating-doc-drift.yml` | `operating-doc-drift` | `ubuntu-24.04` |
| `security.yml` | `gitleaks`, `trivy-config` | `ubuntu-24.04` |
| `pip-audit.yml` | `dependency-scan` | `ubuntu-24.04` |
| `security-deep.yml` | `gitleaks-history`, `trivy-image` | `ubuntu-24.04-arm` |
| `codeql.yml` | `codeql` | `ubuntu-24.04` |
| `release.yml` | `preflight`, `verify` | `ubuntu-24.04` |
| `release.yml` | `mutation`, `stage`, `e2e`, `publish` | `ubuntu-24.04-arm` |

Cada fila sale de la clave `runs-on:` del job citado: `ci.yml` (jobs en líneas 86 a 1080, `runs-on` por job), `pr-gate.yml:75`, `pr-gates.yml:30` y `:70`, `operating-doc-drift.yml:41`, `security.yml:38` y `:123`, `pip-audit.yml:31`, `security-deep.yml:29` y `:91`, `codeql.yml:20`, `release.yml:20`, `:90`, `:185`, `:249`, `:354` y `:455`. Si un job cambia de runner, cambie la fila y la cita en el mismo PR.
