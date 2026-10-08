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

El gate `scripts/check_workflows.py` rechaza cualquier job alcanzable desde un
pull request que use `self-hosted`, una matriz o una expresión dinámica.

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
| `ci.yml` | `review-budget`, `quality`, `required` | `ubuntu-24.04` |
| `ci.yml` | `mutation` | runner propio del VPS (transitorio; lo migra #785) |
| `security.yml` | `gitleaks`, `trivy-config`, `pip-audit` | `ubuntu-24.04` |
| `security-deep.yml` | `gitleaks-history`, `trivy-image` | `ubuntu-24.04-arm` |
| `codeql.yml` | `codeql` | `ubuntu-24.04` |
| `release.yml` | `preflight`, `verify` | `ubuntu-24.04` |
| `release.yml` | `e2e`, `publish` | `ubuntu-24.04-arm` |
