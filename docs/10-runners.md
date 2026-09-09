# Runners de GitHub Actions

Este documento define la frontera de confianza de los runners del repositorio.
La API de GitHub es la fuente de verdad para el inventario operativo.

## Frontera de confianza

Los pull requests públicos contienen código no confiable. Todos los jobs que
pueden ejecutarlos usan `ubuntu-24.04`, un runner efímero hospedado por GitHub.

| Origen | Infraestructura permitida |
|---|---|
| `pull_request` | Runner hospedado por GitHub con etiqueta literal |
| `push` a `main` | Runner hospedado por GitHub |
| `schedule` y `workflow_dispatch` | Runner propio cuando el job lo exige |
| Tag de release `v*` | Runner propio |

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

El 9 de septiembre de 2026, GitHub registraba un único runner:

| Nombre | Estado observado | Etiquetas |
|---|---|---|
| `a2w-coolify-noble-2` | `online`, libre | `self-hosted`, `Linux`, `ARM64`, `oracle-vps`, `a2w` |

El inventario es una instantánea. No copie nombres, rutas ni servicios desde
este documento para operar el host; vuelva a consultar GitHub y el runbook de
infraestructura autorizado.

## Cargas permitidas en el runner propio

| Workflow | Jobs | Entrada confiable |
|---|---|---|
| `ci.yml` | `mutation` | Programación semanal o ejecución manual |
| `security-deep.yml` | Todos | Programación semanal o ejecución manual |
| `release.yml` | Todos | Tag `v*` creado por un mantenedor |

`ci.yml`, `security.yml` y `codeql.yml` ejecutan sus jobs de PR en runners
hospedados. Un cambio de esta tabla exige actualizar el gate y sus pruebas.

## Credenciales y registro

No conserve un PAT como `RUNNER_TOKEN`. GitHub emite tokens de registro de corta
duración; genérelos justo antes de registrar o sustituir un runner.

No documente credenciales, rutas del host ni nombres de servicios sin comprobarlos
en la infraestructura autorizada. GitHub puede mostrar un runner `online` sin
probar el estado del daemon Docker ni del servicio del host.

## Diagnóstico desde GitHub

1. Consulte `status` y `busy` mediante la API anterior.
2. Abra el job en cola y confirme las etiquetas solicitadas.
3. Compruebe que sólo las cargas de la tabla usan las etiquetas `self-hosted`.
4. Si el runner sigue bloqueado, aplique el runbook del host con autorización
   explícita para esa infraestructura.

Los jobs Docker ejecutan un `docker info` con timeout antes del escaneo. Este
preflight convierte un daemon bloqueado en un fallo explícito, pero no repara el
host.

## Referencias

- [`calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) — contrato de gates y
  protección de `main`.
- `oracle-vps-github-runners` — procedimiento operativo general para runners.
- Issues #117, #130, #131, #135 y #552 — evolución del enrutado y sus gates.

[← Volver a architecture.md](architecture.md)
