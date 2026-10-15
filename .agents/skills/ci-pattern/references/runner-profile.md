# Perfil de runners (P51–P53) — self-hosted por defecto para repos privados

Motivo documentado (ardelperal/APAP_WEB, 2026-10-12): repo **privado** de
cuenta **Free** — 2 000 min/mes hosted y 500 MB artifacts. Agotados los
minutos, TODO job con `runs-on: ubuntu-*` muere en 2-3 s y el CI queda
cegado mientras el pool propio está ocioso: tres runners del pool `ci`
conectados y `ci.yml` pidiendo runners hosted que ya no podían correr.
El perfil de runners no era un parámetro del patrón: no se podía
declarar, ni gobernar, ni heredar entre repos. P51–P53 lo son desde
aquí; HR-61 lo gobierna.

## Que goberna el gate

`assets/runner-profile/check_runner_profile.py <repo>` lee el
`ci-pattern.yaml` del consumer y sus workflows:

| Declaración | Veredicto |
| --- | --- |
| P51 ausente o `hosted` | gate inocuo (retrocompatible con consumers existentes) |
| P51 `self-hosted` + job `runs-on` del pool P52 (conjunto exacto de etiquetas) | verde |
| job hosted sin entrada en P53 | `HOSTED-SIN-EXCEPCION` (rojo) |
| job hosted con entrada P53 `job: motivo` | verde — el motivo es la causa de indivisibilidad (estilo HR-15) |
| entrada P53 sin motivo (vacío/nulo) | `MOTIVO-VACIO` (rojo) |
| `runs-on` self-hosted con etiquetas ≠ P52 | `FUERA-DE-POOL` (rojo) |
| job sin `runs-on` sobre default hosted | `SIN-RUNS-ON` (rojo: hereda hosted) |
| `runs-on` `${{ ... }}` u otro valor no evaluable | `RUNS-ON-DINAMICO` / `RUNS-ON-DUDOSO` (rojo: el gate no adivina) |
| sin `ci-pattern.yaml`, P51 fuera del enum, o self-hosted sin P52 | entrada inválida, exit 2 (fail-closed, HR-3) |

Alcance: solo workflows con disparador `pull_request` o `push` (los que
consumen minutos por PR). Schedule-only queda fuera y es deliberado: su
consumo se goberna por el plan, no por el pool; un repo que quiera
gobernarlo lo declara en P53 con motivo.

## Adopcion en un consumer (Fase 4 del porting-guide)

1. Declarar el perfil en `ci-pattern.yaml` del consumer:

   ```yaml
   P51_runners_profile: self-hosted
   P52_runners_pool_labels: [self-hosted, Linux, ARM64, apap, oracle, coolify, noble, deploy]
   P53_hosted_exceptions:
     deploy: accion de terceros sin build ARM64 (causa de indivisibilidad, HR-15)
   ```

2. Fijar las etiquetas del pool **por dato**, no de memoria: la lista
   real de etiquetas de cada runner del pool (Fase 0, G0.5 del
   porting-guide: inventario de runners) es la que va en P52. La
   precedente APAP_WEB está documentada en `references/asset-inventory.md`
   (P29). Un pool `self-hosted` pelado goberna CUALQUIER runner
   conectado — incluido un portátil personal —; por eso P52 vacío con
   perfil self-hosted es entrada inválida, no default benigno.

3. Pasar los jobs de PR/push del workflow al bloque del pool:

   ```yaml
   jobs:
     test:
       runs-on: [self-hosted, Linux, ARM64, apap, oracle, coolify, noble, deploy]
   ```

4. Correr el gate una vez ANTES del primer push del cambio (paridad
   local, HR-17) y cablearlo despues como job requerido del consumer con
   su propio `needs` (HR-29).

## Compatibilidad ARM64 (el pool típico de un Oracle Free Tier)

Los Ampere A1 del Free Tier son ARM64. Lo que hay que verificar antes de
migrar un job, por runner del pool (lista de binarios del G0.5), no por
fe:

- `actions/setup-python`: provee builds CPython aarch64 para Linux desde
  hace años; funciona en pool ARM.
- `actions/setup-node`: binarios node-linux-arm64 oficial; funciona.
- Playwright: los builds de navegador para Linux ARM64 existen para
  Chromium y Firefox; **WebKit no esta soportado oficialmente en
  arm64-linux** — los proyectos con cobertura multi-navegador dejan el
  job de WebKit en runner hosted (excepción P53 con motivo) o lo
  declaran fuera del agregador hasta tener evidence. Verificar SIEMPRE
  con el workflow real del pool antes de prometer la migración (HR-18:
  la prueba es el workflow corriendo, no la doc).
- Imagenes de servicio (postgres/mysql/redis): multi-arch oficial en la
  gran mayoria; confirmar con `docker manifest inspect <img>` desde el
  runner antes de migrar el job `services:`.
- accions Docker-based (`node:`/`docker:`): requieren imagen arm64 del
  action; actionlint ya avisa con `runner.labels` configurados (ver
  snapshot APAP_WEB en `testing/consumer-snapshot/`).

## Alerta temprana de consumo (el hueco que esta regla cierra)

Los minutos agotados se manifiestan como fallos 2-3 s sin señal previa
clara. Deteccion, en orden de costo:

1. free-by-hours: en el runner, el job ni arranca — `Get-RunnerStatus`
   local no existe; la senal es el motivo del fallo del job:
   `This workflow recently had a scheduled event that failed due to
   insufficient minutes` vs `no runners meet...` (pool caido). Son dos
   incidentes distintos; no los mezcleis en el runbook (HR-31: el
   veredicto se lee del dato, no de la costumbre).
2. API de billing: `GET /orgs/{org}/settings/billing/actions` devuelve
   `total_minutes_used` y `paid_minutes`; **limitacion documentada**:
   endpoint de organizacion (una cuenta personal Free sin org no lo
   tiene); requiere token de owner; cota 1 req/min. Un job scheduled
   `schedule: cron: '17 6 * * *'` en el pool propio, con secret de
   lectura, puede comparar uso/limites y abrir una issue de aviso
   (`gh issue create --label type:chore`) al 80 %.
3. canario barato: job `no-op` hosted 1-min cada madrugada solo en
   repos que aun no migraron — si muere en segundos, los minutos se
   acabaron; es la senal que APAP_WEB tuvo que deducir a mano.

El patron recomienda (manual, HR-62): perfilar self-hosted ANTES de
agotar minutos, montar la alerta 2 con la limitacion declarada, y usar
3 solo como transitorio hasta completar la migracion.
