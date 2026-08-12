# Runners self-hosted

Este repo corre en `a2w-oracle-vps`, una instancia self-hosted sobre el VPS Oracle
compartido. Hay **dos** runners registrados para redundancia — uno cae y el
otro sigue tomando jobs, sin que un reboot de Coolify o un `docker daemon` colgado
paralice toda la CI.

## Estado actual

| Nombre | `RUNNER_NAME` | Ubicación en el VPS | Servicio systemd |
|---|---|---|---|
| Primario | `a2w-oracle-vps` | `/home/ubuntu/github-runner/a2w/` | `actions.runner.DysTelefonica-access2web-blueprint.a2w-oracle-vps.service` |
| Secundario | `a2w-oracle-vps-2` | `/home/ubuntu/github-runner/a2w-2/` | `actions.runner.DysTelefonica-access2web-blueprint.a2w-oracle-vps-2.service` |

Ambos usan el **mismo** `RUNNER_TOKEN` (PAT con scope `repo`) registrado contra
`DysTelefonica/access2web-blueprint`, así que un `systemctl restart` de cualquiera
de los dos no requiere generar un registration token nuevo (que caduca en 1h).

Los workflows los distinguen por `RUNNER_NAME`; las labels son las mismas:

```
self-hosted, Linux, ARM64, oracle-vps, a2w
```

## Convenciones (PRs #126 / #129 / #131 / #135)

1. **Un `_work` por runner, nunca compartido.** El directorio de trabajo es
   relativo a la instalación (`workFolder: _work`), así que cada runner
   trabaja en su propia ruta:
   - `a2w` → `/home/ubuntu/github-runner/a2w/_work/`
   - `a2w-2` → `/home/ubuntu/github-runner/a2w-2/_work/`

   Si dos contenedores montaran el mismo host path a `_work`, harían checkout
   de dos ramas distintas en el mismo directorio y el resultado es basura
   no-determinista que no se parece a un fallo de test.

2. **Etiquetas idénticas, nombres distintos.** Las labels son filtro de
   enrutado, no identidad. Mantenerlas iguales permite que GitHub Actions
   mande jobs a cualquiera de los dos sin reglas especiales. Lo que los
   distingue es `RUNNER_NAME`.

3. **Sin `services:` con puerto de host fijo** (vacuamente — no tenemos todavía).
   `ports: - 5432:5432` con dos runners compartiendo el mismo host publica el
   mismo puerto y los jobs concurrentes comparten base de datos en silencio.
   La forma correcta es publicar solo el puerto del contenedor
   (`ports: - 5432`) y leer el asignado con
   `${{ job.services.<id>.ports['5432'] }}` dentro de un step, con guardia de
   vacío explícita. El gate `scripts/check_workflows.py` (PR #135) corta este
   patrón antes de que entre.

## Procedimiento para añadir un tercer runner

1. SSH al VPS: `ssh oracle`.
2. `cp -r /home/ubuntu/github-runner/a2w /home/ubuntu/github-runner/a2w-3`.
3. `rm _work .runner .credentials .credentials_rsaparams .env .path` dentro de la copia.
4. Generar PAT con scope `repo` (o reusar el actual).
5. Generar un registration token:
   ```bash
   curl -X POST -H "Authorization: token $PAT" \
     https://api.github.com/repos/DysTelefonica/access2web-blueprint/actions/runners/registration-token
   ```
6. Configurar el runner:
   ```bash
   sudo -u ubuntu ./config.sh \
     --token "$REG_TOKEN" \
     --url https://github.com/DysTelefonica/access2web-blueprint \
     --name a2w-oracle-vps-3 \
     --labels 'self-hosted,Linux,ARM64,oracle-vps,a2w' \
     --work _work \
     --replace \
     --unattended
   ```
7. Instalar el servicio systemd:
   ```bash
   sudo ./svc.sh install ubuntu
   sudo ./svc.sh start
   ```
8. Verificar en GitHub:
   ```bash
   gh api repos/DysTelefonica/access2web-blueprint/actions/runners
   ```
   Debe listar los tres runners como `online`.

## Capacidad del host

| Recurso | Total | Libre (medido) |
|---|---|---|
| vCPU | 4 | — |
| RAM | 23 GiB | 19 GiB (≈ 80 %) |
| Disco | 194 GiB | 66 GiB (≈ 34 %) |

Compartido con: Coolify, Engram, AdGuard, Vaultwarden, los runners de
`ardelperal/APAP_WEB` (cadete, apap-coolify-noble, apap-web) y los 18 contenedores
que mantienen esos servicios. La capacidad libre es suficiente para un tercer
runner, pero un cuarto requeriría medición concurrente.

## Diagnóstico de zombies

Si un runner aparece `online` en GitHub pero sus jobs quedan en cola
indefinidamente, el síntoma típico es un proceso zombie del daemon docker
(observado en el run 31489715170 del #117). Pasos de diagnóstico:

```bash
ssh oracle
systemctl status actions.runner.DysTelefonica-access2web-blueprint.a2w-oracle-vps.service
docker info | head -3
ps -ef | grep a2w-oracle | grep -v grep
```

El preflight `docker info >/dev/null` que añade el gate en `security.yml` y
`security-deep.yml` (PR #129) acorta este caso a 30 segundos en vez de los
15 minutos de timeout — pero no es auto-recovery. El watchdog cron para
auto-restart del daemon queda como seguimiento infra (no implementado).

## Referencia

- Skill canónica con el procedimiento completo: `oracle-vps-github-runners`
  en `DysTelefonica/team-skills`.
- Issues relacionadas: #117 (umbrella original), #130 (RUNNER_TOKEN expiry),
  #131 (security-deep no tenía los gates), #132 (este file), #135 (gate
  en los propios workflows).