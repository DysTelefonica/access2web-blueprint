# prerelease (#1627, slice 8)

- **Ensayo nocturno en seco:** `python night_gates.py dry-night --changed <fichero> --head <sha> --record <json>`. Sale `0`, `1` (falta o no vale el ensayo) o `2` (una entrada ilegible). Solo lee. Entrada `--changed`, documentada: `git -c core.quotePath=false diff --no-renames --name-only <base>...<head>`. Con `core.quotePath=false` git no cita los nombres no ASCII; con `--no-renames` un cambio de nombre trae la ruta vieja y la nueva. Una lista vacía, o solo comentarios y líneas en blanco, sale con `2`. Una ruta citada entre comillas también sale con `2`, y el mensaje da el comando correcto. Un BOM inicial (PowerShell 5.1) se descarta. Es un gate autónomo: no está cableado en el merge gobernado; eso corresponde al gate de parada de línea (slice 2) y al productor del registro (slice 3). Las rutas del pipeline son `NIGHT_PATHS`, o las de `--paths-file <json>` (lista de globs, sustituye a las del código); `--night-path` añade una más.
- **Sin coste en PR ordinario:** el ejemplo de CI de `references/porting-guide.md` omite `upgrade-path` salvo que el PR cambie SQL o migraciones (`tests/test_no_cost_pr.py`).

Tests: `tests/test_dry_night_gate.py`, `tests/test_no_cost_pr.py`.
## Informe nocturno (#1627, slice 4)
# prerelease: canalización nocturna (perfil pre-release, #1627, slice 3)

Canalización nocturna sobre la cabeza de `main`: construye, ejecuta los checks caros y
despliega solo si todo está en verde sobre el mismo SHA. Cada paso con regla fija es un
script con tests, no un turno de modelo.

## Piezas

| Fichero | Qué es |
|---|---|
| `nightly_pipeline.py` | `plan`: decide DEPLOY o HOLD; solo lee; llama a la API REST, no al CLI `gh` |
| `redispatch_nightly.py` | Relanza el workflow nocturno sobre main; `--dry-run` no llama a gh |
| `nightly-pipeline.template.yml` | Plantilla del workflow: `schedule`, `workflow_dispatch` y `workflow_call`, con `dry_run` |
| `tests/` | Tests de cada script y de la plantilla |

## Regla de despliegue

`plan` dice DEPLOY solo si se cumplen, en este orden:

1. El perfil `.github/prerelease.json` es esquema 1 y `enabled` es `true`.
2. El SHA que construyó la ejecución es la cabeza actual de `main`. Si `main` se ha
   movido, retiene y dice que hay que relanzar.
3. Los seis jobs requeridos (`build`, `e2e-full`, `integration`, `upgrade-path`,
   `trivy`, `security-deep`) están en `success`. Un job que falta o no terminó retiene.
4. No hay migraciones pendientes. La canalización nunca aplica una sola (perfil,
   regla 4): si hay alguna, retiene.

Cualquier lectura fallida (perfil ilegible, la API que falla, cuerpo que no es JSON)
retiene: falla cerrado. Su mensaje da el comando para desbloquear,
`redispatch_nightly.py --repo <owner/name>`.

La última línea dice el resultado: `DEPLOY:` o `HOLD:`. Código 0 para DEPLOY, 2 para HOLD.

## Uso

```bash
# Decisión (solo lectura). La llama el job gate de la plantilla.
python ci-pattern/assets/prerelease/nightly_pipeline.py plan \
  --repo OWNER/REPO --run RUN_ID --profile .github/prerelease.json --pending-migrations N

# Con los nombres de jobs del consumer (#1699): cada --job lee un job requerido de otro job.
python ci-pattern/assets/prerelease/nightly_pipeline.py plan \
  --repo OWNER/REPO --run RUN_ID --profile .github/prerelease.json --pending-migrations N \
  --job e2e-full=e2e --job trivy=build

# Relanzar la nocturna sobre main; primero en ensayo (no llama a gh).
python ci-pattern/assets/prerelease/redispatch_nightly.py --repo OWNER/REPO --dry-run
python ci-pattern/assets/prerelease/redispatch_nightly.py --repo OWNER/REPO
```

`redispatch_nightly.py` termina con `REDISPATCHED:` (código 0) o `FAILED:` (código 1 si
GitHub rechaza el relanzamiento o no se puede ejecutar gh). Con `--dry-run` la última
línea es `DRY-RUN: would dispatch ...`.

`nightly_pipeline.py` llama a la API con `runner-controls/github_rest.py` (urllib, stdlib) y
lee el token de `GITHUB_TOKEN` o `GH_TOKEN`; no usa el CLI `gh`, que el runner del
consumer no trae (#1691). Los demás assets de este directorio sí lo usan:
`$CI_PATTERN_GH` elige el `gh` como en el resto de assets (#1609).

## Plantilla: qué sustituir antes de usarla

Cada job de la plantilla es un marcador que **falla en cerrado** hasta que el proyecto
pone su comando. El paso `mig` del job `build` debe contar las migraciones sin aplicar
y escribir su número en `$GITHUB_OUTPUT`. Mientras no lo haga, el paso falla y el gate
recibe un valor vacío: `plan` sale con 2 y retiene el deploy. El despliegue exige copia
de seguridad verificada antes (perfil, regla 4); el marcador del job `deploy` lo recuerda.

Con `workflow_call`, GitHub pone a cada job el prefijo `<caller> / `. `plan` compara el
nombre completo y también el sufijo tras ` / `, así que el gate funciona igual en los dos
casos. Un nombre repetido cuenta por su peor resultado.

Un consumidor con su noche ya montada y otros nombres de job no necesita duplicarla:
`--job <requerido>=<job del consumer>` (repetible) hace que `plan` lea el job requerido
del job del consumer, sobre la ejecución programada de su propio workflow (#1699). Por
ejemplo, `--job e2e-full=e2e --job trivy=build` lee `e2e-full` del job `e2e` y `trivy`
del job `build`; dos requeridos pueden leer el mismo job. Los requeridos sin `--job`
conservan su nombre, el prefijo `<caller> / ` se trata igual y los mensajes nombran el
job requerido (`trivy=failure`). Sin `--job`, el comportamiento no cambia. Sale con 2,
con una línea `HOLD:` y sin llamar a la API, si un valor no tiene la forma
`<requerido>=<job>` (falta `=` o un lado está vacío), si el requerido no es uno de los
seis o si el mismo requerido se mapea dos veces.

La plantilla llama al script por la ruta `ci-pattern/assets/prerelease/`. Un consumidor
que vendoriza el asset en otra ruta debe ajustarla en el job `gate`.

## Límites

- `plan` decide con la última ejecución que se le pasa; no mira la siguiente.
- El ensayo del pipeline completo (ejecutar todo sin desplegar) lo da la entrada
  `dry_run` de la plantilla; `plan` solo decide.

# prerelease — informe nocturno (#1627, slice 4)

`night_report.py` convierte una instantánea JSON de la noche en una página HTML:
PRs desde la última noche buena (agrupados y enlazados a sus issues), resultado de cada
job caro y si hubo deploy.

El script no habla con GitHub: la instantánea la recoge el paso de la canalización
nocturna. Así el informe es determinista y se prueba sin red.

## Uso

    python3 night_report.py --snapshot snap.json --dry-run          # imprime, no escribe
    python3 night_report.py --snapshot snap.json --out report.html  # escribe

- `--dry-run` no escribe ningún fichero. Sin `--out` ni `--dry-run` sale con 2.
- Instantánea no válida: sale con 2 y nombra la clave que falta.
- La misma instantánea da el mismo HTML (idempotente).

## Instantánea

Claves obligatorias: `repository`, `night` (con `started_at`), `last_good_night` (con
`finished_at`, o `null`: entonces entran todos los PRs), `prs`, `jobs` y `deploy` (con
`deployed`). Cada PR exige `number`, `title`, `merged_at`, `merge_sha` e `issues` (lista,
puede ir vacía). Cada job exige `name`, `expensive`, `conclusion` y `run_url` (pueden ser
`null`). Las fechas van en ISO 8601 con zona horaria.

Una instantánea que no cumple el contrato sale con 2 y nombra la clave o el campo, sin
traceback. Lo mismo si el fichero no existe.

## Tests

    PYTHONUTF8=1 PYTHONDONTWRITEBYTECODE=1 python3 tests/test_night_report.py

## Runner nocturno (slice 6)

`provision-night-runner.sh` aprovisiona el directorio del runner nocturno de forma
idempotente. Lo lee `night-runner.pins` (las versiones de Python y Node de la CI).

    bash provision-night-runner.sh --runner-dir ~/night-runner --dry-run   # no escribe nada
    bash provision-night-runner.sh --runner-dir ~/night-runner            # aprovisiona

- Si falta una herramienta o su versión no es la fijada, sale con `2` y lo nombra; no escribe nada.
- Si el directorio ya tiene la marca `.night-runner` con `profile=1`, no cambia nada.
  Si la marca tiene otro contenido, sale con `2` y no la sobrescribe.
- Los tests se ejecutan en WSL: `python3 tests/test_provision_night_runner.py`.

## Instantánea (`snapshot.py`, slice 4)

`snapshot.py --snapshot <json>` valida la instantánea de la noche. Sale 0 si cumple el contrato; 2 si no, nombrando la clave o el campo, sin traceback. Lo usa el informe antes de renderizar.
