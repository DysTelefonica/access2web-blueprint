# runner-controls — assets que el runner ejecuta sobre el evento

Assets del patrón que corren en el runner (HR-34) con el token del job: leen el
evento y el contrato del host y dejan un veredicto. Solo stdlib; el acceso a la
API vive en `github_rest.py` (reintentos, timeout y `User-Agent` propio).
Las lecturas y las escrituras idempotentes (`PUT .../merge` con `sha`, `PATCH`
de cierre) reintentan un 5xx, un 429 o un error de red con backoff exponencial
(2, 4 y 8 s); un 4xx no se reintenta y el `POST` que crea un comentario, una
issue o un `workflow_dispatch` hace un solo intento (#1592). La regla y su
tabla de casos: `../gh-api-retry/`.

## Contenido

| Fichero | Papel |
|---|---|
| `check_force_push_guard.py` | Guard de force-push y de borrado de una rama protegida por contrato |
| `check_push_compliance.py` | Detective post-push: lo que llega a la rama protegida debe ser un PR mergeado con sus checks en verde |
| `check_governed_merge.py` | Gate pre-merge: el PR entra por el camino gobernado y el merge sale auditado |
| `check_runner_binding.py` | Enlace entre el runner y el contrato del host |
| `github_rest.py` | Cliente de la API de GitHub que usan los tres assets de arriba |
| `safe_merge.py`, `safe_cleanup.py` | Fusión con permisos acotados y limpieza con salvaguarda de Leap |
| `rerun_run.py` | El coordinador relanza un run caído con un solo comando (#1630) |
| `tests/` | Suite ejecutable (solo stdlib, sin red) |

`safe_merge.py` y `rerun_run.py` ejecutan `gh` (como `pr-size` y `ci-pattern report-friction`):
toma `$CI_PATTERN_GH` si está fijada, y una ruta `.py` corre con el mismo
Python; si no, `gh` del PATH, como siempre. Los tests fijan esa variable con la
ruta del `gh` falso y se niegan a correr, sin hacer ninguna llamada, si el `gh`
resuelto queda fuera de su directorio de fixtures (`../tests/_fake_gh.py`, #1609).

## El merge gobernado y el padre de cadena sin fusionar (#1597)

Cuando no hay otro hallazgo, `check_governed_merge.py` lista las PRs abiertas
de cualquier base, paginadas y con el reintento de #1592. Un commit es de la PR
más vieja: solo una PR de número menor que el hijo puede ser su padre. No
cuentan los descendientes declarados (PRs cuya base es la rama del hijo en este
repo) ni una PR con la misma cabeza. Si la cabeza de la otra PR es un commit del
hijo fuera de la base, o el hijo se separa de ella en uno de esos commits
(`compare/{su head}...{head del hijo}?per_page=1`, el padre que avanzó), el hijo
no se fusiona (1):

```
  BLOCKED  PR #<hijo> carries the commits of open PR #<padre> (head <sha>), not yet merged into '<base>': merge or close the parent first, or rebuild the child branch from the base without the parent's commits
```

Si el listado o un compare fallan, o traen otra forma, falla cerrado (2) y el
mensaje nombra la PR del compare. Límites: un padre abierto después que su hijo
o reescrito con force-push no se detecta.

## Parada de línea (#1627)

La regla es la misma en todo repo y la decide `stop_the_line.py`, que el merge gobernado usa como un gate más.

- **Rojo** = hay al menos una issue abierta con la etiqueta `prerelease:night-red` (`--label` la cambia) y la
  última ejecución **completada** de la nocturna (`--nightly-workflow`, por defecto `tests.yml`) en la **rama
  por defecto** del repo no es `success`. Cuenta la ejecución `schedule`. Un `workflow_dispatch` cuenta solo si
  el consumer fija `--dispatch-marker <texto>` y el `display_title` del run (su `run-name`) contiene ese
  texto (p. ej. `suite=full`); sin marcador, ningún dispatch cuenta (#1702). Una ejecución en curso se omite en
  favor de la última completada.
- Con rojo, el PR pasa solo si enlaza una de esas issues **del mismo repositorio**. Si no, sale con 1 y
  muestra la ejecución, sus issues y el comando para relanzarla: `gh workflow run <workflow> --repo <owner/name> --ref <rama por defecto>`.
- Sin señal (ninguna issue etiquetada) el gate pasa. Con issue etiquetada, un workflow nocturno inexistente
  (404) **retiene** (sale con 2) nombrando el workflow y `--nightly-workflow` (#1702); antes pasaba o retenía
  sin decir por qué.
- Con el perfil pre-release activo (`.github/prerelease.json` con `enabled: true`, en la rama por defecto),
  una lectura fallida o una nocturna sin ejecución completada **retiene** (sale con 2) y da el comando.
  Sin perfil, esas dos situaciones pasan.
- Cualquier lectura de issues o de enlaces que falle retiene (fallo cerrado).
- Todas las lecturas van a la rama por defecto, nunca a la base del PR. El gate solo lee.
- **Segunda señal: main en rojo (#1708).** Mientras haya una issue **abierta** con la etiqueta
  `prerelease:main-red` (`--main-red-label` la cambia), un PR pasa solo si enlaza una de esas issues del mismo
  repositorio; si no, sale con 1 y da el desbloqueo: `arregle main en el PR que enlaza #N; el siguiente run verde
  de main la cierra`. Se activa por defecto: la llamada del consumer no necesita ningún flag nuevo. Ambas señales
  se evalúan y cualquiera puede bloquear. Una lectura fallida de esas issues retiene (sale con 2).

## Relanzar un run caído: el coordinador, con un solo comando (#1630)

El colaborador no tiene `actions:write`: pide el relanzamiento al coordinador
del repositorio, que lo ejecuta con su propia App. Cerrar y reabrir la PR no es
la vía de relanzar.

```
python rerun_run.py --repo <owner/name> --run <id> [--all-jobs]
```

Antes de relanzar comprueba que el run es del repo, que terminó `cancelled`,
`failure` o `timed_out` y que su SHA sigue siendo la cabeza de su PR (o de su
rama si no tiene PR); si el SHA ya no es la cabeza, pide empujar. Relanza con
`--failed` salvo `--all-jobs`. La última línea dice el resultado (`RERUN:`,
`REFUSED:` o `FAILED:`). Salida: 0 relanzado · 1 GitHub rechazó el rerun ·
2 una comprobación lo rechazó o una lectura falló.

## El guard de force-push: qué mide y con qué evento (#1266)

El guard se dispara con dos formas de evento y cada una lleva el rango y las
banderas en un sitio distinto:

| Evento | Rango | Banderas `forced`/`deleted` | Veredicto |
|---|---|---|---|
| `push` | en el nivel superior (`before`, `after`) | las trae | `forced: true` o `deleted: true` sobre una rama protegida abre incidencia y sale 1; un push ordinario sale 0 |
| `workflow_dispatch` (el que despacha el merge gobernado, #817) | en `inputs` (`inputs.before`, `inputs.after`) | no las trae: no existen en ese payload | el rango bien formado sale 0 y dice qué midió; el rango mal formado falla cerrado con 2, nombrando la causa |

El rango se lee con `push_range`, el lector de `check_push_compliance.py` que ya
usa el detective post-push (#933): hay una sola implementación de «el rango de
este evento» y el guard no trae la suya. Un dispatch no trae las banderas, y su
ausencia se declara en la salida en vez de callar un veredicto que no se midió:

```
OK: dispatch on protected branch 'main' audited the range <before> -> <after>: well-formed range, no force-push nor deletion reported for it; the event carries no 'forced'/'deleted' flag
```

El guard no busca reescrituras dentro del rango despachado: las reescrituras
llegan como evento `push`, que es el camino que ya audita.

Salidas del guard: 0 veredicto conforme · 1 incidencia abierta (force-push o
borrado de una rama protegida) · 2 fallo cerrado: evento malformado o
incidencia que no se pudo abrir (HR-3, nunca verde sin medición).

## Destino en el consumer

Copie el directorio a `scripts/runner-controls/` del repo consumer (o la ruta
equivalente) y versione allí su propio contrato. Si el consumer también usa
`required-jobs/`, copie `runner-controls/` junto a él: `build_annotations.py`
carga `github_rest.py` por ruta relativa (`../runner-controls/github_rest.py`).
El workflow del consumer declara los `inputs` `before` y `after` bajo
`on.workflow_dispatch.inputs` si quiere que el guard de force-push mida el
rango del merge gobernado; declarar solo uno de los dos bloquea el merge.

## Main en rojo (#1708, parte A)

`main_red.py sync` es el productor del rojo de `main`: lo llama el workflow del consumer sobre el
evento `workflow_run` de su CI (plantilla `main-red.template.yml`). Nadie lo vigila: es determinista.

- Run **completado** en la **rama por defecto**. Si no, sale con `SKIP` (código 0) y no escribe nada.
- Run en **rojo** (`failure`, `timed_out`): reutiliza la issue **abierta** con la etiqueta
  `prerelease:main-red` (comentario con la URL). Si no hay ninguna, abre **una** con el título
  `main en rojo: <workflow> <sha corto>`, el SHA y la PR fusionada (o `desconocido`). Una cancelación
  (concurrencia de main) no es rojo: `SKIP`.
- Run en **verde** (`success`): cierra cada issue abierta con la etiqueta, con un comentario.
- `--dry-run` dice lo que haría y no escribe.
- Cualquier lectura o escritura fallida sale con **2**, con el comando para relanzarlo (es idempotente).

```bash
python .agents/skills/ci-pattern/assets/runner-controls/main_red.py sync --repo OWNER/REPO --run RUN_ID [--dry-run]
```

El gate que bloquea los PR mientras la issue está abierta es la segunda señal de `stop_the_line.py`
(parte B de #1708).
