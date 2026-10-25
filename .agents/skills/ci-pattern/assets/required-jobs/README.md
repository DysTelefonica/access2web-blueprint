# required-jobs — agregador fail-closed de jobs requeridos

Primer asset ejecutable de `ci-pattern` (issue DysTelefonica/team-skills#132).
Sustituye al `check_required_jobs.py` del consumer de origen: la lógica es la
misma, pero ningún nombre de job, evento o skip vive en el código — todo eso
es política (`required-jobs.policy.example.json` como molde de ejemplo).

## Contenido

| Fichero | Papel |
|---|---|
| `check_required_jobs.py` | Gate: veredicto fail-closed sobre el objeto `needs` serializado |
| `build_annotations.py` | Generador del fichero de anotaciones del check rojo, a partir de la API de check runs (#1219) |
| `check_parity.py` | Comprobación de paridad entre el workflow, el `needs` del agregador y la política |
| `required-jobs.policy.example.json` | Política de ejemplo: jobs conocidos y skips aceptados por evento |
| `tests/` | Suite ejecutable (solo stdlib): veredictos, fail-closed y cableado documentado |

## Destino en el consumer

Copie el directorio a `scripts/required-jobs/` del repo consumer (o la ruta
equivalente) y versione allí su propia política; la de ejemplo es un molde,
no una configuración. Copie también `runner-controls/` junto a él, en
`scripts/runner-controls/`: `build_annotations.py` carga `github_rest.py` por
ruta relativa (`../runner-controls/github_rest.py`). Sin ese fichero el
generador avisa y no escribe anotaciones, y la causa degrada a `no annotation`.
Fuera de esos dos directorios no hay nada más que instalar: el gate es stdlib-only.

## Cableado del job agregador

```yaml
required:
  needs: [lint, compile, smoke, unit, full]   # la misma lista que la política declara
  if: always()
  runs-on: ubuntu-latest
  # `checks: read` es lo que permite leer las anotaciones de los checks del PR.
  permissions:
    contents: read
    checks: read
  steps:
    - uses: actions/checkout@v4
    - name: Veredicto del agregador
      env:
        NEEDS_JSON: ${{ toJSON(needs) }}
        EVENT_NAME: ${{ github.event_name }}
        RUN_ATTEMPT: ${{ github.run_attempt }}
        BASE_REF: ${{ github.base_ref }}
        BASE_SHA: ${{ github.event.pull_request.base.sha }}
        REPO: ${{ github.repository }}
        HEAD_SHA: ${{ github.event.pull_request.head.sha || github.sha }}
        GH_TOKEN: ${{ github.token }}
      run: |
        policy=scripts/required-jobs/required-jobs.policy.json
        # Las anotaciones de un check solo existen en la API de check runs: el
        # generador las lee y escribe el fichero. Si no puede (API caida, sin
        # token), avisa y no escribe: el gate degrada a «no annotation» y el
        # veredicto no cambia.
        printf '%s' "$NEEDS_JSON" > "$RUNNER_TEMP/needs.json"
        python3 scripts/required-jobs/build_annotations.py \
          --policy "$policy" --needs-file "$RUNNER_TEMP/needs.json" \
          --repo "$REPO" --sha "$HEAD_SHA" --event "$EVENT_NAME" \
          --out "$RUNNER_TEMP/annotations.json"
        set -- --policy "$policy" --event "$EVENT_NAME" --run-attempt "$RUN_ATTEMPT"
        if [ -f "$RUNNER_TEMP/annotations.json" ]; then
          set -- "$@" --annotations "$RUNNER_TEMP/annotations.json"
        fi
        if [ "$EVENT_NAME" = "pull_request" ]; then
          # La comparación es contra el SHA base fijo del PR (el que
          # GitHub registró al abrir/sincronizar el PR), no contra la
          # punta móvil de la rama base: una actualización de la base
          # durante el PR no cambia lo que el gate toma como "antes".
          git fetch --no-tags origin "refs/heads/$BASE_REF"
          base_entry=$(git ls-tree --name-only "$BASE_SHA" -- "$policy")
          if [ -n "$base_entry" ]; then
            git show "$BASE_SHA:$policy" > "$RUNNER_TEMP/base-policy.json"
            set -- "$@" --base-policy "$RUNNER_TEMP/base-policy.json"
          fi
        fi
        printf '%s' "$NEEDS_JSON" | python3 scripts/required-jobs/check_required_jobs.py "$@"
```

El objeto `needs`, el nombre del evento, el intento de ejecución y la rama base
llegan al script por variables de entorno (`env:`), nunca interpolados dentro
de `run:`. Una expresión `${{ ... }}` escrita en el script se sustituye antes
de que el shell lo lea: una comilla en una salida de job rompería el paso.

El generador de anotaciones necesita el token del propio workflow
(`github.token`) y el permiso `checks: read` para leer los check runs del PR.
No hace falta ningún PAT ni ningún secreto nuevo.
`toJSON(needs)` entrega, por cada job, un objeto `{"result": ..., "outputs": ...}`.
El gate lee la conclusión del campo `result`; también admite la forma plana
`{"job": "success"}`. Cualquier otra forma sale por exit 1. La entrada llega
por stdin o por `--needs-file`.

Marque `required` como check requerido en la protección de rama.

## El intento juzgado y la causa del check rojo (#1208)

Dos entradas opcionales amplían el veredicto; ninguna cambia qué se considera
verde.

`--run-attempt N` declara qué intento de la ejecución juzga el gate
(`github.run_attempt`) y el veredicto lo dice (`judged attempt 2`): un re-run
que pasa de rojo a verde —o al revés— nunca se lee como el estado del intento
anterior. Sin la opción esa línea no se imprime y el veredicto queda igual que
antes. El número debe ser un entero positivo; otro valor es una duda y sale
por exit 1.

`--annotations FICHERO` nombra el check rojo de verdad: su nombre, su
conclusión y su primera anotación de error cuando existe. El fichero es un
objeto JSON de `job -> mensaje` (por ejemplo
`{"test": "pytest: 3 failed"}`); el gate no llama a ninguna API, quien lo usa
produce el fichero y se lo pasa. Sin fichero, la causa dice `no annotation` en
lugar de callarse (`test (failure): no annotation`); un fichero ilegible o
mal formado es una duda y el gate sale por exit 1.

El fichero lo produce `build_annotations.py`: lee los check runs del SHA de la
PR por la API (`GET repos/{repo}/commits/{sha}/check-runs` y, para cada check
rojo, `GET repos/{repo}/check-runs/{id}/annotations`) y escribe la primera
anotación de error de cada uno. Si el API no se puede leer, avisa por stderr y
**no escribe el fichero**: el paso no lo pasa al gate y la causa degrada a
`no annotation`. El veredicto no depende de las anotaciones, así que un API
caído no pone el job en rojo.

Solo se piden las anotaciones de un check **rojo**: uno ya terminado con el
nombre publicado y una conclusión que no sea verde, el mismo criterio que el
lado del `needs` usa para decidir que un job está rojo. Un check verde no
presta su anotación a un job rojo, porque su texto no es la causa: el gate lo
imprimiría como si lo fuera. Cuando no hay ningún check rojo con el nombre
publicado, el job se queda sin anotación (la causa degrada a `no annotation`) y
el generador avisa por stderr del motivo.

### El nombre del check publicado

El runner publica cada job con el nombre que declara su workflow, que no
siempre es el id del job (`name: lint (${{ github.event_name }})` publica
`lint (pull_request)`). Para que la causa se pueda casar con el check que se
ve en la UI, la política declara el mapeo en la clave `check_names`:

```json
"check_names": {"lint": "lint ({event})"}
```

Solo se declaran las divergencias: si el nombre publicado es el id del job, no
hace falta declararlo. El marcador `{event}` se sustituye por el evento en
curso, para los nombres que lo llevan. Con la declaración, la causa nombra el
check publicado y conserva el id del job:

```
blocked by: lint (pull_request) [lint] (failure)
root-cause failures (fix these first):
  lint (pull_request) [lint] (failure): PHP syntax error in src/Foo.php:12
```

Sin declaración, la causa sale como siempre (`lint (failure): ...`). Declarar
el nombre no es una relajación: no cambia qué jobs exige el gate, así que no
exige campo `relaxation`.

## Contrato de la política

| Clave | Obligatoria | Contenido |
|---|---|---|
| `required_jobs` | Sí | Lista no vacía de nombres de job, sin duplicados |
| `events` | Sí | Un objeto por evento que el workflow atiende |
| `events.<evento>.accepted_skips` | Sí | Jobs de `required_jobs` que ese evento puede saltar; `[]` si ninguno |
| `dependencies` | No | Grafo `{downstream: [upstream, ...]}` para aceptar skips en cascada cuando algún upstream está en `accepted_skips` (#853) |
| `docs_only_lane` | No | Lane condicional: `event`, `gate_job`, `output` y `skips`. El job `gate_job` debe publicar `output` = `"true"` en `event`; si no, esos skips no se aceptan (#857) |
| `marker_lanes` | No | Lista de lanes condicionales con la misma forma que `docs_only_lane` más `equals` opcional: el lane acepta sus `skips` solo si el marker de `gate_job.output` vale exactamente `equals` en `event` (#863) |
| `state_dependent_gates` | No | Lista de gates cuyo veredicto depende de estado externo (HR-39). La lee el gate de política de workflows del **mismo** fichero: cada entrada declara `name`, `command`, `pr_workflow`, `pr_job`, `manifests`, `scheduled_workflow` y `alert_command` |
| `check_names` | No | Nombre con el que el runner publica cada check: `job -> nombre`. Solo las divergencias; `{event}` se sustituye por el evento en curso (#1219) |
| `relaxation` | Solo al relajar | Objeto con `reason` y `reference`, ambos cadenas no vacías |

`accepted_skips` es obligatoria en cada evento declarado. Un evento que la
omite invalida la política completa, se ejecute o no ese evento. Una clave
desconocida, en cualquier nivel, también la invalida.

### Lanes condicionales: `docs_only_lane` y `marker_lanes`

Un lane acepta un skip solo cuando un job de la propia ejecución prueba un marker.
`equals` es `"true"` por defecto y `"false"` es el único otro valor admitido; el
marker se compara como cadena exacta, de modo que un valor distinto, una salida
ausente o otro evento dejan los skips sin declarar. Un lane nunca amplía
`accepted_skips`. `marker_lanes` admite varios lanes a la vez; dos lanes no
pueden compartir la firma `(event, gate_job, output)`, ni entre sí ni con
`docs_only_lane`.

Ejemplo, el carril `no_ui`: una PR que no toca UI puede saltar `e2e`, y solo `e2e`:

```json
{
  "required_jobs": ["pr-gate", "lint", "typecheck", "e2e"],
  "events": {
    "pull_request": {"accepted_skips": []},
    "push": {"accepted_skips": []}
  },
  "marker_lanes": [
    {
      "event": "pull_request",
      "gate_job": "pr-gate",
      "output": "ui_changed",
      "equals": "false",
      "skips": ["e2e"]
    }
  ]
}
```

Con `ui_changed` igual a `"true"`, el skip de `e2e` sale por exit 1. Añadir un lane
o un skip a un lane existente es una relajación frente a la base (tabla anterior)
y exige `relaxation`. `tests/test_readme_marker_lanes.py` ejecuta este ejemplo.

## Comparación contra la rama base (HR-15)

En `pull_request`, el paso obtiene la copia de la política de la rama base
(`git fetch` de la rama y `git show` de la ruta) y la pasa por `--base-policy`.
El gate compara la política del PR contra esa copia.

| Cambio respecto a la base | Clase | Exige `relaxation` |
|---|---|---|
| Job retirado de `required_jobs` | Relajación | Sí |
| Entrada añadida a `accepted_skips` de un evento | Relajación | Sí |
| Evento añadido | Relajación | Sí |
| Skip añadido a `docs_only_lane.skips`, o cambio de `event`, `gate_job` u `output` del lane | Relajación | Sí |
| Job añadido, skip retirado o evento retirado | Endurecimiento | No |

Una relajación solo pasa si la política del PR lleva el campo `relaxation`:

```json
"relaxation": {
  "reason": "full pasa a ejecutarse solo en la ejecución nocturna",
  "reference": "mi-org/mi-repo#123",
  "issue": 123
}
```

El campo debe ser distinto del que ya lleva la copia base. Un campo heredado
justificó un cambio anterior y no ampara una relajación nueva: sustituya
`reason` y `reference` por los de la relajación actual.

La ejecución se evalúa con la política del PR, no con la de la base. Añadir un
job exige tocar el workflow, el `needs` del agregador y la política en el mismo
PR (HR-29); la copia base no conocería el job nuevo.

La primera línea tras el veredicto declara si la comparación se hizo:

- `base comparison: performed, ...` lista cada relajación detectada y si queda
  amparada o rechazada.
- `base comparison: NOT performed ...` indica que el gate no recibió
  `--base-policy` y no comprobó ninguna relajación.

Si `--base-policy` apunta a un fichero ilegible, no válido o idéntico al de
`--policy`, el gate sale por exit 1.

### Límites de la comparación

- **El PR puede editar el workflow que invoca el gate.** En `pull_request`
  se ejecuta el workflow del propio PR: un PR que retire `--base-policy` del
  paso, o cambie la ruta de la política, elude la comparación. El gate no lo
  detecta. La línea `NOT performed` deja constancia en el log, y la revisión
  de los cambios en `.github/workflows/` queda como control humano, sin test
  que lo observe fallar (`documented-only`).
- **Copia base ausente.** El PR que instala el gate, o cualquier PR mientras
  la ruta de la política no exista en la rama base, no tiene copia que
  comparar: el paso lo detecta con `git ls-tree`, omite `--base-policy` y la
  salida declara `NOT performed`: revise esa política completa a mano. Hay un
  test que cubre este camino.
- **Otros eventos.** `push`, `schedule` y `workflow_dispatch` no tienen rama
  base: la comparación no se hace y la salida lo declara.
- **Copia base no válida.** Si la política de la rama base no cumple este
  contrato, todo PR sale en rojo hasta que la rama base se corrija.

## Paridad con el workflow (HR-29)

Tres conjuntos deben ser iguales: los jobs del workflow menos el agregador, el
`needs` del agregador y `required_jobs`. `check_parity.py` los compara contra
el workflow real del consumer. Añádalo como paso de un job del mismo workflow:

```yaml
    - name: Paridad de workflow, needs y política
      run: |
        python3 scripts/required-jobs/check_parity.py \
          --workflow .github/workflows/ci.yml \
          --aggregator required \
          --policy scripts/required-jobs/required-jobs.policy.json
```

| Argumento | Contenido |
|---|---|
| `--workflow` | Ruta del workflow que contiene el job agregador |
| `--aggregator` | Identificador del job agregador; otros jobs pueden declarar `needs` |
| `--policy` | Ruta de la política |

Sale por exit 0 si los tres conjuntos coinciden y por exit 1 ante cualquier
diferencia o duda. Detecta un job del workflow ausente de la política, un job
de la política sin definir o sin cablear en `needs`, y una entrada de `needs`
que la política desconoce.

El script lee el workflow con un subconjunto de YAML, sin dependencias:

- `jobs:` en estilo de bloque, con una clave simple por job.
- `needs` del agregador como escalar (`needs: a`), lista en línea
  (`needs: [a, b]`) o lista en bloque (`needs:` seguido de líneas `- a`).

Lo que queda fuera de ese subconjunto sale por exit 1 y nombra la
construcción: una expresión `${{ ... }}`, un ancla o un alias, una clave de
fusión `<<`, un job definido en línea, una lista en línea partida en varias
líneas, la ausencia del bloque `jobs:` o del job agregador.

Todo job distinto del agregador debe figurar en la política. El script no
admite todavía jobs excluidos a propósito del agregado.

## Veredicto

El veredicto es fail-closed. Salen por exit 1: entrada ilegible, política mal
formada, evento no declarado, job conocido ausente de `needs`, clave de
`needs` no cubierta por la política, skip no declarado, conclusión `failure`,
`cancelled`, `timed_out` o `startup_failure` y relajación sin amparo frente a
la copia base. La salida separa las causas raíz de los skips en cascada: un
skip descendiente nunca se lee como la causa.

## Códigos de salida

- **exit 0**: todo job requerido tuvo éxito o fue un skip declarado, y ninguna
  relajación quedó sin amparo.
- **exit 1**: el veredicto es rojo, o hubo una duda: entrada ilegible, política
  mal formada, un `--annotations` ilegible o mal formado, o un `--run-attempt`
  que es un entero pero no positivo.
- **exit 2**: error de uso de la línea de órdenes. Un `--run-attempt` que no es
  un entero (`abc`) lo corta argparse antes de que el gate juzgue; el job sale
  rojo igual, pero la causa no llega a imprimirse.

## Tests

| Fichero | Qué observa |
|---|---|
| `tests/test_required_jobs.py` | Veredictos del gate, formas de `needs` y validación de la política |
| `tests/test_aggregator_edges.py` | Bordes del agregador: AND exacto sobre los 12 jobs requeridos, `timed_out` y `startup_failure` como causa raíz, re-run en ambos sentidos y causa a nivel de check (#1208) |
| `tests/test_published_name.py` | `check_names`: el nombre publicado del check en la causa, y su validación (#1219) |
| `tests/test_build_annotations.py` | El generador del fichero de anotaciones y su degradación a «no annotation» (#1219) |
| `tests/test_base_comparison.py` | `--base-policy`: relajaciones, endurecimientos y campo `relaxation` |
| `tests/test_docs_only_lane.py` | Lane docs-only: skips aceptados solo con el marcador `"true"` en el evento declarado, combinado con `dependencies` y su relajación contra la base (#857) |
| `tests/test_marker_lane.py` | `marker_lanes`: lanes simultáneos, `equals` y su relajación contra la base (#863) |
| `tests/test_readme_marker_lanes.py` | El ejemplo del carril `no_ui` de este README, ejecutado contra el gate (#973) |
| `tests/test_readme_wiring.py` | Los pasos de este README, ejecutados con `bash` y `git` tal como están escritos |
| `tests/test_parity.py` | `check_parity.py`: paridad, cada diferencia y cada construcción no soportada |

```bash
for t in scripts/required-jobs/tests/test_*.py; do python3 "$t"; done
```

`test_readme_wiring.py` lee este README y la política de ejemplo: copie
ambos junto con el directorio `tests/`. Necesita `bash` y `git`; si faltan,
falla cuando la variable `CI` está definida y se salta en local.
