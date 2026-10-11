# required-jobs — agregador fail-closed de jobs requeridos

Primer asset ejecutable de `ci-pattern` (issue DysTelefonica/team-skills#132).
Sustituye al `check_required_jobs.py` del consumer de origen: la lógica es la
misma, pero ningún nombre de job, evento o skip vive en el código — todo eso
es política (`required-jobs.policy.example.json` como molde de ejemplo).

## Contenido

| Fichero | Papel |
|---|---|
| `check_required_jobs.py` | Gate: veredicto fail-closed sobre el objeto `needs` serializado |
| `check_parity.py` | Comprobación de paridad entre el workflow, el `needs` del agregador y la política |
| `required-jobs.policy.example.json` | Política de ejemplo: jobs conocidos y skips aceptados por evento |
| `tests/` | Suite ejecutable (solo stdlib): veredictos, fail-closed y cableado documentado |

## Destino en el consumer

Copie el directorio a `scripts/required-jobs/` del repo consumer (o la ruta
equivalente) y versione allí su propia política; la de ejemplo es un molde,
no una configuración. No hay nada más que instalar: el gate es stdlib-only.

## Cableado del job agregador

```yaml
required:
  needs: [lint, compile, smoke, unit, full]   # la misma lista que la política declara
  if: always()
  runs-on: ubuntu-latest
  steps:
    - uses: actions/checkout@v4
    - name: Veredicto del agregador
      env:
        NEEDS_JSON: ${{ toJSON(needs) }}
        EVENT_NAME: ${{ github.event_name }}
        BASE_REF: ${{ github.base_ref }}
        BASE_SHA: ${{ github.event.pull_request.base.sha }}
      run: |
        policy=scripts/required-jobs/required-jobs.policy.json
        set -- --policy "$policy" --event "$EVENT_NAME"
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

El objeto `needs`, el nombre del evento y la rama base llegan al script por
variables de entorno (`env:`), nunca interpolados dentro de `run:`. Una
expresión `${{ ... }}` escrita en el script se sustituye antes de que el shell
lo lea: una comilla en una salida de job rompería el paso o inyectaría órdenes.

`toJSON(needs)` entrega, por cada job, un objeto `{"result": ..., "outputs": ...}`.
El gate lee la conclusión del campo `result`; también admite la forma plana
`{"job": "success"}`. Cualquier otra forma sale por exit 1. La entrada llega
por stdin o por `--needs-file`.

Marque `required` como check requerido en la protección de rama.

## Contrato de la política

| Clave | Obligatoria | Contenido |
|---|---|---|
| `required_jobs` | Sí | Lista no vacía de nombres de job, sin duplicados |
| `events` | Sí | Un objeto por evento que el workflow atiende |
| `events.<evento>.accepted_skips` | Sí | Jobs de `required_jobs` que ese evento puede saltar; `[]` si ninguno |
| `relaxation` | Solo al relajar | Objeto con `reason` y `reference`, ambos cadenas no vacías |

`accepted_skips` es obligatoria en cada evento declarado. Un evento que la
omite invalida la política completa, se ejecute o no ese evento. Una clave
desconocida, en cualquier nivel, también la invalida.

## Comparación contra la rama base (HR-15)

En `pull_request`, el paso obtiene la copia de la política de la rama base
(`git fetch` de la rama y `git show` de la ruta) y la pasa por `--base-policy`.
El gate compara la política del PR contra esa copia.

| Cambio respecto a la base | Clase | Exige `relaxation` |
|---|---|---|
| Job retirado de `required_jobs` | Relajación | Sí |
| Entrada añadida a `accepted_skips` de un evento | Relajación | Sí |
| Evento añadido | Relajación | Sí |
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
`needs` no cubierta por la política, skip no declarado, conclusión `failure`
o `cancelled` y relajación sin amparo frente a la copia base. La salida separa
las causas raíz (`failure`, `cancelled`) de los skips en cascada: un skip
descendente nunca se lee como la causa.

## Tests

| Fichero | Qué observa |
|---|---|
| `tests/test_required_jobs.py` | Veredictos del gate, formas de `needs` y validación de la política |
| `tests/test_base_comparison.py` | `--base-policy`: relajaciones, endurecimientos y campo `relaxation` |
| `tests/test_readme_wiring.py` | Los pasos de este README, ejecutados con `bash` y `git` tal como están escritos |
| `tests/test_parity.py` | `check_parity.py`: paridad, cada diferencia y cada construcción no soportada |

```bash
for t in scripts/required-jobs/tests/test_*.py; do python3 "$t"; done
```

`test_readme_wiring.py` lee este README y la política de ejemplo: copie
ambos junto con el directorio `tests/`. Necesita `bash` y `git`; si faltan,
falla cuando la variable `CI` está definida y se salta en local.
