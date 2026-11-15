# change-scope — qué suites puede romper una PR

`classify_changes.py` decide, a partir de la lista de ficheros de una PR, qué
suites de pruebas pueden encontrar en ella un error. Publica una marca por
suite, `<suite>_needed=true|false`, y el agregador `required`
(`assets/required-jobs`) acepta el salto de una suite **solo** cuando esa marca
vale `false` en la propia ejecución (#1440).

## Por qué listas positivas

El criterio es el de `Gentleman-Programming/gentle-ai`: cada suite cara declara
la lista de **código** que puede romperla (`on.pull_request.paths` en
positivo), no una lista de documentación excluida. Una ruta nueva que nadie
declaró no dispara suites; por eso la red de seguridad es una ejecución
completa programada, sin clasificador, que corre todas las suites cada noche.

Usted declare en cada lista lo que la suite ejecuta o carga: el código, sus
pruebas, sus dependencias y su configuración. Una PR de documentación o de
configuración de flota no toca ninguna lista y entra sin esperar a las suites
caras.

## Cuándo corre todo (fallo cerrado, HR-3)

| Situación | Resultado |
|---|---|
| La declaración no existe, no es JSON o no cumple el formato | exit 2, sin salida |
| La lista de ficheros no se puede leer | exit 2, sin salida |
| `--sealed-files` dado y no legible | exit 2, sin salida |
| La lista de ficheros está vacía | todas las suites `true` |
| Una ruta llega entrecomillada por `git` (no se puede casar) | todas las suites `true` |
| Una ruta casa con `always` | todas las suites `true` |
| Una ruta está bajo `.agents/skills/ci-pattern/assets/**`, es `.github/change-scope.json` o `.governance-manifest.json` | todas las suites `true` |
| Una ruta es un fichero sellado del manifiesto de gobierno | todas las suites `true` |

Con exit 2 el script no escribe nada en `$GITHUB_OUTPUT`. El paso **no debe
enmascarar** el código de salida (nada de `continue-on-error` ni de `|| true`):
el job `change-scope` queda rojo y, como es un job obligatorio, el agregador
también. Las suites se condicionan con `!= 'false'`, de modo que una marca
ausente las ejecuta.

## La declaración: `.github/change-scope.json`

```json
{
  "version": 1,
  "suites": {
    "phpunit": ["www/**", "tests/**", "composer.*", "phpunit.xml*"]
  },
  "always": [".github/workflows/**"]
}
```

| Clave | Obligatoria | Contenido |
|---|---|---|
| `version` | Sí | `1` |
| `suites` | Sí | Objeto `suite -> lista no vacía de patrones`. El nombre casa con `[A-Za-z0-9][A-Za-z0-9_-]*` y da la salida `<suite>_needed` |
| `always` | No | Lista no vacía de patrones que exigen todas las suites |

Los patrones siguen la semántica de `.gitattributes`, la misma que el
presupuesto de revisión (`assets/budget/vendored_budget.py`): sin `/` casan con
el nombre en cualquier nivel (`composer.*`); con `/` se anclan a la raíz;
`**` solo cruza directorios junto a una barra (`www/**`, `**/fixtures`), y
`docs/**.md` no cruza directorios. No se admiten la negación (`!`) ni la barra
final: escriba `dir/**`. Una clave desconocida invalida la declaración.

`change-scope.example.json` es un ejemplo completo.

## El job en el workflow

El job lee la declaración y el patrón de la **BASE** de la PR: una PR no puede
sacar sus propias rutas de una lista. Solo mide el árbol, así que dispara con
`opened`, `synchronize` y `reopened` (HR-73) y tiene su propia cola. Use
`--no-renames`: sin él, un fichero movido de `www/` a `docs/` solo aparece con
su ruta nueva y la suite de `www/` no correría.

```yaml
on:
  pull_request:
    types: [opened, synchronize, reopened]
  push:
    branches: [main]
  schedule:
    - cron: "0 3 * * *"   # red de seguridad: todas las suites, sin clasificador

jobs:
  change-scope:
    runs-on: [self-hosted, Linux, mi-consumer]
    timeout-minutes: 5
    concurrency:
      group: change-scope-${{ github.event.pull_request.number || github.run_id }}
      cancel-in-progress: ${{ github.event.action == 'synchronize' }}
    outputs:
      phpunit_needed: ${{ steps.scope.outputs.phpunit_needed }}
    steps:
      - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
        if: github.event_name == 'pull_request'
        with:
          ref: ${{ github.event.pull_request.head.sha }}
          path: pr-head
          fetch-depth: 0
          persist-credentials: false
      - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
        if: github.event_name == 'pull_request'
        with:
          ref: ${{ github.event.pull_request.base.sha }}
          path: base
          persist-credentials: false
      - id: scope
        if: github.event_name == 'pull_request'
        env:
          BASE_SHA: ${{ github.event.pull_request.base.sha }}
          HEAD_SHA: ${{ github.event.pull_request.head.sha }}
        run: |
          git -C pr-head diff --no-renames --name-only "$BASE_SHA...$HEAD_SHA" > changed.txt
          python3 -c 'import json,sys; print("\n".join(json.load(open(sys.argv[1]))["files"]))' \
            base/.governance-manifest.json > sealed.txt
          python3 base/.agents/skills/ci-pattern/assets/change-scope/classify_changes.py \
            --changed-files changed.txt --scope base/.github/change-scope.json \
            --sealed-files sealed.txt --github-output "$GITHUB_OUTPUT"

  phpunit:
    needs: change-scope
    if: needs.change-scope.outputs.phpunit_needed != 'false'
    runs-on: [self-hosted, Linux, mi-consumer]
    steps:
      - run: echo "la suite"
```

Fuera de `pull_request` el job no clasifica: las salidas quedan vacías y las
suites corren. Un consumer sin `.governance-manifest.json` quita la línea de
`sealed.txt` y la opción `--sealed-files`.

## La política del agregador

`change-scope` es un job obligatorio y es el `gate_job` del lane. Cada suite
que puede saltarse lleva su lane con `equals: "false"`; añadir el lane es una
relajación frente a la base y exige `relaxation` (`assets/required-jobs`,
HR-15 y HR-33):

```json
{
  "required_jobs": ["change-scope", "lint", "phpunit"],
  "events": {
    "pull_request": {"accepted_skips": []},
    "push": {"accepted_skips": []},
    "schedule": {"accepted_skips": []}
  },
  "marker_lanes": [
    {
      "event": "pull_request",
      "gate_job": "change-scope",
      "output": "phpunit_needed",
      "equals": "false",
      "skips": ["phpunit"]
    }
  ],
  "relaxation": {
    "reason": "phpunit solo corre cuando la PR toca su lista de código; la ejecución nocturna corre todo",
    "reference": "mi-org/mi-repo#123",
    "issue": 123
  }
}
```

Con `phpunit_needed` igual a `"true"`, o sin la salida, el salto de `phpunit`
deja el agregador en rojo. En `push` y `schedule` no hay lane: todas las
suites son obligatorias. `tests/test_classify_changes.py` ejecuta esta
política con `check_required_jobs.py`.

## Regla para los consumers

Un consumer que adopte el clasificador `MUST` ejecutar todas sus suites en un
workflow programado (`schedule`) sin clasificador. La lista positiva deja
fuera lo que nadie declaró; la ejecución nocturna es la que lo encuentra.
