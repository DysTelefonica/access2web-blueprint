# Fase 2 — Aislamiento del entorno y toolchain de medición

> **Qué es.** La declaración de la fase 2 del contrato de adopción: qué árbol
> escanea el gate de aislamiento (HR-28), qué variables puede leer la suite, con
> qué intérprete y dependencias se miden las cifras, y cuántos actores trabajan
> a la vez sobre este repositorio.

> **Qué no es.** Un cambio de la suite, del plugin de cobertura ni de los
> umbrales: el piso de cobertura (`--cov-fail-under=69`) y los cuatro objetivos
> críticos quedan como están. Tampoco declara `governance` ni adelanta fases.

## La raíz que se escanea: `tests`

El gate escanea un árbol de suite (`*.sh`, `*.py`) y exige que no lea el entorno
local. El árbol declarado es **`tests`**, que es donde vive la suite para
`pytest` (`app/pyproject.toml`: `testpaths = ["../tests"]`). Las otras dos
candidatas se midieron antes de decidir, con la lista vacía de variables:

| `--root` | Resultado medido | Por qué no |
|---|---|---|
| `tests` | 4 hallazgos, todos el interruptor de integración | Es el sujeto: el árbol que la suite corre |
| `app` | 7 hallazgos de configuración de la aplicación (`DATABASE_URL` ×3, `PLATFORM_SECRET_KEY`, `GLOBAL_ADMIN_EMAILS`, `FERNET_KEY`, `SECRET_KEY`) | Mide la aplicación, no la suite: esas variables son el contrato de despliegue de `app/src` y `app/migrations`, y declararlas en la allowlist sería meter configuración de la aplicación en el contrato de la suite, justo lo que el criterio 2 prohíbe |
| `.` | 80 hallazgos y un hallazgo de dominio con 1.293 ficheros fuera del dominio del gate (`.bas` 544, `.cls` 723, `.mjs` 1, `.ps1` 16, `.ts` 9) | Mide el repositorio entero, no una suite: la mayoría del árbol es Access/VBA y herramientas, ajenas a este control |

## La allowlist: una sola variable

`.github/env-isolation-allowlist.json` declara **`APAP_INTEGRATION_ENABLED`** y
nada más. Con la lista vacía el gate nombra exactamente 4 lecturas de esa misma
variable (el interruptor con el que los casos que tocan base de datos real se
saltan: `tests/lanzadera/migrations/*.py` y
`tests/lanzadera/adapters/test_user_repository_pg.py`). Ninguna otra variable
aparece. La suite no lee configuración de la aplicación.

## Toolchain de medición, pineada

| Pieza | Valor | Dónde |
|---|---|---|
| Intérprete del runner | `3.12.11` | `.github/workflows/ci.yml` (`env.PYTHON_VERSION`), usado por `actions/setup-python` |
| Intérprete del worktree | `3.12.11` | `app/.python-version` (era `3.12`, nivel menor; se alinea en este PR porque el propio repositorio trata un cambio de parche como un veredicto distinto) |
| Dependencias | `python -m pip install --editable "app[dev]"` sobre `app/pyproject.toml`, con `app/uv.lock` versionado | Paso `install` del job `unit` de `ci.yml` |
| Medición local | `make test` (`pytest -c app/pyproject.toml --rootdir=app -m "not integration" --cov --cov-fail-under=69`) | `Makefile` |

Premisa declarada: una cifra medida en local solo es comparable con el runner si
el intérprete es el mismo parche (3.12.11) y las dependencias salen del mismo
`app/pyproject.toml`; cualquier otra combinación mide otra cosa.

## Actores concurrentes y sus worktrees

Un worktree por actor, ninguno compartido. La convención del repositorio es
`access2web-blueprint-worktrees/<tipo>/<slug>`:

| Actor | Worktree | Rol |
|---|---|---|
| Agente de implementación | `access2web-blueprint-worktrees/chore/fase-2-aislamiento` | Escribe esta fase |
| Agente de implementación | `access2web-blueprint-worktrees/docs/fase-1-propagacion` y `…/chore/fase-0-inventario` | Fases 0 y 1 ya cerradas; se conservan sin escrituras |
| Auditor | copias del PR | Revisa; no escribe en ningún worktree del repositorio |

## Premisa del entorno: el `.env`

No hay `.env` en el árbol y `.gitignore` no lo ignora, así que un `.env` local
quedaría visible como fichero sin versionar en lugar de quedar oculto. El gate
comprueba que la suite no lo lee: cero referencias a `.env`/`dotenv` bajo
`tests`. Si alguna vez aparece una, el hallazgo nombra fichero y línea.

## Validación ejecutada

```text
$ python3 <skill>/assets/env-isolation/check_env_isolation.py \
    --root tests --allowlist .github/env-isolation-allowlist.json
ENV-ISOLATION OK: 173 suite file(s) scanned — no .env reads, no undeclared environment variables.
EXIT=0

$ python3 <skill>/assets/bin/ci-pattern adoption check --phase 2 .
adoption check: fases 0..2 en verde sobre .
EXIT=0

# con la allowlist vacía, el gate no pasa: nombra las 4 lecturas de la suite
HR-28 FAIL: 4 finding(s) — …
EXIT=1

# al crear la allowlist, el gate de la fase 0 la nombró como artefacto de
# gobierno sin entrada de inventario (HR-46) hasta registrarla
hallazgo: fase 0: artefacto de gobierno sin entrada de inventario: .github/env-isolation-allowlist.json
EXIT=1
```

## Ficheros y comandos de esta fase

```bash
# lo declarado
.github/ci-pattern-adoption.json      # phases."2".env_isolation {root, allowlist} + entrada de inventario
.github/env-isolation-allowlist.json  # el contrato de variables de la suite
app/.python-version                   # 3.12.11, el mismo parche que el runner

# las medidas (solo lectura)
python3 <skill>/assets/env-isolation/check_env_isolation.py --root tests --allowlist .github/env-isolation-allowlist.json
python3 <skill>/assets/bin/ci-pattern adoption check --phase 2 .
git worktree list
```
