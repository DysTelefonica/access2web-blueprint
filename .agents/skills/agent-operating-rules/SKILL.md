---
name: agent-operating-rules
description: "Trigger: trabajar como agente en oracle o rosa, crear un checkout de trabajo, scratch temporal, borrar un directorio, lanzar tests en local, vigilar la CI, credencial en un comando, antes de commitear, MiniMax agotado. Reglas de trabajo del dueño para los agentes en sus máquinas compartidas."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-11-01
  scope: ['universal', 'ops']
  auto_invoke: ['working as an agent on a shared host', 'creating a work checkout', 'cleaning up scratch files']
  tiers: ['universal', 'ops']
---

# Reglas de trabajo de los agentes en las máquinas del dueño

Estas reglas las fijó el dueño para los agentes que trabajan en sus máquinas compartidas (oracle, rosa): workers de Pi, sesiones delegadas y coordinadores. Cada una responde a un incidente real: un `rm -rf` con la variable vacía, un token impreso en un error o una máquina sin memoria por cinco suites en paralelo.

Esta skill no repite lo que ya está en otras:

| Tema | Skill que manda |
| --- | --- |
| Rol colaborador o coordinador, quién fusiona, wrappers de identidad | `github-roles` |
| Actualizar la rama, cadenas, etiquetas, `status:approved`, 400 líneas, conservar la rama remota | `governed-delivery` |
| Comprobación anti-slop antes de cualquier cambio | `ai-slop-discipline` |

## Activation

Cargue esta skill cuando:

- trabaje como agente en oracle, en rosa o en otra máquina que comparte con otros agentes;
- vaya a crear un checkout de trabajo, un directorio temporal o un venv;
- vaya a borrar cualquier cosa del disco;
- vaya a lanzar tests o a vigilar la CI desde la máquina;
- un comando necesite una credencial;
- vaya a hacer un commit;
- trabaje como worker y MiniMax se haya agotado.

**No la cargue** para la estación Windows del dueño: allí rige el diseño de worktrees de `worktree-reorg-per-project` (`<proyecto>-worktrees\<nombre>`).

## Hard Rules

- **HR-1** — Nunca ejecute `rm -rf` ni un `rm` sobre una ruta construida con variables. Un `rm -rf "$VAR/cadete"` con `$VAR` vacío borra `$HOME/cadete`. Los generados que ignora `.gitignore` (`__pycache__`, `*.egg-info`, cachés, `node_modules`, informes) tampoco se borran.
- **HR-2** — Cree el trabajo temporal con `d=$(mktemp -d)` y déjelo donde está. Si hace falta rehacerlo, cree otro con `mktemp -d`. Al terminar, liste las rutas de scratch y de worktrees que ha creado, para que el coordinador decida qué limpiar.
- **HR-3** — Ponga el venv y los artefactos intermedios (cuerpos de PR, borradores, salidas de gate) en ese directorio privado. No use rutas predecibles en `/tmp`, porque otro agente puede pisarlas. Tampoco cree el venv dentro del repositorio: un `uv sync` a secas crea `./.venv` en el árbol de trabajo.
- **HR-4** — Cree los checkouts de trabajo solo con `~/bin/merge-tools/agent-wt.sh`. Prohibido el `git clone` a mano. Los checkouts principales (`~/repos/cadete`, `~/repos/team-skills`) no se tocan.
- **HR-5** — Nunca ponga un secreto en la línea de comandos: ni `$(gh auth token)` como argumento, ni un `gho_…` o un `ghs_…` literal, ni un `--token <valor>`. La credencial va por variable de entorno (`GH_TOKEN`) y no se imprime. Si la salida de un error puede contener un secreto, muestre solo las líneas útiles en lugar del error entero.
- **HR-6** — Use solo la identidad de agente que trae el shell (`dys-agentes`, a través del wrapper). Nunca ejecute `gh auth token -u ardelperal`, ni siquiera para leer, ni sustituya `GH_TOKEN` o `GITHUB_TOKEN` por el token del dueño. Si una operación falla por permisos con la identidad de agente, pare y avise al coordinador, que la hace él.
- **HR-7** — Si el cambio toca `.github/workflows/`, no lo empuje: la App no puede publicar workflows. Haga el commit en local y entregue al coordinador la rama, el SHA y la ruta del worktree.
- **HR-8** — En local, ejecute solo los tests del cambio y con `nice -n 15`. No lance la suite completa (`testing/run-unit-suites.sh`, la `propagate suite`): esa la corre la CI. Si hay una suite larga suya en marcha, párela y deje que decida la CI.
- **HR-9** — Deje un mínimo de 120 s entre dos consultas a `gh` cuando vigile una PR o una ejecución. Nada de bucles con `sleep` cortos.
- **HR-10** — Antes de cada commit, ejecute `ruff format --check` y `ruff check` con la versión de ruff que fija el lock del repositorio, no con la del `PATH`.
- **HR-11** — Cierre una PR con `gh pr close`, nunca con `gh issue close`.
- **HR-12** — Como worker, trabaje solo con MiniMax. Nunca pase a Claude (Haiku ni ningún otro modelo): la cuenta personal de Claude del dueño, con la mitad de capacidad que la de empresa, queda para los coordinadores.

## Decision Gates

| Condition | Action |
| --- | --- |
| Necesita un directorio temporal | `d=$(mktemp -d)`; se queda ahí y su ruta va en el informe final. |
| Quiere «limpiar» un directorio | No lo borre. Anote la ruta en el informe y deje que el coordinador decida. |
| Necesita trabajar en otro repositorio | `agent-wt.sh add <owner/repo> <su-agente> <slug>` (paso 1 de Execution Steps). |
| La rama del worktree está atrasada respecto de `main` | `git merge origin/main`, como manda `governed-delivery`. |
| Un comando o un gate necesita el token | `export GH_TOKEN=...` en una línea propia y el comando siguiente sin el valor. |
| La API devuelve 403 «Resource not accessible by integration» | Pare y avise al coordinador con el endpoint. No cambie de identidad. |
| Quiere seguir una ejecución de la CI y `gh run watch` o `gh pr checks` dan 403 | Lea `GET /repos/{r}/commits/{sha}/check-runs`, que la App sí puede leer. |
| El diff toca `.github/workflows/` | Commit en local; entregue la rama, el SHA y el worktree (HR-7). |
| Quiere comprobar el cambio en local | Solo los tests del cambio, con `nice -n 15` (HR-8). |
| Hay que matar un proceso suyo | Recorra `/proc` y mate solo los procesos cuyo `cwd` cae en su worktree, empezando por el padre. `pkill -f` coincide con su propio comando. |
| Va a hacer un commit | Primero `ruff format --check` y `ruff check` (HR-10). |
| MiniMax se ha agotado | Aparque el trabajo al terminar el paso en curso (paso 6 de Execution Steps). |

## Execution Steps

1. **Checkout.** Cree el worktree con el script; devuelve la ruta en `~/agent-wt/<repo>/<agente>-<slug>`:

   ```bash
   W=$(~/bin/merge-tools/agent-wt.sh add DysTelefonica/cadete <su-agente> <slug> [--branch <rama>] [--ref <ref>])
   ```

   El `<su-agente>` es el nombre que le dio el coordinador, no su rol. Sin `--branch`, el worktree queda en detached HEAD sobre `origin/main`: cree la rama con `git switch -c` cuando ya conozca el número de issue.

   Para ponerlo al día: `agent-wt.sh fetch <owner/repo>` (o `git -C "$W" fetch origin`) y después `git merge origin/main`. El `main` local del espejo compartido suele estar atrasado, así que el diff real se mide contra `origin/main`.

   Para retirarlo: `agent-wt.sh remove "$W"`. Rechaza la retirada si quedan commits sin empujar; `--discard` solo vale cuando el trabajo sobra de verdad. Para listar: `agent-wt.sh list [owner/repo]`.

   Si un worktree se creó con un nombre equivocado, déjelo donde está: renombrarlo rompe los metadatos de `git worktree`.

2. **Scratch y venv.**

   ```bash
   VDIR=$(mktemp -d)
   UV_PROJECT_ENVIRONMENT="$VDIR/venv" uv sync --frozen --extra dev
   ```

   En cadete, `www/vendor/` está versionado: un `composer install` en el worktree modifica ficheros versionados. Restáurelos con `git checkout -- www/vendor` antes del commit. No use `git clean`.

3. **Credenciales.** Si un gate necesita el token, expórtelo en una línea propia y no lo repita:

   ```bash
   export GH_TOKEN="$(gh auth token)"   # la expansión no aparece en el comando siguiente
   ```

   Ni `--token <valor>`, ni una tubería que vuelque el valor a stdout.

4. **Tests y CI.** Ejecute los tests del cambio con `nice -n 15`. Para vigilar la CI, consulte cada 120 s como mínimo y use la API de check-runs si la App no puede leer `/actions/runs`.

5. **Commit.** Ejecute `ruff format --check` y `ruff check` con la versión del lock. Los commits son convencionales (`feat`, `fix`, `docs`, `chore`, `ci`, `test`, `perf`, `refactor`, `revert`). `style` no está en esa lista: un reformateo puro va como `chore(...)`. El mensaje del commit no lleva atribución de IA (`Co-Authored-By:`, `Generated-by:` ni `Generated with`); el cuerpo de la PR sí puede llevar la marca de Claude Code.

6. **Aparcar y retomar (workers).** Si MiniMax se agota, termine el paso en curso y guarde su posición: tarea, rama, worktree y siguiente paso. Su coordinador avisa al dueño de que el worker queda aparcado. En la siguiente ventana de MiniMax (10, 15, 20, 01 y 06 UTC; avisa el cron `minimax-cycle`), el coordinador lo retoma con `/new` y un brief autocontenido que parte de esa posición, y vuelve a avisar al dueño.

7. **Entrega.** Informe con el Output Contract.

## Output Contract

| Key | Type | Description |
| --- | --- | --- |
| `status` | `"done" \| "handoff" \| "parked" \| "blocked"` | `handoff` cuando el coordinador tiene que empujar o hacer algo con permisos que el agente no tiene; `parked` cuando MiniMax se agotó. |
| `branch` | string | Rama de trabajo. |
| `sha` | string | Último commit, empujado o local. |
| `worktree` | string | Ruta absoluta del worktree. |
| `scratch_paths` | string[] | Directorios de `mktemp -d` creados, para que el coordinador los limpie. |
| `tests_run` | string[] | Tests ejecutados en local, con su resultado. |
| `needs_coordinator` | string \| null | Lo que el coordinador tiene que hacer (push de workflows, endpoint con 403) o `null`. |
| `next_step` | string | Siguiente paso. Con `parked`, es la posición desde la que se retoma. |

## Anti-patterns

| Symptom | Fix |
| --- | --- |
| `rm -rf "$DIR"` para rehacer un directorio | Otro `mktemp -d`; el viejo se queda y su ruta va en el informe. |
| `git clone --filter=blob:none` en `/tmp` para una tarea | `agent-wt.sh add`: el espejo compartido evita clones repetidos y `remove` tiene guardas. |
| `ci-pattern pr contract --token "$(gh auth token)"` | La CLI no admite `--token` y su error imprime el valor. `export GH_TOKEN` y el comando sin el token. |
| Un 403 en `/branches/{b}/protection` y se cambia al token del dueño | La publicación queda firmada por el dueño. Pare y avise al coordinador. |
| `repos/{r}.permissions.push` es `false` y se concluye que la App no puede abrir PR | Ese campo refleja permisos de usuario, no de la instalación. Abra la PR con el wrapper. |
| `testing/run-unit-suites.sh` completo en oracle | Solo los tests del cambio con `nice -n 15`; la suite completa la corre la CI. |
| `while :; do gh pr checks; sleep 10; done` | Una consulta cada 120 s como mínimo, o la API de check-runs. |
| `pkill -f <patrón>` para parar sus tests | El patrón coincide con el propio `pkill`. Use `/proc/<pid>/cwd` para matar solo lo suyo. |
| Commit `style(ci): ruff format` | `chore(...)`: el gate de commits no admite `style`. |
| El ruff del `PATH` dice «limpio» y la CI falla por formato | Use la versión que fija el lock del repositorio. |
| `gh issue close <N>` sobre una PR | `gh pr close <N>`. |
| Worker sin MiniMax que sigue con Haiku | Aparcarlo al terminar el paso y retomarlo en la siguiente ventana. |

## References

- `github-roles`: roles, identidades y wrappers `~/bin/app-gh/{agentes,orquestador}`.
- `governed-delivery`: actualización de rama, cadenas, etiquetas y merge.
- `~/bin/merge-tools/agent-wt.sh`, en las máquinas del dueño. Por defecto, `AGENT_WT_MIRRORS=$HOME/.cache/agent-mirrors` y `AGENT_WT_ROOT=$HOME/agent-wt`. Usa `flock` sobre el espejo, así que varios agentes pueden pedir el mismo repo a la vez.
- Team-skills#1461: la issue que reúne estas reglas desde Engram.
