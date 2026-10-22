# XCUT #785 — el job `mutation` lleva ocho semanas en rojo

## Outcome

El job `mutation` del workflow `ci` vuelve a medir de verdad: su paso `run mutants` deja de morir con exit 127 y ejecuta una sesión real de cosmic-ray; el ratchet `scripts/check_mutation.py` decide. Diagnóstico en `~/.local/share/team-skills-auditor/p2-785.md` (host del auditor).

## Branch

`fix/785-mutation-cosmic-ray` desde `main @ a9f37c0`. Worktree `~/repos/access2web-blueprint-worktrees/fix/fase-785-mutation`.

## Causa raíz (medida, no inferida)

- `cosmic-ray` no está en las extras `dev` de `app/pyproject.toml`, ni en `app/uv.lock`, ni en `app/requirements-dev.lock`. El pin (`cosmic-ray==8.7.0`) vive solo en el asset `skills/deterministic-quality-harness/assets/pyproject.fragment.toml:18`; el fragmento nunca se consolidó.
- `docs/quality/cosmic-ray.toml` no existe y nunca existió; el `Makefile` apuntaba a otro nombre inexistente (`mutation.toml`).
- Resultado: `ci.yml:409` ejecuta `cosmic-ray init ...` sobre un binario ausente → `cosmic-ray: command not found` + exit 127 en 15 s.

## Presupuesto

`app/pyproject.toml` (1 línea), `app/uv.lock` y `app/requirements-dev.lock` regenerados, `docs/quality/cosmic-ray.toml` (nuevo, ~10 líneas), `Makefile` (3 líneas), `tests/test_ci_workflow.py` (1 test). Los locks pueden dominar el diff; vigilar las 400 líneas.

## Estrategia

1. RED: test en `tests/test_ci_workflow.py` que ata lo que el job invoca con lo que su install provee y comprueba que las rutas de config existen.
2. `cosmic-ray==8.7.0` en las extras `dev` + `uv lock` + `uv export --frozen --no-emit-project --extra dev`.
3. `docs/quality/cosmic-ray.toml` desde el asset de la skill; `Makefile` a la misma ruta.
4. GREEN local, luego `workflow_dispatch` sobre la rama y resultado real del job.

## Verificación local pre-PR

```bash
pytest -c app/pyproject.toml --rootdir=app tests/test_ci_workflow.py
ruff check --config app/pyproject.toml .
```

## Cierre

PR con `Closes #785`, `type:bug`, anti-slop N/23 y tabla de trazabilidad. El job `mutation` tiene que pasar en verde de verdad (workflow_dispatch), y ese resultado va al PR y al aviso.
