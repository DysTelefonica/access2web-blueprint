# Dónde trabajar

Una fila por tarea canónica. Cada fila dice **qué leer**, **qué tocar** y **qué
probar**, para que un agente que llega nuevo sepa dónde está el trabajo sin
explorar el código.

La fila `local` es de este repo: su índice y las skills obligatorias.

| Tarea | Lee | Toca | Prueba |
|---|---|---|---|
| `form` | `docs/codebase/ejemplo-modulo.md` | `app/modules/ejemplo` | `tests/test_ejemplo.py` |
| `ci` | `docs/codebase/ci-cd.md` | `.github/workflows/ci.yml` | `make verify` |
| `schema` | `docs/codebase/esquema.md` | `app/core/migration/sql/001_init.sql` | `tests/test_esquema.py` |
| `local` | `docs/agents-repo-guide.md` | `skills/ejemplo-local/SKILL.md` | `make lint` |

Reglas de la página:

- Las rutas son relativas a la raíz del repo y **existen**: lo comprueba el gate
  (`assets/where-to-work/check_where_to_work.py`). Una ruta que se movió deja la
  página vieja, y eso es lo que el gate caza.
- Un marcador de posición (`<...>`) o un glob no cuentan como ruta.
- Las tres celdas de cada fila se rellenan. Si una receta no aplica en este
  repo, se dice por qué en la celda: una fila vacía es peor que una que
  explica su excepción.
