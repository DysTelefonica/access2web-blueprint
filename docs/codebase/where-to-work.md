[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Dónde trabajar

Una fila por tarea canónica. Cada fila dice **qué leer**, **qué tocar** y **qué
probar**, para que un agente que llega nuevo sepa dónde está el trabajo sin
explorar el código.

La fila `local` es de este repo: su índice y las skills obligatorias.

| Tarea | Lee | Toca | Prueba |
|---|---|---|---|
| `form` | `docs/codebase/repository-map.md` (propiedad por paquete) y `docs/03-aplicaciones/lanzadera/epic.md` (la app que sirve la pantalla) | `app/src/modules/lanzadera/delivery/http/templates/base.html`, `app/src/modules/lanzadera/application/`, `app/src/modules/lanzadera/domain/` | `tests/lanzadera/application/` |
| `ci` | `docs/codebase/testing-and-ci.md` (qué corre en cada nivel) y `docs/gates.md` (manual de los gates) | `.github/workflows/ci.yml`, `scripts/check_complexity.py` | `scripts/local-preflight.sh` (el CI en local) |
| `schema` | `docs/codebase/maintainer-playbook.md` (la lista de la migración) y `docs/codebase/repository-map.md` | `app/migrations/versions/0001_core_schema.py` | `tests/lanzadera/migrations/` |
| `local` | `docs/agents-repo-guide.md` (índice de skills) y `AGENTS.md` (front-door de lectura, §F.2) | `skills/architecture-guardrails/SKILL.md`, `skills/documentation-alan-style/SKILL.md` | `scripts/local-preflight.sh` |

Reglas de la página:

- Las rutas son relativas a la raíz del repo y **existen**: lo comprueba el gate
  `where-to-work` del canon. Una ruta que se movió deja la página vieja, y eso
  es lo que el gate caza.
- Un marcador de posición (`<...>`) o un glob no cuentan como ruta.
- Las tres celdas de cada fila se rellenan. Si una receta no aplica en este
  repo, se dice por qué en la celda: una fila vacía es peor que una que
  explica su excepción.
- Se pueden añadir filas que no sean canónicas; el gate exige las cuatro y no
  prohíbe las demás.

## En este repo «migración» es una sola cosa

A diferencia de otros consumers de la fleet, aquí la migración de esquema es
**Alembic** y vive en `app/migrations/versions/`; no hay dos motores que se
confundan. La receta de la fila `schema` incluye la regla D82 (Expand &
Contract) y el test de bajada que pide la lista del mantenedor.
