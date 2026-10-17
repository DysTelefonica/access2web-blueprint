# Registro de decisiones

Una decisión por fichero `AAAA-MM-DD-<slug>.md`. Este índice lista exactamente las decisiones del directorio; lo comprueba `tests/test_decisions_registry.py`.

## Formato

Tras el título, una lista con estas claves:

| Clave | Obligatoria | Valor |
|---|---|---|
| `estado` | Sí | `adoptada`, `rechazada`, `sustituida` o `aplazada` |
| `fecha` | Sí | Fecha ISO `AAAA-MM-DD` |
| `decide` | Sí | Rol o persona que decide |
| `motivo` | Sí | Por qué; no puede estar vacío |
| `sustituida-por` | Si `sustituida` | Slugs separados por comas de decisiones existentes; no a sí misma ni en ciclo |
| `revision` | Si `aplazada` | Condición o fecha de revisión |
| `origen` | No | De dónde procede (topic_key, issue o PR) |

## Decisiones

| Decisión | Estado | Fecha | Decide |
|---|---|---|---|
| [D36](2026-08-08-d36-preservar-hashes-heredados.md) | sustituida | 2026-08-08 | usuario (dueño del producto) |
| [D37](2026-08-08-d37-rehash-transparente.md) | sustituida | 2026-08-08 | usuario (dueño del producto) |
| [D88](2026-08-08-d88-auth-argon2id.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [D89](2026-08-08-d89-migracion-sin-hash-heredado.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [PRs siempre contra main](2026-10-15-prs-siempre-contra-main.md) | adoptada | 2026-10-15 | dueño del repositorio |
| [El registro de decisiones debe ser completo, sin baseline](2026-10-17-registro-de-decisiones-completo.md) | adoptada | 2026-10-17 | dueño del repositorio |

Las D-n anteriores a este registro viven en `../08-decisiones-y-preguntas-abiertas.md` y se migran aquí en la issue #985; hasta entonces, las cuatro primeras filas de arriba (D36, D37, D88, D89) son la única cadena de sustitución migrada.
