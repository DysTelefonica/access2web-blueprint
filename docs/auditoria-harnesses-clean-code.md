# Auditoría: harnesses estilo Uncle Bob / swarm-forge

> **Sentence que organiza este documento**: los harnesses correctos hacen innecesaria la revisión mecánica del código. Lo que queda para revisar es lo que importa: intención, dominio y edge cases no obvios.

## Qué es y qué no es

| Es | No es |
|---|---|
| Auditoría del setup de quality gates del MVP contra los principios de Clean Code y la orquestación por roles de swarm-forge. | Adopción literal de `unclebob/swarm-forge` (tmux + Babashka + worktrees, infraestructura pesada). |
| Mapa de cada principio Uncle Bob al gate que lo enforce, o a un gap explícito. | Réplica del workflow `six-pack` con sus seis roles ejecutándose en paralelo. |
| Identificación de lo que el harness puede reemplazar y de lo que sigue requiriendo revisión humana. | Lista exhaustiva de todas las prácticas de Uncle Bob. |
| Propuesta de Fase 2 con disparador, coste y prioridad por gate. | Roadmap de implementación ni épica nueva. |

## Roles swarm-forge mapeados al setup MVP

| Rol swarm-forge | Concern original | Gate o artefacto MVP | Estado |
|---|---|---|---|
| `specifier` | Convertir intención en specs Gherkin aceptadas. | `openspec/changes/<change>/specs.md` + revisión humana. | Pendiente del MVP Lanzadera. |
| `coder` | Implementar comportamiento con TDD + tests unitarios. | pytest en `pyproject.toml` con `--strict-markers --strict-config` + strict TDD flag. | Por activar en Día 6. |
| `cleaner` | CRAP, DRY, encapsulation, separation of concerns. Umbrales absolutos upstream: `CRAP ≤ 6` y split obligatorio por encima de 100 mutation sites. | `check_complexity.py` (`CC ≤ 15`), `check_crap.py` (`CRAP ≤ 6`, el mismo número de upstream) y `check_dry.py` (clones type-2 sobre AST). `check_module_size.py` sigue excluido. | Cubierto salvo el split por mutation sites, que depende de tener mutación. |
| `architect` | Module structure, dependency direction, property-test coverage. | `scripts/check_layers.py` con `ALLOWED_IMPORTS`, `PURE_LAYERS` y `_check_slice`. | Por implementar en Día 3. |
| `hardender` | Mutation hardening, CRAP y DRY verification, language mutation. | cosmic-ray en `docs/quality/cosmic-ray.toml`. | Diferido a Fase 2 (sin código ni suite madura). |
| `refactorer` | Behavior-preserving cleanup, coverage improvement, mutation-site scans. | `scripts/pytest_plugin/coverage_gate.py` con `CRITICAL_HELPERS = ["hash_password", "verify_password"]`. | Por implementar en Día 1. |
| `QA` | Specs ejecutables, verificación final, completion notification. | `ci.yml` + `security.yml` + `pr-size.yml` + `Makefile`. | Por implementar en Días 1-2. |

## Mapeo: principios Clean Code → gates del MVP

| Principio | ¿Mecánico? | Gate actual | Gap explícito |
|---|---|---|---|
| **Naming reveals intent** | Parcial. | Ruff N-series (`N802`, `N803`, `N806`, `N815`, `N818`). | Naming quality (intención, pronunciable, sin disinformation) sigue requiriendo revisión humana. |
| **Funciones pequeñas, una sola cosa** | Sí. | `check_complexity.py` con techo absoluto global `CC ≤ 15`. | El gate mide ramificación, no longitud: una función de 200 líneas con CC=10 sigue pasando. |
| **Comentarios: por qué, no qué** | Sí. | `check_docstring_coverage.py` con floor `73%`. | Excluido del MVP. Docstrings faltantes se acumulan sin ruido. |
| **Formato vertical y horizontal** | Sí. | ruff format + `line-length = 100`. | Por decidir si `ruff format` se enforce o solo `ruff check`. |
| **Error handling: excepciones, no códigos de retorno** | Parcial. | Ruff TRY/ERA rules. | No activas en MVP. `try/except: pass` pasa silencioso. |
| **Boundaries: third-party envuelto** | Sí. | `check_layers.py` con `PURE_LAYERS` que prohíbe `fastapi`, `sqlalchemy`, `jinja2`, `httpx`, `asyncpg`, `alembic` en `domain`, `ports`, `application`. | — |
| **SOLID SRP** | Parcial. | `check_complexity.py` + `check_module_size.py`. | Module size excluido del MVP. El techo de CC cubre ramificación, no cohesión: una clase con diez responsabilidades triviales pasa limpia. |
| **SOLID OCP** | No. | — | Revisión humana. El harness no detecta código que se modifica en lugar de extenderse. |
| **SOLID LSP** | No. | — | Revisión humana. Subtipos mal diseñados pasan los tests si los mocks coinciden. |
| **SOLID ISP** | Parcial. | Ruff `PLR0913` (parameter count) + `PLR0912` (branches). | No activas en MVP. Handlers de 26 parámetros pasan. |
| **SOLID DIP** | Sí. | `check_layers.py` con `ALLOWED_IMPORTS`. | — |
| **DRY** | Sí. | `check_dry.py` (clones type-1/type-2 sobre AST normalizado, 5+ sentencias, 0 tolerados). | Detecta copia-pega renombrado, que es el caso real. No detecta duplicación semántica con estructura distinta. |
| **YAGNI** | No. | — | Revisión humana. El harness no distingue código de un feature en uso de código especulativo. |
| **Encapsulation** | No. | — | Revisión humana. Atributos públicos innecesarios pasan `check_layers.py`. |
| **Separation of concerns** | Parcial. | `check_layers.py` slicing + `check_complexity.py`. | — |
| **TDD discipline** | Sí. | pytest + `--strict-markers --strict-config` + flag strict TDD. | El flag enforce order, no coverage del proceso red-green-refactor. |
| **Architecture boundaries** | Sí. | `check_layers.py` dependency direction. | — |
| **Tests como documentación de comportamiento** | Parcial. | pytest doctest collection. | Doctests no exigidos. Examples viven solo en los docs. |

## Gaps identificados que Uncle Bob enfatiza

| Gap | Riesgo si no se cubre | Disparador para activarlo |
|---|---|---|
| Cobertura sin mutación. | Un test sin aserciones cuenta como cobertura. CRAP hereda esa ceguera: confía en el número de cobertura. | Cuando ≥ 3 módulos superen el 70% de cobertura. Receta upstream: mutación diferencial contra manifest, un archivo por vez. |
| `check_module_size.py` excluido. | Módulos de 700 líneas se aceptan como baseline. | Cuando el módulo `lanzadera` supere 500 líneas. |
| `check_docstring_coverage.py` excluido. | Funciones públicas sin docstring se acumulan. | Cuando el módulo `lanzadera` alcance 80% de cobertura de tests. |
| Ruff format no enforce. | Diferencias de formato entre contribuciones generan diffs ruidosos. | Día 1 del MVP. |
| TRY/ERA rules no activas. | `try/except: pass` y comparaciones con strings laxas pasan. | Cuando el módulo `lanzadera.auth` tenga ≥ 3 meses de vida. |
| PLR0913/PLR0912 no activas. | Handlers con > 5 parámetros y > 12 branches pasan. | Cuando se introduzcan los primeros forms de la UI web. |
| cosmic-ray excluido. | Mutaciones que escapan a la suite no se detectan. | Cuando ≥ 3 módulos tengan cobertura > 70% y ≥ 5 tests por path real. |
| OCP, LSP, ISP sin gate mecánico. | Diseño orientado a herencia, no a composición, pasa los tests superficiales. | Revisión de PR obligatoria; no automatizable sin AI reviewer. |

## Lo que el harness no puede reemplazar

El setup MVP reduce la revisión mecánica al mínimo posible, pero deja cinco áreas donde la revisión humana sigue siendo insustituible:

- **Intención y dominio**. ¿La función hace lo que el nombre dice y lo que el caso de uso necesita? Ni ruff ni cosmic-ray responden esto.
- **Edge cases no obvios**. ¿El caso `None` se maneja donde el lenguaje no lo enforce? ¿El overflow se detecta? Cobertura al 100% no implica robustez.
- **OCP, LSP, YAGNI, encapsulation**. Cuatro pilares de Clean Code que ningún linter mide de forma fiable.
- **Naming que revela intención**. Las reglas N de ruff detectan convenciones; no detectan nombres vagos como `process_data` o `handle_stuff`.
- **Tests como spec ejecutable**. Que un test pase no significa que verifique lo correcto. La revisión del assert sigue siendo humana.

## Propuesta Fase 2 con disparador

| Gate | Disparador de activación | Coste | Prioridad |
|---|---|---|---|
| `ruff format --check` en CI. | Día 1 del MVP. | 30 min. | Alta. |
| Ruff TRY/ERA select activo. | Módulo `lanzadera.auth` con ≥ 3 meses de vida. | 1 h config + limpieza de violations. | Media. |
| Ruff PLR0913 + PLR0912 select activo. | Primeros forms de la UI web. | 2 h config + refactor de handlers. | Media. |
| Ajuste de `MIN_STATEMENTS` en `check_dry.py`. | Ruido medido en PRs reales, en cualquier dirección. | 30 min + revisión del BASELINE. | Baja. |
| `check_module_size.py` activado. | Módulo `lanzadera` > 500 líneas. | 1 día (script + BASELINE). | Baja. |
| `check_docstring_coverage.py` activado. | Módulo `lanzadera` con cobertura ≥ 80%. | 1 día (script + BASELINE). | Baja. |
| cosmic-ray mutation testing. | ≥ 3 módulos con cobertura > 70% y ≥ 5 tests por path real. | 1 día setup + wrapper con degenerate-run guard. | Diferible hasta Fase 2. |
| AI reviewer para OCP, LSP, ISP, naming. | Cualquier PR. | Dependiente del backend. | Explorar Fase 3. |

## Convenciones heredadas de Uncle Bob

Estas prácticas no se defienden con scripts; se defienden con revisión de PR y con la constitución del swarm.

| Convención | Lectura |
|---|---|
| Nombres que revelan intención. | No hay abreviaturas crípticas (`usr`, `mgr`, `proc`). No hay disinformation (`List` que no es lista). |
| Funciones a un solo nivel de abstracción. | Si una función mezcla `try/except` con SQL con validación de input, refactorizar. |
| Argumentos mínimos. | Más de tres parámetros sugiere objeto de parámetros. Handlers de formularios usan Pydantic models. |
| Sin efectos secundarios ocultos. | Una función llamada `check_x` no modifica estado. Si lo hace, su nombre miente. |
| Sin código muerto. | Comentarios que desactivan código, funciones no llamadas, ramas inalcanzables. Vulture activa Fase 2. |
| Sin retorno de null. | Usar `Optional[T]` explícito o Result types. `try/except` para errores, no para flujo. |
| Comentarios cuentan por qué. | Si un comentario reescribe el código en español, sobra. Si explica una decisión no obvia, falta. |
| Tests son la documentación. | El nombre del test describe el escenario: `test_user_locked_after_5_failed_attempts`. |
| Errores son contexto. | Una excepción trae el input que falló, la operación intentada y la causa raíz. No solo `raise ValueError`. |
| Boundaries wrapped. | `S3Client` no aparece en `domain/`. Aparece en `adapters/storage/s3.py` detrás de `StoragePort`. |

## Quick map inverso

| Si necesita | Abra primero | Y luego consulte |
|---|---|---|
| Entender cómo se mapea swarm-forge al setup. | §Roles swarm-forge. | §Mapeo principios → gates. |
| Saber qué falta contra Clean Code. | §Gaps identificados. | §Propuesta Fase 2. |
| Decidir si activar un gate ahora o esperar. | §Propuesta Fase 2 (columna Disparador). | §Lo que el harness no puede hacer. |
| Convencer a un reviewer de una práctica sin gate mecánico. | §Convenciones heredadas de Uncle Bob. | La lectura original de Clean Code + Clean Architecture. |
| Auditar un PR específico. | §Lo que el harness no puede reemplazar. | §Mapeo principios → gates. |

## Referencias

| Recurso | Ruta o enlace |
|---|---|
| Uncle Bob swarm-forge (rama `main` documental, ramas `two-pack`/`four-pack`/`six-pack` ejecutables). | `https://github.com/unclebob/swarm-forge` |
| Constitución shared articles de swarm-forge. | `swarmforge/constitution/articles/` en `main`. |
| Setup de quality gates del MVP. | [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) |
| Decisiones D8 (hexagonal global), D66-D82 (stack, monolito, Expand & Contract). | [`docs/08-decisiones-y-preguntas-abiertas.md`](08-decisiones-y-preguntas-abiertas.md) |
| Clean Code + Clean Architecture, Robert C. Martin. | Lectura de referencia para §Convenciones heredadas. |

## Lista de comprobación final

- [ ] Cada rol swarm-forge tiene al menos un gate o artefacto asociado en el MVP, o está marcado como Diferido con disparador explícito.
- [ ] Cada principio de Clean Code mapeado en §Mapeo tiene una respuesta explícita: gate activo, gate excluido con disparador, o revisión humana residual.
- [ ] La lista de §Lo que el harness no puede reemplazar se revisa en cada PR de plataforma.
- [ ] Las convenciones de §Convenciones heredadas se referencian desde `AGENTS.md` raíz cuando se introduzca la primera.
- [ ] El setup de Fase 2 se discute cuando se cumple el primer disparador, no antes.
