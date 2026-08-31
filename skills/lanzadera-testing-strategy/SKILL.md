---
name: lanzadera-testing-strategy
description: "Trigger: añadir test a tests/lanzadera/, revisar capa de un test, decidir mock vs fake, auditar clasificación de un test, debate sobre auth_bypass. Decide la categoría de test (1–5) según el árbol de decisión de docs/testing/testing-strategy.md y aplica las HR-N operativas a ese test."
license: MIT
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-08-31
  based_on: "skills/skill-style-guide/SKILL.md y skills/documentation-alan-style/SKILL.md"
---

[← Back to skills/README.md](../../skills/README.md)

## Activation Contract

Cargue esta skill **antes** de:

- Crear un archivo en `tests/lanzadera/` que no sea un fixture (categoría 1–5).
- Modificar un test existente para hacerlo más rápido o más lento.
- Revisar un PR que añade tests (verificar que la capa es la correcta).
- Decidir si mockear una dependencia o sustituirla por un fake.
- Responder "¿este test debería ser unit o integration?".
- Auditar la clasificación de un test existente (verificar que cumple HR-1..HR-8).

No la cargue para:

- Decisiones de arquitectura o modelo de datos (eso es `architecture-guardrails`).
- Diseño visual o tokens de marca (eso es `telefonica-brand-design`).
- Cambiar los `check_*.py` o añadir uno nuevo (eso es `deterministic-quality-harness`).
- Escribir un doc human-facing (eso es `documentation-alan-style`).

## Hard Rules

- **HR-1 — No importen** `tests/lanzadera/_fakes` ni sesiones de DB en `tests/lanzadera/domain/`. Si un test de dominio necesita repositorio, está testeando la capa equivocada — muévalo a `application/`.
- **HR-2 — Sustituyan** dependencias inyectadas por `FakeFixtures` (`tests/lanzadera/_fakes.py`) en `tests/lanzadera/application/`. `unittest.mock.MagicMock` queda prohibido en esta capa porque pierde la señal del contrato del puerto.
- **HR-3 — Prueben** sólo el `Protocol` del puerto en `tests/lanzadera/adapters/test_*_contract.py`. Si el test necesita datos sembrados, es Categoría 5 (Postgres e2e), no Categoría 3 (contract).
- **HR-4 — Prueben** la cadena HTTP con `TestClient` + `LanzaderaContainer` real con fakes en `tests/lanzadera/delivery/test_*_routes*.py`. Mockear el container está prohibido; se monta el árbol real y se sustituyen sólo los puertos de infraestructura.
- **HR-5 — Gaten** tests de Categoría 5 con `APAP_INTEGRATION_ENABLED=1`. Sin esa env var, el test ejecuta `pytest.skip(...)` con mensaje explícito. CI por defecto corre sin la env var.
- **HR-6 — Marquen** la fixture `auth_bypass` en tests de `delivery/` que toquen `admin_routes`. La fixture monkeypatchea `admin.require_global_admin` con un mensaje que cita el gate real que sustituirá (M02 / A01–A03).
- **HR-7 — Incluyan** assert verificable en cada test. `assert True` o `pass` sin acción detectable es una mentira; el gate de mutation semanal lo detecta, pero un humano no debería poder añadirlo. Si cree que necesita `pass`, escriba por qué en el docstring.
- **HR-8 — Ubiquen** tests bajo `tests/lanzadera/<capa>/test_*.py`. No cree tests en la raíz de `tests/lanzadera/` salvo que sean meta-tests del harness (CI workflow, gate wiring). Ejemplo válido: `tests/lanzadera/test_*_wiring.py`.

## Decision Gates

| Condición | Acción |
|---|---|
| Si la pieza es lógica pura sin I/O ni dependencias inyectadas | Categoría 1 (unit / domain). Carpeta: `domain/`. |
| Si la pieza depende de un `Protocol` declarado en `ports/` | Categoría 3 (contract) o Categoría 5 (Postgres e2e) según coste. Ver árbol completo en `docs/testing/testing-strategy.md`. |
| Si la pieza NO depende de un `Protocol` y es un use case | Categoría 2 (unit / use case). Carpeta: `application/`. |
| Si la pieza es la capa delivery (HTTP route, CLI) | Categoría 4 (integration HTTP). Carpeta: `delivery/`. |
| Si la pieza es un test del CI workflow o de un `check_*.py` | Meta-test. Carpeta: `tests/` raíz o `tests/lanzadera/test_*_wiring.py`. |
| Si una nueva capa se introduce sin Categoría documentada | Error de taxonomía. Pedir aclaración antes de escribir el test. No improvise la categoría. |

## Execution Steps

1. **Identifique** la pieza a testear y su capa (domain, application, adapters, delivery, di, exp, auth, migrations, notifications).
2. **Aplique** el árbol de decisión de `docs/testing/testing-strategy.md` §Cómo elegir el tipo correcto.
3. **Confirme** la Categoría (1–5) y ancle el archivo a la ruta esperada de la tabla §Las 5 categorías.
4. **Escriba** el test siguiendo las HR-N que apliquen a la Categoría.
5. **Verifique** que el path respeta HR-8 (`tests/lanzadera/<capa>/test_*.py`).
6. **Compruebe** que no hay asserts vacíos (HR-7).
7. **Ejecute** `python -m pytest tests/lanzadera/<capa>/test_<nombre>.py -v` localmente antes de push.
8. **Documente** la categoría en el docstring del test (sigue el patrón de `tests/lanzadera/adapters/test_user_repository_pg.py` líneas 4–19).

## Output Contract

Cuando use esta skill para clasificar un test, devuelva:

| Key | Type | Description |
|---|---|---|
| `categoria` | `1 \| 2 \| 3 \| 4 \| 5` | Categoría asignada por el árbol de decisión. |
| `capa_path` | string | Ruta esperada: `tests/lanzadera/<capa>/test_<nombre>.py`. |
| `fakes_imported` | string[] | Lista de fakes importados de `_fakes.py` (NO `mock.MagicMock`). Vacío si Categoría 1 o 3. |
| `requires_postgres` | bool | `true` sólo si Categoría 5. |
| `auth_bypass_needed` | bool | `true` sólo si Categoría 4 y el test toca `admin_routes`. |
| `gates_que_aplican` | string[] | Subset de `[layers, complexity, dry, mutation, coverage_gate]` que aplican a este test. |
| `hr_satisfied` | string[] | Lista de HR-N que el test cumple (ej. `["HR-1", "HR-2"]`). |
| `rationale` | string | Una frase citando el árbol de decisión y la evidencia (path del código). |
| `status` | `"classified" \| "blocked"` | `classified` si hay Categoría; `blocked` si el árbol no produce categoría. |
| `blocked_reason` | string | Sólo si `status == "blocked"`. Motivo por el cual no se pudo clasificar. |

## Anti-patterns

| Symptom | Fix |
|---|---|
| `unittest.mock.MagicMock()` en `application/test_*.py` | Sustituir por `FakeFixtures`. Ver HR-2. |
| Test de dominio que importa `_fakes` | Mover a `application/`. Ver HR-1. |
| `test_admin_route_creates_user` sin fixture `auth_bypass` | Añadir la fixture o etiquetar el test como `@pytest.mark.requires_auth`. Ver HR-6. |
| `assert response.status_code == 200` sin verificar el cuerpo | El test sólo prueba que el router no 500ea. Usar el container real y verificar la fila en el fake repo. |
| Test con `time.sleep(0.5)` | Reescribir con `freezegun` o `monkeypatch` del clock. El CI tiene 30 min budget; los sleeps silenciosos son drift. |
| Test marcado `@pytest.mark.skip` permanente con TODO | El skip debe ser condicional (`APAP_INTEGRATION_ENABLED`), no permanente. Los skips permanentes son deuda. |
| Fixture de DB compartida entre tests | Cada test crea su propio `FakeFixtures()` o su propio Postgres efímero. El orden de tests no debe importar. |
| Test Categoría 5 sin gating | Añadir skip condicional con `APAP_INTEGRATION_ENABLED`. Ver HR-5. |
| Path del test en raíz de `tests/lanzadera/` sin ser wiring | Mover a la carpeta de capa correspondiente. Ver HR-8. |

## Self-compliance

Esta skill se audita a sí misma contra el rubric de `skill-style-guide` HR-1..HR-13:

- [x] HR-1: frontmatter declara `name`, `description`, `license`, `metadata.author`, `metadata.version`, `metadata.last_verified`.
- [x] HR-2: body budget ≤ 1000 líneas. (Actual: ~250.)
- [x] HR-3: `description` arranca con `Trigger:`.
- [x] HR-4: `description` es single-language (castellano + identificadores técnicos en inglés, sin mezcla EN/ES en keywords).
- [x] HR-5: 8 HR-N numeradas y con verbo observable al inicio.
- [x] HR-6: Decision Gates es tabla `| Condición | Acción |`.
- [x] HR-7: Anti-patterns es tabla `| Symptom | Fix |`.
- [x] HR-8: Output Contract es tabla con todas las keys.
- [x] HR-9: la sección Output Contract lista todas las keys (10 keys).
- [x] HR-10: `metadata.last_verified` actualizado a 2026-08-31.
- [x] HR-11: secciones canónicas en orden: Activation → Hard Rules → Decision Gates → Execution Steps → Output Contract → Anti-patterns.
- [x] HR-12: la prosa sigue las reglas que prescribe (sin emojis decorativos en headings ni cuerpo).
- [x] HR-13: no duplica contenido de `CONTRIBUTING.md`, `AGENTS.md` ni `README.md`; referencia por path.

## Companion skills

- [`documentation-alan-style/SKILL.md`](../documentation-alan-style/SKILL.md) — patrón que esta skill aplica (estructura de sección, tono).
- [`skill-style-guide/SKILL.md`](../../../.agents/skills/skill-style-guide/SKILL.md) — rubric HR-1..HR-13 que esta skill cumple.
- [`architecture-guardrails/SKILL.md`](../architecture-guardrails/SKILL.md) — front-door a `docs/architecture.md` cuando el test cubre una decisión arquitectónica (D-n).

## Cross-references

- [`docs/testing/testing-strategy.md`](../../docs/testing/testing-strategy.md) — el doc fuente. Esta skill es su codificación operativa.
- [`docs/testing/epic.md`](../../docs/testing/epic.md) — la epic del proyecto testing-strategy (D-1..D-7).
- [`tests/lanzadera/_fakes.py`](../../tests/lanzadera/_fakes.py) — los fakes compartidos (HR-2).
- [`tests/lanzadera/_presence_fakes.py`](../../tests/lanzadera/_presence_fakes.py) — fakes especializados de presence.
- [`docs/calidad-de-codigo-y-ci.md`](../../docs/calidad-de-codigo-y-ci.md) §Hexagonal layer gate — la pureza de capa que HR-1 enforce.
- [`docs/architecture.md`](../../docs/architecture.md) — fuente de verdad única de las decisiones D-n referenciadas.
