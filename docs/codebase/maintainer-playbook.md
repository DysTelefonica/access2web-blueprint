[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Manual del mantenedor

> Esta página rastrea `main`, que puede contener cambios sin liberar. Para la última versión liberada, véase la etiqueta correspondiente.

Use esta página antes de revisar o hacer un cambio. Convierte el mapa del repositorio en guardrails prácticos.

## Capacidades del producto

| Capability | Vista del mantenedor |
|---|---|
| Multi-app Lanzadera | ¿La nueva app declara su `[<APP>]` en el título y su `(app)` o `(platform)` en el commit? |
| Hexagonalidad | ¿El nuevo código respeta las capas y las reglas del bloque `architecture` de `docs/architecture.md`, que impone el job `docs-architecture`? |
| Auth (Argon2id + JWT) | ¿La ruta nueva pasa por el middleware y los servicios `app/src/modules/lanzadera/domain/services/issue_reset_token.py` / `consume_reset_token.py` si toca reset? |
| CI por niveles | ¿El cambio sabe en qué nivel corre (PR, nightly, release) y respeta `scripts/check_pr_size.py` (≤400)? |
| Migración Alembic | ¿La nueva migración sigue Expand & Contract (D82), tiene una de bajada y un test en `tests/lanzadera/migrations/`? |
| E2E | ¿El test nuevo usa `e2e/<spec>.spec.ts` y se reserva al release, no al PR? |
| Quality gates | ¿Si añadió un gate, lo cableó en `.github/workflows/ci.yml`, en `docs/gates.md` y en `tests/test_ci_workflow.py`? |

## Cómo navegar como mantenedor

1. Lea primero la doc de usuario (`README.md`, `DOCS.md`).
3. Encuentre el paquete que owns el cambio en [Mapa del repositorio](repository-map.md).
4. Inspeccione los tests cercanos antes de cambiar código.
5. Actualice la página de doc más pequeña y relevante.

## Guardrails

- **No toque el estado de runtime local**: `data/staging/<app>/` se R2-pullea; las fixtures viven en `tests/lanzadera/` o las crea `scripts/generate_lanzadera_fixtures.py`.
- **No invente internos externos**: Engram, dysflow y Gentleman-Programming tienen owner externo; apunte a su contrato en vez de replicarlo.
- **No evite los adapters**: las rutas HTTP y los paths de `scripts/` viven en sus paquetes, no en `application`.
- **No duplique referencias**: enlace a `docs/gates.md`, `docs/architecture.md` o las páginas de `docs/codebase/` en vez de copiar el detalle.
- **No rebase ni force-push en una PR abierta**: use merge de origin/main (Hard Rule 12).

## Checklists por tipo de cambio

### Cambiar una ruta HTTP de Lanzadera

- [ ] Identifique el router correcto en `app/src/modules/lanzadera/delivery/http/`.
- [ ] Añada o actualice el test del flujo en `tests/lanzadera/delivery/` (unit) o `tests/lanzadera/application/` (unit de use case).
- [ ] Si toca Postgres, marque el test con `@pytest.mark.integration`.
- [ ] Si añade plantilla, declárela en [Interfaces](interfaces.md) §Plantillas.
- [ ] Si añadió un endpoint nuevo a la tabla de rutas en [Interfaces](interfaces.md).

### Cambiar un caso de uso de Expedientes

- [ ] Cree o edite `app/src/modules/expedientes/application/<case>/{command,service}.py`.
- [ ] Añada test en `tests/expedientes/application/<case>/`.
- [ ] Si la ruta HTTP la expone Lanzadera o Expedientes, coordine el router en su `app/src/modules/lanzadera/delivery/http/`.

### Cambiar un quality gate

- [ ] Implemente en `scripts/check_<name>.py`.
- [ ] Documente el contrato en `docs/gates.md` §2.
- [ ] Añada el wiring en `tests/test_ci_workflow.py`.
- [ ] Invoque el gate en `.github/workflows/ci.yml::quality` o en `scripts/quality_report.py`.
- [ ] Si el gate afecta a la matriz `required`, actualice también `scripts/check_required_jobs.py`.

### Cambiar CI

- [ ] Si añade un job, declárelo en la matriz `required` de `.github/workflows/ci.yml`.
- [ ] Si añade un trigger nuevo (`schedule`, `workflow_dispatch`), verifique que el `if` del job no lo excluye por error.
- [ ] Si renombra un step, actualice `tests/test_ci_workflow.py` (wiring pinado).
- [ ] No cree un gate de CI adicional sin PR que documente el contrato: este repo modela los gates como `scripts/check_*.py`, no como workflows nuevos.

### Cambiar doc

- [ ] Valide las citas contra el código fuente.
- [ ] Use los patrones de `skills/documentation-alan-style/SKILL.md` (tono castellano peninsular formal, tablas, sin marketing).
- [ ] Si cambia una ruta citada en una página de `docs/codebase/`, regenere la comprobación con `test -e` y péguela en el cuerpo de la PR.
- [ ] Mantenga filenames en inglés y contenido en castellano peninsular formal.

## Playbook de revisión de PR

| Pregunta de revisión | Por qué importa |
|---|---|
| ¿El cambio tiene un único owner? | Evita comportamiento disperso y drift futuro. |
| ¿El test está en la suite correcta (`unit` / `integration` / `e2e_smoke` / `e2e`)? | Una suite mal etiquetada gasta runner o esconde fallos. |
| ¿Si toca Postgres, el test lleva `@pytest.mark.integration`? | Sin marker, el test corre en `quality` sin DB y falla. |
| ¿El PR respeta el presupuesto de 400 líneas? | PRs más largas rompen la disciplina de revisión. |
| ¿Las citas de paths en la doc están verificadas con `test -e`? | Citar paths falsos erosiona la confianza del codeguide. |
| ¿El commit usa Conventional Commits con scope `(app)` o `(platform)`? | La convención mantiene el changelog navegable. |
| ¿Si añade un gate, actualizó `docs/gates.md` y `tests/test_ci_workflow.py`? | Sin docs y sin wiring, el gate es invisible al siguiente mantenedor. |
| ¿Se evitó rebase y force-push? | El repo exige historial lineal vía merge de origin/main. |

## Si necesitas X, lee/toca Y

| Si necesitas... | Empieza por... | Luego revisa... |
|---|---|---|
| Revisar un PR | La sección «Playbook de revisión de PR» de esta página | `docs/gates.md` |
| Saber qué comprobar según el tipo de cambio | La sección «Listas por tipo de cambio» de esta página | `docs/codebase/repository-map.md` |
| Decidir dónde va un fichero nuevo | `docs/codebase/repository-map.md` | `docs/architecture.md` |

## Regla de colocación

Esta página no aloja código ni reglas nuevas: solo enruta. Una regla de revisión nueva va en la lista de su tipo de cambio y, si es un gate, en `docs/gates.md`.

## Lista de comprobación del contribuidor

- [ ] El cambio respeta los 5 gates del workflow (`scripts/check_pr_size.py`, el `check_issue_spec.py` del canon, `scripts/check_walkthrough_schema.py`, ruff, mypy) — ver `docs/gates.md` §2.
- [ ] Si toca `app/`, `tests/` u `openspec/`, leyó `docs/decisions/README.md` antes de empezar.
- [ ] Si toca `docs/`, leyó `skills/documentation-alan-style/SKILL.md` y aplica tono castellano peninsular formal con usted.
- [ ] El commit usa Conventional Commits con el scope `(app)` o `(platform)` correspondiente.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, parte por unidad de trabajo o encadena.
- [ ] Si añadió un nuevo gate, lo documentó en `docs/gates.md` y en `tests/test_ci_workflow.py`.
- [ ] Ningún `Co-Authored-By` ni atribución de IA en los commits.

## Siguiente paso

Ha completado el codeguide. Vuelva a [Codebase Guide](../CODEBASE-GUIDE.md) o, si el cambio toca un área concreta, a la página especializada correspondiente.