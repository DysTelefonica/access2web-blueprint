<!-- Change: architectural-guards-over-metrics · Project: access2web-blueprint · Issue: #266 · Date: 2026-08-13 -->

# Propuesta: guards arquitectónicos en lugar de métricas derivadas

> **Frase que organiza**: el repositorio deja de defender su arquitectura con una métrica derivada que hoy no señala nada, y pasa a defenderla con tests que codifican decisiones registradas, más un gate que comprueba mecánicamente que ninguna decisión se quede sin su guard.

## Intención

`scripts/check_crap.py` lleva `BASELINE = {}` y una ejecución en vivo confirma cero infractores. No obstante, no está ocioso: como `(1 − cobertura)³ ≥ 0`, se cumple siempre `CRAP(f) ≥ CC(f)`, de modo que **el techo de complejidad efectivo del repositorio es 6, no el 15 nominal de `check_complexity.py`**. CRAP mantiene las funciones pequeñas mientras reporta cero hallazgos, y para hacerlo exige un run completo de cobertura por PR.

Un techo derivado que nadie declaró no explica ninguna decisión de arquitectura. La sustitución que se propone es de naturaleza distinta, no de calibre: guards que citan la decisión que protegen, y un gate que impide que «escribir un guard por decisión» degenere en una convención que se recuerda a ratos.

## Alcance

### Incluye

- Retirada de CRAP: `scripts/check_crap.py`, su entrada en `GATES` de `scripts/quality_report.py`, `tests/lanzadera/test_crap_wiring.py` y los árboles `tests/fixtures/crap_clean/` y `crap_violation/`. Arrastra referencias en `Makefile`, `.github/workflows/ci.yml`, `tests/test_ci_workflow.py`, `tests/test_gate_liveness.py`, `tests/test_gate_smoke.py`, `tests/lanzadera/test_quality_report_smoke.py`, `.gitignore`, `app/pyproject.toml`, `app/README.md` y los docstrings de `check_complexity.py` y `check_mutation.py`.
- `MAX_COMPLEXITY` de `check_complexity.py` pasa de 15 a **10**, con fecha de revisión explícita. Hoy ninguna función supera CC 6: el cambio no exige tocar código y concede holgura real. Dejarlo en 15 elevaría el techo efectivo de 6 a 15 en silencio.
- Nuevo `scripts/check_decision_guards.py`, cableado en `GATES`, con su `test_decision_guards_wiring.py` y sus fixtures clean/violation.
- Patrón de guard test de dos tiers documentado y aplicado a las decisiones `DA-*` ya cubiertas de hecho.

### No incluye

- Cualquier cambio en `scripts/check_layers.py`. Es el activo arquitectónico más fuerte del repositorio y nada de lo aquí propuesto lo sustituye.
- Escribir un guard nuevo por cada una de las trece decisiones `DA-*`. Esta propuesta entrega el gate y el patrón; el `BASELINE` con `target_date` obligatoria calendariza las decisiones aún descubiertas.
- Vendorizar la skill `deterministic-quality-harness` en el repositorio (decisión explícita: sus gates ya viven aquí como código).
- El fichero de registro `.sdd/decision-guards.yml` (opción B de la exploración). Queda como vía documentada para repositorios cuyo `design.md` no tenga tabla con columna de ID.

## Contrato de capabilities con `sdd-spec`

- **New capability**: `decision-guard-coverage` — toda decisión arquitectónica registrada con ID debe tener un guard que la cite, o una cobertura declarada por un gate estructural.
- **Modified capabilities**: ninguna. `openspec/specs/` está vacío hoy.

## Enfoque

Opción A de la exploración. `check_decision_guards.py` lee las filas con columna de ID de cada tabla de decisiones de `openspec/**/design.md` y exige que cada ID aparezca en una de dos formas:

1. la cabecera `HARNESS-PROVENANCE` o el docstring de un guard test bajo `tests/`, o
2. una entrada `covered-by` de un gate, formalizando lo que la tabla «Quality gates wiring» de `lanzadera-mvp/design.md` ya hace a mano.

Los guards siguen la forma de dos tiers que el repositorio ya escribe de modo informal: un test de corrección del detector contra fixtures sintéticas clean y violation, para que el scanner no pueda estar vacuamente verde, y una aserción de producción sobre el conjunto **cerrado** de ficheros que la decisión nombra. El gate falla cerrado si no encuentra ninguna tabla de decisiones bien formada, igual que `check_crap.py` se negaba a correr sin `coverage.json`.

### Universal frente a específico del repositorio

| Universal — expresable en la skill | Específico — bloque `CONFIGURATION` |
|---|---|
| Forma del gate de cruce decisión↔guard | `ROOT_PACKAGE = "app"`, prefijos de módulo |
| Forma de dos tiers del guard test | Nombres de capa hexagonal |
| Idioma `BASELINE` + `target_date` obligatoria | Prefijos de ID en uso: `DA-`, `QC-`, `D-` |
| Regla de conjunto cerrado frente a glob abierto | Raíz de `tests/` y de `openspec/` |
| Vía de escape `covered-by` | La decisión concreta de retirar CRAP |

La skill vive en `C:\Proyectos\skills\skills\deterministic-quality-harness\SKILL.md` y se actualiza allí, fuera de este repositorio.

## Coste aceptado

Retirada CRAP, **ningún gate exige ya que el código ramificado esté bien testeado por función y por PR** cuando ese código no está ligado a una decisión registrada. `check_complexity` ignora la cobertura, `--cov-fail-under=69` es una media de paquete y la mutación corre semanalmente. Los guards cubren solo lo que alguien decidió registrar. Se acepta con esta mitigación: techo de complejidad a 10 con fecha de revisión, `check_mutation_sites` como proxy de superficie estática, y la regla de que toda lógica de negocio no trivial que llegue sin decisión registrada se discute en revisión, no en un gate.

## Decisiones de producto confirmadas

Resueltas por el responsable del proyecto el 2026-08-13. Dejan de ser supuestos.

| Decisión | Resolución |
|---|---|
| Fecha de revisión del techo 10 | **2027-02-13**. La revisión decide si el margen se usó: si el código sigue en complejidad 6, el techo baja hacia 6. El 10 es transitorio, no destino |
| Cobertura de las trece `DA-*` | Solo las ya cubiertas de hecho reciben guard en esta entrega: `DA-1` vía `check_layers`, `DA-13` vía su test existente. Las once restantes se calendarizan con `BASELINE` y `target_date` |
| Alcance del recorrido del gate | **Todos** los `design.md` bajo `openspec/`, incluidos los archivados. Una decisión archivada arrastra sus IDs y sigue exigiendo guard |
| Decisión superada | Su guard **se borra con ella**. El gate lee la tabla como fuente de verdad, de modo que sin fila no hay guard exigible |

### Hallazgo bloqueante: `DA-1` está desactualizada

`lanzadera-mvp/design.md` declara en `DA-1` que `check_layers.py` usa
`ROOT_PACKAGE = "platform.src.modules"`. Los cuatro gates que declaran esa
constante —`check_layers`, `check_complexity`, `check_crap`, `check_dry`— usan
`"app"`, y no existe ningún directorio `platform/` en el repositorio.

Esto obliga a reconciliar `DA-1` **dentro de este cambio**. El gate nuevo lee la
tabla de decisiones como fuente de verdad: si nace apuntando a una configuración
que ningún gate usa, su primera lectura es falsa.

Es también la mejor evidencia a favor del gate. Esa deriva lleva tiempo ahí y
ningún mecanismo la detectó, porque hasta ahora nada leía la tabla.

## Áreas afectadas

| Área | Impacto | Descripción |
|---|---|---|
| `scripts/check_crap.py` + fixtures + wiring test | Eliminado | Superficie de retirada. Verificado: **20 ficheros rastreados** mencionan `crap`, no 13 |
| `openspec/changes/lanzadera-mvp/design.md` `DA-1` | Corregido | `ROOT_PACKAGE` declarado no coincide con el real (`app`) |
| `scripts/check_complexity.py` | Modificado | `MAX_COMPLEXITY` 15 → 10 + fecha de revisión |
| `scripts/check_decision_guards.py` | Nuevo | Gate de cruce decisión↔guard |
| `scripts/quality_report.py` | Modificado | `GATES`: fuera `crap`, dentro `decision_guards` |
| `.github/workflows/ci.yml`, `Makefile` | Modificado | Steps y targets |
| `openspec/changes/lanzadera-mvp/design.md` | Modificado | Tabla de wiring: fila CRAP → fila del gate nuevo |
| `scripts/check_layers.py` | Sin tocar | Explícitamente fuera de alcance |

## Riesgos

| Riesgo | Probabilidad | Mitigación |
|---|---|---|
| Pérdida de la garantía por función y por PR fuera de las decisiones registradas | Alta | Aceptada y nombrada arriba; techo 10 + `check_mutation_sites` |
| Disciplina de etiquetado: el guard no cita el ID | Media | Vía de escape `covered-by` + `BASELINE` con `target_date` |
| Falsos positivos del parser de tablas markdown ante un `design.md` mal formado | Media | Fixtures clean/violation del propio gate; fallo cerrado explícito |
| El presupuesto de 400 líneas se agota entre retirada y gate nuevo | Alta | `sdd-tasks` debe planificar PRs encadenadas: retirada primero, gate después |

## Plan de rollback

Cada PR es un commit convencional revertible. `git revert` de la PR de retirada restaura `check_crap.py`, sus fixtures y su cableado en `GATES`; nada externo depende de ellos. El gate nuevo es aditivo: retirarlo de `GATES` lo desactiva sin tocar ningún test de producción. El techo 10 se revierte cambiando una constante.

## Criterios de éxito

- [ ] `python scripts/quality_report.py` corre sin `crap` y con `decision_guards`, y ambos hechos están pinneados por test.
- [ ] `python scripts/check_decision_guards.py` falla cerrado sin tabla de decisiones y detecta un ID sin guard sobre la fixture violation.
- [ ] Toda fila `DA-*` de `lanzadera-mvp/design.md` está cubierta por un guard, por una entrada `covered-by`, o por una entrada `BASELINE` con `target_date`.
- [ ] `python scripts/check_complexity.py` pasa con `MAX_COMPLEXITY = 10` sin modificar código de aplicación.
- [ ] Ninguna referencia a `crap` queda en `Makefile`, `ci.yml`, `pyproject.toml` ni en la documentación.
- [ ] `DA-1` declara el `ROOT_PACKAGE` que los gates usan realmente, y el gate nuevo lo lee sin discrepancia.
- [ ] `check_complexity.py` declara su fecha de revisión `2027-02-13` junto al techo.
- [ ] El patrón queda expresado como universal frente a configuración en la skill `deterministic-quality-harness`.

## Siguiente paso

`sdd-spec` y `sdd-design` sobre esta propuesta. `sdd-tasks` debe forzar encadenado de PRs: el presupuesto de 400 líneas no cubre retirada y gate nuevo en una sola.
