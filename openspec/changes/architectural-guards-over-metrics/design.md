<!-- Change: architectural-guards-over-metrics · Project: access2web-blueprint · Issue: #266 · Date: 2026-08-13 -->

# Design: guards arquitectónicos en lugar de métricas derivadas

> **Frase que organiza**: una decisión arquitectónica sin guard es prosa, y un guard que nadie cruza contra la tabla de decisiones es folclore. Este diseño construye el cruce mecánico entre ambos y se somete a él antes que ningún otro documento.

## Autocomprobación de dogfooding — primera comprobación

Este documento es el **primer artefacto que `scripts/check_decision_guards.py` debe poder leer sin quejarse**. Su tabla §Decisiones arquitectónicas está escrita en la forma exacta que el gate parsea (primera columna `ID`, fila delimitadora válida, un único token registrado por celda) y su tabla §Cableado de guards declara, decisión por decisión, la forma de cobertura que la sostiene. Si el gate no pudiera parsear la tabla de este diseño, el diseño estaría mal, no el gate. La verificación es literal y ejecutable:

```
python scripts/check_decision_guards.py --root .
# debe listar los IDs DG-1..DG-13 de este fichero, reconocer UNA sola tabla de decisiones en él,
# y no emitir malformed_table ni malformed_row sobre ninguna de sus otras tablas
```

**Esta autocomprobación ya ha atrapado un defecto en este mismo documento.** La primera redacción tenía una tabla §Cableado de guards encabezada también por `ID`, con celdas de lista (`DG-1, DG-2, DG-3, …`): satisfacía las tres condiciones de reconocimiento de DG-2 y el gate la habría tomado por una segunda tabla de decisiones, parseando la celda-lista como un identificador malformado. `DG-13` añade el discriminante que faltaba y la cabecera de esa tabla pasa a ser `Decisiones`. El fallo no era hipotético: estaba escrito en la línea 72 del propio diseño que declara el gate.

## Enfoque técnico

Opción A de la exploración. Un gate stdlib-only, con bloque `CONFIGURATION` separado del mecanismo, idéntico en forma a los seis gates ya presentes: recorre todos los `design.md` bajo `openspec/`, extrae los IDs de cada tabla de decisiones, recorre las cabeceras `HARNESS-PROVENANCE` de `tests/`, y exige que cada ID quede cubierto por un guard, por una entrada `COVERED_BY` o por una entrada `BASELINE` con `target_date` obligatoria. El envelope JSON y el contrato de `--root`/`--json` son los de la familia; `quality_report.py` lo agrega sin código nuevo salvo su fila en `GATES` y sus significados de indicador.

## Decisiones arquitectónicas

| ID | Decisión | Rationale | Alternativa descartada | Heredada |
|---|---|---|---|---|
| DG-1 | El registro de decisiones es la tabla markdown del propio `design.md`; no se crea `.sdd/decision-guards.yml`. El gate aísla el parseo en `parse_decision_tables()` para que un backend alternativo sea sustituible sin tocar `evaluate()`. | La tabla ya existe y es la que el humano lee y edita; un segundo fichero introduce dos verdades que divergen. La frontera de función deja la opción B abierta sin reescritura. | Fichero de registro YAML como fuente primaria (opción B de la exploración). | Exploración §4, propuesta §Enfoque |
| DG-2 | Una tabla es tabla de decisiones si su primera celda de cabecera normaliza a `id`, su fila delimitadora es válida y **al menos una** fila de datos es fila de decisión en el sentido estricto de DG-13. Las tablas con cabecera `ID` sin ninguna fila así (la tabla de riesgos `R-*` de `lanzadera-mvp`) se ignoran sin verdicto. | Verificado: `lanzadera-mvp/design.md` tiene dos tablas con cabecera `ID` (líneas 148 y 333). «Cabecera `ID`» sola barrería los riesgos `R-1..R-15` y exigiría un guard por riesgo. | Firma de cabecera completa (`ID\|Decisión\|Rationale\|…`): frágil ante traducción y ante cualquier columna nueva. | Verificación directa del fichero |
| DG-3 | Fallo cerrado explícito: sin ningún `design.md`, con conjunto de IDs vacío, o con un `design.md` sin tabla reconocible y no declarado en `UNGOVERNED_DESIGNS`, el gate sale `1` con `status: "error"`. | Hard Rule 18 y el precedente de `check_crap.py`: «no pude medir» y «está limpio» no pueden compartir verdicto. Un gate que no encuentra decisiones y reporta verde es el peor resultado posible. | Salir `0` cuando no hay nada que comprobar. | `check_crap.py` §CoverageUnavailable, `check_layers.py` §files_seen |
| DG-4 | Un guard declara la decisión que codifica **solo** en su bloque `HARNESS-PROVENANCE`: el tramo contiguo de líneas `#` de cabecera que abre con ese marcador, con continuación en líneas siguientes. El docstring no otorga cobertura. | Un único sitio de declaración, extraíble mecánicamente. La prosa del docstring menciona IDs para explicar, no para reclamar, y contarla generaría cobertura falsa por cita. | Buscar el token en todo el fichero o en el docstring: falsos positivos por mención. | `tests/lanzadera/auth/test_no_legacy_compat_smoke.py` |
| DG-5 | La vía de escape `covered-by` vive en `COVERED_BY` del bloque `CONFIGURATION`, y cada entrada nombra un script que debe existir **y** figurar en la tupla `GATES` de `quality_report.py`, leída por AST sin importar el módulo. | Una cobertura declarada por un gate desactivado es cobertura inexistente. Este cambio retira CRAP: sin esta validación, una decisión cubierta por `check_crap.py` habría quedado huérfana en silencio. | Parsear la tabla «Quality gates wiring» del `design.md`: segundo parser de markdown, segundo punto de rotura. | Propuesta §Enfoque, `quality_report.py` §GATES |
| DG-6 | Una entrada `COVERED_BY` puede declarar `asserts={"CONSTANTE": valor}`; el gate lee esa asignación de módulo del script nombrado por AST y falla con `constant_drift` si difiere. La deriva **entre prosa y código** queda fuera de alcance declarado. | Responde a la deriva de `DA-1` de forma determinista y sin parsear prosa. Reconoce el límite: el gate verifica que la expectativa declarada case con el código, no que la prosa case con la expectativa. | Detección general de constantes citadas en prosa: semántica, no determinista, falsos positivos altos. | Hallazgo bloqueante de la propuesta |
| DG-7 | Los IDs son únicos globalmente en todo `openspec/`. Dos ficheros con el mismo ID fallan con `duplicate_id`. Este cambio usa el prefijo propio `DG-`. | Un guard reclama un token desnudo; si `DA-1` existiera en dos changes, la cobertura sería ambigua y el gate mentiría en uno de los dos. | Cualificar el ID por nombre de change (`lanzadera-mvp/DA-1`): rompe las cabeceras `HARNESS-PROVENANCE` ya escritas. | DG-4 |
| DG-8 | La deuda calendarizada usa `BaselineEntry(count, target, target_date)`, campo por campo el mismo de `check_layers.py`, con `target_date` obligatoria. Un ID en `BASELINE` pasada su fecha y aún descubierto falla. | Idioma idéntico en los siete gates: quien lee uno lee todos. Hard Rule 12: un ratchet es rampa, no destino. | Lista de exenciones sin fecha. | `check_layers.py` §BaselineEntry |
| DG-9 | `decision_guards` va **último** en `GATES`, después de `legacy_hashes`. | El orden está pinneado en código porque es carga útil: primero lo barato y estructural, cuyos fallos mueven los números de abajo. Este gate no es una métrica y no consume ni invalida ninguna cifra previa; además `COVERED_BY` depende de que los gates citados hayan corrido, así que el fallo de base debe verse antes que el de la reclamación. | Colocarlo tras `layers`: adelantaría un fallo de disciplina documental por delante de un fallo de arquitectura real. | `quality_report.py` §GATES, comentario de `legacy_hashes` |
| DG-10 | El recorrido cubre **todos** los `design.md` bajo `openspec/`, archivados incluidos. La tabla es la fuente de verdad: sin fila no hay guard exigible, y un guard que cita un ID inexistente falla con `orphan_guard`. | Una decisión archivada sigue rigiendo el código vivo. `orphan_guard` es el mecanismo que hace que «la decisión superada se borra con su guard» sea comprobable en lugar de recordable. | Recorrer solo `openspec/changes/` activo. | Decisiones de producto 3 y 4 |
| DG-11 | `MAX_COMPLEXITY` pasa de 15 a 10 con la fecha de revisión `2027-02-13` escrita junto a la constante. La cobertura de esta decisión es `COVERED_BY` con `asserts={"MAX_COMPLEXITY": 10}`. | Retirado CRAP, el techo efectivo saltaría de 6 a 15 sin que nadie lo decidiera. El 10 es transitorio: la revisión decide si baja hacia 6. Hoy ninguna función pasa de CC 6, así que el cambio no toca código. | Dejar 15: eleva el techo real en silencio. Bajar a 6 ya: convierte una retirada en un refactor. | Decisión de producto 1, DG-6 |
| DG-12 | CRAP se retira por completo: script, fixtures, wiring test y toda referencia rastreada. Un guard dedicado impide su reintroducción parcial. | Verificado: 20 ficheros rastreados mencionan `crap`. Una retirada a medias deja un `make check-crap` que apunta a un script inexistente. | Dejar el script sin cablear en `GATES`: código muerto que parece gate. | Propuesta §Alcance |
| DG-13 | Una fila es **fila de decisión** solo si su primera celda normalizada contiene **exactamente un** token de `DECISION_ID_PREFIXES` y ese token abre la celda. Una celda de lista o con separador (`DG-1, DG-2, DG-5`) no lo es, y una tabla cuyas filas son todas así no es tabla de decisiones. Por defensa adicional, las tablas de cableado encabezan su primera columna con `Decisiones`, no con `ID`. | Verificado sobre este mismo documento: su tabla de cableado cumplía las tres condiciones de DG-2 —cabecera `ID`, delimitadora válida, primera celda abriendo por `DG-`— y el gate la habría clasificado como segunda tabla de decisiones, leyendo `DG-1, DG-2, DG-3, …` como un identificador malformado. Es exactamente el modo de fallo que DG-2 pretendía evitar, del que la tabla de riesgos `R-*` se libró solo por el prefijo. Contar tokens es determinista, independiente del idioma y estable ante columnas nuevas, que es lo que DG-2 exigía de su discriminante. | Excluir por posición («solo la primera tabla con cabecera `ID` del fichero»): arbitraria y silenciosa, rompe en cuanto un diseño reordena sus secciones. Excluir por lista negra de cabeceras: reintroduce la fragilidad de firma que DG-2 ya descartó. Renombrar la cabecera y nada más: apoya la corrección en la disciplina de quien escribe, que es justo lo que este change existe para dejar de suponer. | DG-2 |

## Contrato de parseo — la junta más débil

El parseo de markdown es la pieza que puede convertir el gate en ruido. Se acota así.

**Forma esperada.** Tabla GFM de tuberías. Fila de cabecera cuya primera celda normalizada (sin espacios, sin backticks, sin `**`, en minúsculas) es `id`. Fila delimitadora inmediatamente posterior, con el mismo número de columnas y todas las celdas casando `^:?-{3,}:?$`. Filas de datos contiguas hasta la primera línea que no empieza por `|`.

**Fila de decisión (DG-13).** La primera celda, normalizada, casa `^(DA|DG|QC)-\d+\b` **y** contiene una sola aparición de ese patrón en toda la celda. El token capturado es el ID. Se tolera sufijo tras el token —la forma `R-1 (H1)` de la tabla de riesgos demuestra que el idioma existe— siempre que ese sufijo no aporte un segundo token registrado. Una celda como `DG-1, DG-2, DG-5` aporta tres y por tanto no es fila de decisión: es una celda de cableado, y así queda excluida sin depender de la cabecera.

`D-` queda **fuera** de `DECISION_ID_PREFIXES`: verificado que las decisiones heredadas se escriben `D8`, `D88`, sin guion, y nunca aparecen en columna `ID`; incluir `D-` sería precisión inventada. `R-` queda fuera porque ninguna decisión se registra con ese prefijo.

**Dos discriminantes, no uno.** El recuento de tokens (DG-13) es el que decide; el nombre de la cabecera es defensa en profundidad. Ningún fichero depende de que quien escribe recuerde no usar `ID` en una tabla de cableado, porque esa dependencia es precisamente la que este change existe para eliminar.

**Verdictos.** Nada se salta en silencio.

| Condición | Clave | Verdicto |
|---|---|---|
| No hay `design.md` bajo `openspec/`, o el total de IDs es 0 | — | `status: error`, exit 1 |
| `design.md` sin tabla reconocible y ausente de `UNGOVERNED_DESIGNS` | `no_decision_table` | fail |
| Cabecera `ID` sin delimitadora válida o con recuento de columnas distinto | `malformed_table` | fail |
| Fila de una tabla ya reconocida cuyo primer campo no casa | `malformed_row` | fail |
| Mismo ID en dos ficheros | `duplicate_id` | fail |
| ID sin guard, sin `COVERED_BY` y sin `BASELINE` | `uncovered` | fail |
| `COVERED_BY` que nombra un script ausente o fuera de `GATES` | `stale_coverage` | fail |
| Constante de `asserts` con valor real distinto | `constant_drift` | fail |
| Cabecera `HARNESS-PROVENANCE` que cita un ID inexistente | `orphan_guard` | fail |
| Entrada `BASELINE` cuyo ID ya no existe en ninguna tabla | — | `NOTE`, no falla |

**Frontera de exclusión.** `EXCLUDED_PARTS` incluye `fixtures` además de la lista de la familia. Sin ello, los árboles `tests/fixtures/decision_guards_violation/` —que contienen a propósito cabeceras que citan IDs falsos— entrarían en el recorrido de producción y producirían `orphan_guard` sobre el repositorio real.

## Reconciliación de `DA-1`

`DA-1` declara `ROOT_PACKAGE = "platform.src.modules"`. Verificado: `check_layers`, `check_complexity`, `check_crap` y `check_dry` declaran `ROOT_PACKAGE = "app"` con `MODULE_PREFIX = ("src", "modules")`, y no existe directorio `platform/`.

**Se corrige la tabla, no el código.** `DA-1` pasa a declarar `ROOT_PACKAGE = "app"` y `MODULE_PREFIX = ("src", "modules")`, con nota de corrección fechada que cita este change. Renombrar el paquete a `platform` se rechaza por dos motivos: arrastra todos los ficheros del árbol y todas las rutas de los gates, y `platform` es un módulo de la biblioteca estándar de Python, de modo que un paquete raíz con ese nombre lo ensombrece para cualquier `import platform` ejecutado desde la raíz del repositorio.

**¿Debe el gate detectar esta clase de deriva?** Parcialmente, y con el límite escrito. Vía `asserts` (DG-6) detecta de forma determinista que una expectativa declarada no case con la constante real, y `DA-1` se registra con `asserts={"ROOT_PACKAGE": "app"}`, de modo que la deriva corregida hoy no puede volver en silencio. Lo que queda fuera de su alcance, declarado como no-objetivo: comprobar que la **prosa** de la fila diga lo mismo que la expectativa declarada. Eso exige leer lenguaje natural y no es determinista; lo sostiene la revisión humana. La fila `DA-13` cita rutas `platform/src/...` con la misma deriva de prosa: se corrige en la misma PR por lectura, no por gate.

## Cableado de guards

La primera columna se encabeza `Decisiones`, nunca `ID`, y sus celdas son listas: por DG-13 el gate no puede confundir esta tabla con una tabla de decisiones ni por su forma ni por su cabecera.

| Decisiones | Forma de cobertura | Artefacto | Vence |
|---|---|---|---|
| DG-1, DG-2, DG-3, DG-4, DG-7, DG-10, DG-13 | guard | `tests/lanzadera/test_decision_guards_wiring.py` (dos tiers) | — |
| DG-5, DG-6, DG-8 | guard | mismo fichero, casos `stale_coverage`, `constant_drift` y `BASELINE` vencida sobre fixture | — |
| DG-9 | guard | `tests/lanzadera/test_quality_report_smoke.py` (modificado: pinea posición y orden) | — |
| DG-11 | covered-by | `scripts/check_complexity.py`, `asserts={"MAX_COMPLEXITY": 10}` | — |
| DG-12 | guard | `tests/lanzadera/test_crap_retired_guard.py` | — |
| DA-1 | covered-by | `scripts/check_layers.py`, `asserts={"ROOT_PACKAGE": "app"}` | — |
| DA-13 | covered-by | `scripts/check_legacy_hashes.py` + `tests/lanzadera/auth/test_no_legacy_compat_smoke.py` | — |
| DA-2 … DA-12 (once) | BASELINE | `count=1, target=0, target_date="2027-02-13"` | 2027-02-13 |

Segunda comprobación de coherencia con las propias reglas: verificado que el bloque `HARNESS-PROVENANCE` de `tests/lanzadera/test_quality_report_smoke.py` es hoy `deterministic-quality-harness v1.4 + lanzadera-mvp`, **sin ningún token de ID**; su `QC-11` vive en el docstring, que por DG-4 no otorga cobertura. La rebanada que modifique ese fichero debe ampliar su bloque de provenance a `… + architectural-guards-over-metrics DG-9`, o `DG-9` quedaría `uncovered`. El mismo DG-4 es lo que impide que ese `QC-11` del docstring dispare un `orphan_guard`.

Fecha única para las once, no once fechas escalonadas: una sola conversación de revisión, coincidente con la revisión del techo de complejidad. Si el responsable prefiere escalonar por riesgo, el cambio es de datos, no de mecanismo.

## Flujo de datos

```text
openspec/**/design.md ──parse_decision_tables()──→ {ID: (fichero, línea)}
                                                        │
tests/**/*.py  ──collect_guard_claims()───────────→ {ID: [ruta_test]}
   (bloque HARNESS-PROVENANCE, sin fixtures)            │
                                                        │
CONFIGURATION.COVERED_BY ─validate_coverage()────→ {ID: gate}
        │                                               │
        └─AST─→ quality_report.GATES + constante del script citado
                                                        │
CONFIGURATION.BASELINE ──────────────────────────→ {ID: plazo}
                                                        ▼
                                      evaluate(today) → (exit_code, líneas)
                                                        │
                                      build_report() → envelope JSON
                                                        ▼
                                              scripts/quality_report.py
```

## Cambios de ficheros

| Fichero | Acción | Descripción |
|---|---|---|
| `scripts/check_decision_guards.py` | Crear | El gate: `CONFIGURATION` + mecanismo + envelope + `--root/--json` |
| `tests/lanzadera/test_decision_guards_wiring.py` | Crear | Guard de dos tiers del propio gate |
| `tests/fixtures/decision_guards_clean/` | Crear | `openspec/` + `tests/` mínimos: una decisión con su guard **y** una tabla de cableado con celdas-lista que el gate debe ignorar (caso DG-13) |
| `tests/fixtures/decision_guards_violation/` | Crear | ID descubierto, tabla malformada, guard huérfano, cobertura obsoleta |
| `tests/lanzadera/test_crap_retired_guard.py` | Crear | DG-12: la retirada no se deshace por partes |
| `scripts/quality_report.py` | Modificar | `GATES`: fuera `crap`, `decision_guards` al final; `INDICATOR_MEANINGS` nuevos |
| `scripts/check_complexity.py` | Modificar | `MAX_COMPLEXITY = 10` + fecha de revisión `2027-02-13` |
| `openspec/changes/lanzadera-mvp/design.md` | Modificar | `DA-1` reconciliada; `DA-13` sin rutas `platform/`; fila CRAP de la tabla de wiring sustituida |
| `scripts/check_crap.py`, `tests/lanzadera/test_crap_wiring.py`, `tests/fixtures/crap_clean/`, `tests/fixtures/crap_violation/` | Borrar | Superficie de retirada |
| `Makefile`, `.github/workflows/ci.yml`, `app/pyproject.toml`, `app/README.md`, `.gitignore`, `tests/test_ci_workflow.py`, `tests/test_gate_liveness.py`, `tests/test_gate_smoke.py`, `tests/lanzadera/test_quality_report_smoke.py`, docstrings de `check_complexity.py` y `check_mutation.py` | Modificar | 20 ficheros rastreados mencionan `crap` |

## Interfaces

```python
@dataclass(frozen=True)
class Coverage:
    """Cobertura declarada por un gate estructural ya cableado."""
    gate: str                      # nombre de script en scripts/, presente en quality_report.GATES
    reason: str
    asserts: dict[str, object] = field(default_factory=dict)   # constante de módulo -> valor esperado

@dataclass(frozen=True)
class BaselineEntry:
    count: int          # 1 mientras la decisión siga descubierta
    target: int         # 0
    target_date: str    # ISO-8601, obligatoria

COVERED_BY: dict[str, Coverage] = {...}
BASELINE: dict[str, BaselineEntry] = {...}
UNGOVERNED_DESIGNS: dict[str, str] = {}   # design.md sin tabla, cada uno con su motivo
```

Envelope: `{"gate": "decision_guards", "status": ..., "indicators": {"decisions_total", "decisions_uncovered", "guards_total", "orphan_guards", "designs_scanned", "designs_ungoverned"}, "ceilings": {"decisions_uncovered": 0, "orphan_guards": 0, "designs_ungoverned": 0}, "findings": [...]}`.

## Estrategia de tests

| Tier | Qué se prueba | Cómo |
|---|---|---|
| Detector (tier 1) | Cada clave del contrato de verdictos | Ejecutar el gate con `--root` sobre `decision_guards_violation/`; assertar exit 1 y la clave concreta en el envelope |
| Detector negativo | El scanner no puede estar vacuamente verde | `--root` sobre `decision_guards_clean/`; exit 0 y `decisions_total > 0` |
| Discriminante (DG-13) | Una tabla de cableado no se toma por tabla de decisiones | Fixture clean con ambas tablas; assertar `decisions_total` igual al recuento de la tabla real y ausencia de `malformed_row` |
| Producción (tier 2) | El repositorio real cumple hoy | `--root .`; exit 0, `decisions_uncovered == 0`, `orphan_guards == 0` |
| Wiring | Posición y presencia en `GATES` | `test_quality_report_smoke.py` pinea la tupla completa en orden |
| Liveness | El gate falla de verdad | `test_gate_liveness.py` y `test_gate_smoke.py` incorporan el script nuevo y pierden el de CRAP |

## Matriz de amenazas

| Frontera | Aplicabilidad | Respuesta de diseño | Test RED planificado |
|---|---|---|---|
| Rutas tipo documentación | **Aplicable**: el gate clasifica `design.md` y `*.py` como entrada y podría barrer árboles de fixture | `EXCLUDED_PARTS` incluye `fixtures`; raíces cerradas `openspec/` y `tests/`, nunca glob desde la raíz | Fixture con guard de cabecera falsa bajo `tests/fixtures/`: el recorrido de producción no debe verlo |
| Selección de repositorio git | N/A: el gate no invoca git; recibe `--root` como todos sus hermanos y `quality_report.py` mantiene la única llamada a git | — | — |
| Estado de commit | N/A: el gate no lee índice ni worktree de git | — | — |
| Estado de push | N/A: sin operación de red ni de refs | — | — |
| Comandos de PR | N/A: sin automatización de PR. `quality_report.py` lo lanza como subproceso con argv fijo y sin shell, exactamente como a los seis existentes | — | — |

## Universal frente a específico del repositorio

| Universal — a la skill `deterministic-quality-harness` | Específico — bloque `CONFIGURATION` |
|---|---|
| La forma del gate de cruce decisión↔guard y su envelope | `OPENSPEC_ROOT`, `TESTS_ROOT`, `SCRIPTS_ROOT` |
| El reconocimiento de tabla por cabecera `ID` + fila de decisión (DG-2) | `DECISION_ID_PREFIXES = ("DA-", "DG-", "QC-")` |
| La fila de decisión como celda de token único, y la cabecera `Decisiones` para las tablas de cableado (DG-13) | Qué tablas concretas de este repositorio son de cableado |
| El contrato de verdictos y el fallo cerrado (DG-3) | `UNGOVERNED_DESIGNS`, `EXCLUDED_PARTS` |
| El bloque `HARNESS-PROVENANCE` como única sede de la reclamación (DG-4) | `COVERED_BY` y sus `asserts` concretos |
| La validación de `covered-by` contra la tupla de gates cableados (DG-5) | `BASELINE` y sus fechas |
| El idioma `BASELINE` + `target_date` obligatoria (DG-8) | La decisión concreta de retirar CRAP |
| La unicidad global de IDs y `orphan_guard` (DG-7, DG-10) | El techo 10 y su fecha de revisión |

La skill vive en `C:\Proyectos\skills\skills\deterministic-quality-harness\SKILL.md` y se actualiza allí. **No se vendoriza en este repositorio** (decisión explícita del responsable).

## Entrega — frontera de encadenado recomendada a `sdd-tasks`

El presupuesto de revisión es de 400 líneas y el trabajo no cabe en una PR: solo `check_crap.py` son 391 líneas de borrado, y el gate nuevo con sus dos árboles de fixture es código autorizado nuevo de orden similar. Seis rebanadas, cada una verde por sí sola y revertible con un `git revert`:

| # | Rebanada | Contenido | Presupuesto |
|---|---|---|---|
| 1 | Techo de complejidad | `MAX_COMPLEXITY = 10` + fecha `2027-02-13` + su wiring test | ~30 |
| 2 | Descableado de CRAP | Fuera de `GATES`, `Makefile`, `ci.yml`, `pyproject.toml`, docs y tests de wiring | ~250 |
| 3 | Borrado de CRAP | `check_crap.py`, sus dos árboles de fixture y su wiring test | ~500, borrado puro |
| 4 | Gate, sin armar | `check_decision_guards.py` + tests unitarios, **sin** entrada en `GATES` | ~380 |
| 5 | Fixtures de dos tiers | `decision_guards_clean/` y `decision_guards_violation/` + wiring test | ~300 |
| 6 | Armado | `DA-1` reconciliada, `COVERED_BY`, `BASELINE`, entrada en `GATES`, `ci.yml`, `Makefile`, bloque `HARNESS-PROVENANCE` de `test_quality_report_smoke.py` ampliado con `DG-9` | ~200 |

Dos ordenaciones son obligatorias, no preferencias. La rebanada 1 va **antes** que la 2 y la 3: mientras CRAP siga vivo el techo efectivo es 6, así que bajar a 10 no rompe nada; en el orden inverso el techo flota a 15 durante toda la cadena. La rebanada 6 va **la última**: cablear el gate en `GATES` antes de poblar `COVERED_BY` y `BASELINE` deja CI en rojo al fusionar. La rebanada 3 es borrado puro de código ya revisado y es candidata razonable a `size:exception` en lugar de a otro corte.

Cadena de PRs: 1 → 2 → 3 → 4 → 5 → 6, cada hija apuntando a la rama de la anterior sobre `feat/266-guards-arquitectonicos`.

## Migración y rollout

No hay migración de datos. El gate es aditivo hasta la rebanada 6; retirarlo de `GATES` lo desactiva sin tocar ningún test de producción. La retirada de CRAP es irreversible solo en apariencia: `git revert` de las rebanadas 2 y 3 restaura script, fixtures y cableado, y nada externo depende de ellos.

## Convención que `sdd-tasks` debe propagar

Toda tabla que asocie varias decisiones a un artefacto encabeza su primera columna con `Decisiones` y no con `ID` (DG-13, defensa en profundidad). Verificado sobre `lanzadera-mvp/design.md`: su tabla «Quality gates wiring» ya encabeza `| Gate | Mecanismo | Archivos | Decisiones |`, de modo que **no requiere renombrado**; la única tabla con cabecera `ID` que no es de decisiones es la de riesgos `R-*`, excluida por prefijo. La corrección de cabecera se aplica, pues, a este documento; la convención se propaga a los diseños futuros.

## Preguntas abiertas

- [ ] Si algún día se registrara `R-` en `DECISION_ID_PREFIXES`, la tabla de riesgos de `lanzadera-mvp` pasaría a ser tabla de decisiones —sus celdas `R-1 (H1)` son de token único y superarían DG-13— y exigiría un guard por riesgo. El prefijo no se registra y la razón queda escrita aquí.
- [ ] Las once `target_date` se proponen todas en `2027-02-13`. Si el responsable prefiere escalonarlas por riesgo de la decisión, es un cambio de datos en `BASELINE` y `sdd-tasks` puede recogerlo sin tocar el mecanismo.
- [ ] `QC-` figura en `DECISION_ID_PREFIXES` de forma preventiva: hoy ningún `design.md` declara una tabla con IDs `QC-*`. Si se decide que las QC nunca serán filas de tabla, el prefijo sobra.
- [ ] La rebanada 3 (~500 líneas de borrado puro) puede entregarse como `size:exception` justificado o partirse en dos; la decisión corresponde a la estrategia de entrega, no al diseño.
