# Porting guide — checklist de pre-vuelo para adoptar el patrón de CI

Checklist por fases para adoptar el patrón de CI (gates deterministas,
presupuesto de revisión, evidencia por SHA, gobernanza de orquestación) en un
repositorio nuevo. Cada fase declara hard gates que deben pasar antes de iniciar
la siguiente, y cada fase termina en un cierre determinista: el gate
`ci-pattern adoption check --phase N <repo>` = 0, ejecutado sobre el árbol real
y las instantáneas del host. El checklist es determinista: cada paso da el
comando exacto que se ejecuta y el gate que lo cierra; la IA no decide el orden
ni cuándo ha terminado. La regla que gobierna el checklist completo es una: la
adopción no se declara terminada antes de que un PR real recorra el ciclo
completo (fase 6).

Este documento es el pre-vuelo. Los pasos de ejecución de la adopción (auditar,
medir, instalar, validar) viven en §4 de la `SKILL.md` de esta skill; los
parámetros por repo, en `assets/parameters.md`.

Cada gate cita el incidente real de la adopción en Cadete (2026-09-30/10-01)
que lo justifica. El post-mortem blameless completo del proceso de orquestación
vive en el repo `ardelperal/APAP_WEB`:
`docs/postmortems/2026-09-30-ci-pattern-adoption-cadete.md`. Las citas usan el
identificador «C-N» definido en ese documento.

## Fase 0 — Inventario (solo lectura)

Fase de lectura pura: ningún cambio de código, workflow ni gobernanza.

- **G0.1 — Stack y toolchain documentados.** Language, test runner, linters y
  disponibilidad del coverage driver, con versión. Previene C7: la baseline de
  cobertura se midió con la toolchain local (11503) mientras el runner de CI
  medía 11559 (xdebug sin pinear); el gate quedó rojo sobre main limpia.
- **G0.2 — Inventario de gobernanza.** Protecciones de rama, rulesets,
  CODEOWNERS, plantillas de issue/PR y workflows de validación existentes, con
  la lista de eliminación confirmada por el operador. Previene C3: las
  prescripciones rotas (formato abreviado de issue, nombre de rama ilegal)
  aparecieron donde la gobernanza de destino no estaba instalada ni verificada.
- **G0.3 — Ramas y roles.** Qué rama es la base activa, qué ramas son de
  release o de integración y quién puede empujar a cada una.
- **G0.4 — Mecanismo de deploy.** Quién despliega producción y cómo: ¿el push a
  la rama base dispara el deploy o hay herramienta externa? Si no existe
  mecanismo de deploy, el binding evidencia↔deploy de la fase 4 no se instala.
- **G0.5 — Inventario de runners.** GitHub-hosted frente a self-hosted, y qué
  job corre en cada uno. Previene C2: el job de security corría en
  GitHub-hosted, no en la flota self-hosted; el diagnóstico «Docker daemon» ×3
  salió de grepear logs del runner equivocado en lugar de pedir la evidencia
  por paso con `gh api jobs/<id>`. **Censo de tooling (HR-16/HR-17 del
  runner-skill):** para cada job que vaya a migrar, la lista de binarios que
  invocan sus `run:` y contenedores, contrastada por runner del pool (fase 0,
  solo lectura) — alimenta el preflight fail-loud del job; un binario que
  solo existe por instalación manual en el runner histórico se declara
  ausente y aterriza en la definición del servicio, nunca en el host.
  Preflight canónico:
  `assets/tooling-preflight/tooling-preflight.sh` (lista declarada como dato
  versionado del job, `command -v` de cada binario, nombra TODOS los
  ausentes y falla si la lista no existe o está vacía).
- **G0.6 — Estado de release.** Qué está sin liberar en el repo. Nunca adoptar
  a mitad de un fix sin liberar: el ruido del cambio de gobernanza se mezcla
  con el fix pendiente y ambos quedan sin dueño claro.
- **G0.7 — Clones locales frente a workspaces de CI.** Identifique los
  workspaces de CI del repo y declárelos fuera de límites para cualquier
  edición. Previene C5: los workspaces de CI pertenecen al runner; editar en
  ellos compite con el runner sobre el mismo working tree.

Cierre determinista de fase 0: cada artefacto de gobierno del destino queda en
el inventario del contrato de adopción (`.github/ci-pattern-adoption.json`,
fase 0) con su destino —`adopted` hacia el artefacto del patrón que lo
sustituye (campo `target`, que debe pertenecer al catálogo canónico del
patrón: dato versionado en `assets/adoption/pattern-catalog.json` con las
ubicaciones de gobierno que este instala — incluidos los scripts de gate
consumer-side y el canal de propagación de skills — y los espacios
`ruleset:<nombre>`), o
`retired`—, sin tercera opción (HR-46), y `removal_confirmed: true` firma la
confirmación del operador. La fase 0 NO exige que los targets existan todavía:
la instalación y la eliminación del gobierno ajeno ocurren y se verifican en
la fase 3 (#306). La fase 3 y la fase 5 comprueban además que ninguna instrucción del consumer —`AGENTS.md`, el documento operativo y los ficheros que estos enlazan o citan— mande la documentación del proyecto a una ruta absoluta o externa al repositorio: la documentación vive dentro del repo, bajo `docs/`, y los post-mortems en `docs/postmortems/<AAAA-MM-DD>-<slug>.md` (HR-48, #319).

### Registro operativo del incidente (HR-21)

De un incidente de producción se guardan cuatro niveles, enlazados por
identificadores —ID de incidente, número de ticket, SHA, ejecución de CI,
issues, PRs y nombres de sistemas—: el **post-mortem** (parámetro 9, el qué
pasó y sus identificadores); el **registro operativo** (parámetro 24), con
el detalle técnico que hace falta para seguir atendiendo los tickets
abiertos —IPs, hostnames, comandos, logs, cronología fina y el siguiente
paso de cada ticket—; los **secretos** en el gestor declarado (parámetro
25); y los **datos personales** en el sistema de tickets, referenciados
por su número. Por defecto el registro operativo vive en el mismo
repositorio, en `docs/incidents/<AAAA-MM-DD>-<slug>.md`; un repositorio de
operaciones aparte es opcional y se declara en el parámetro 24.

**Migración de bitácoras locales.** Si la bitácora del incidente vive hoy en
una carpeta local fuera de cualquier repositorio, muévala a `docs/incidents/`
antes de cerrar la fase 5: lo que no está en el repositorio no existe para el
CI, ni para quien atiende la siguiente llamada sobre el ticket, ni para la IA.
La evidencia binaria —capturas, volcados, paquetes— se cita con su `sha256` y
su ubicación de acceso controlado, y lo redactado se marca como tal; nunca se
borra en silencio. Un repositorio de operaciones aparte solo se justifica si su
audiencia es distinta de la del código. Los SCRIPTS no son gobierno por ubicación sino por
REFERENCIA (#311): un script solo es gobierno si el gobierno lo invoca —las
fuentes declaradas en el catálogo (`referenced_scripts`: los `run:` de los
workflows de CI, los hooks y el entrypoint de paridad local HR-4 declarado en
su política) o su declaración explícita en el contrato del consumer
(`declared_governed_paths`). Un `scripts/*.ps1` operativo de la APP que nadie
del gobierno invoca NO se inventaría ni se retira: el glob genérico lo habría
forzado. La fase 0 ejecuta además la COBERTURA del inventario
(#312): escanea las rutas gobernadas del catálogo y las etiquetas del host
(instantánea JSON declarada en `host_labels`, sin red) y exige una entrada
por cada artefacto encontrado — una omisión es hallazgo que la nombra.

Cierre: `ci-pattern adoption check --phase 0 <repo>` = 0.

## Fase 1 — Verificación de realidad de la propagación

- **G1.1 — El mecanismo de propagación de skills funciona, verificado en
  vivo.** Audit el mecanismo real del destino: hooks activos, scope del
  reconciliador, overlays y markers. Ejecute o inspeccione el mecanismo; no se
  fíe de la documentación ni de la memoria de una sesión previa. Previene C4:
  la premisa de propagación era falsa (el hook post-commit era un no-op
  retirado y el reconciliador solo leía `skills/`), y se descubrió auditando
  después de haberla usado; los espejos quedaron stale durante el tramo.

Gate de salida de fase: una ejecución real del mecanismo que deja el artefacto
esperado en el destino, con la evidencia de esa ejecución declarada en el
contrato (fase 1).

Cierre: `ci-pattern adoption check --phase 1 <repo>` = 0.

## Fase 2 — Aislamiento del entorno

- **G2.1 — Clone de desarrollo fresco.** Todo el trabajo de adopción ocurre en
  un clone de desarrollo fresco, nunca en un workspace de CI. Previene C5.
- **G2.2 — Un worktree por actor concurrente.** Antes de tocar nada, declare
  cuántos actores van a trabajar en paralelo y asígneles un worktree cada uno.
  Previene C5: dos workers colisionaron en el working tree compartido de
  `apap-app` y uno commiteó sobre la rama del otro en pleno vuelo (HR-24).
- **G2.3 — Chequeo de envenenamiento por `.env`.** Compruebe si un `.env` en la
  raíz del repo filtra variables al entorno de tests; aísle la suite o
  documéntelo como premisa del entorno. Previene C6: el `.env` gitignored de la
  raíz filtró `APAP_*` a pydantic Settings y produjo 5 rojos ambientales que
  solo existían en la máquina de desarrollo (HR-28).
- **G2.4 — Toolchain pineada para toda baseline medida.** Toda cifra que un
  gate compare (cobertura, conteos, huellas) se mide con la misma toolchain que
  usará el runner de CI, pineada. Previene C7 (HR-26).
- **G2.5 — Baterías contra la revisión desplegada.** Toda batería contra un
  deploy corre desde un worktree en el SHA desplegado, nunca con los tests de
  la rama por defecto. Previene C1: la batería se lanzó con los tests de main
  contra la revisión desplegada y produjo 18 falsos fallos; el worktree en la
  revisión correcta tuvo que prepararlo la IA de otra sesión.

Gate de salida de fase: cada actor con su worktree, la suite verde sin el
`.env` local y la toolchain de medición pineada y documentada, con la evidencia
declarada en el contrato (fase 2).

Cierre: `ci-pattern adoption check --phase 2 <repo>` = 0.

## Fase 3 — Reemplazo de gobernanza

- **G3.1 — Reemplazo sin hueco.** La gobernanza vieja se retira y la nueva se
  instala en el mismo PR: nunca un PR que solo retira y otro posterior que
  instala, porque el hueco entre ambos es una ventana sin gates.
- **G3.2 — Plantillas canónicas.** La plantilla de PR del patrón (`.github/PULL_REQUEST_TEMPLATE.md` de `DysTelefonica/team-skills`) incluye la sección **Chain Context** (HR-54) además de «Tests que prueban el cierre» (HR-53): cópiela tal cual, porque el gate de contrato de PR (`assets/pr-contract/check_pr_contract.py`) la lee como dato; el formulario de issue sale de `.github/ISSUE_TEMPLATE/`.  Issue y PR con las secciones exactas que los
  gates leen y con el contrato de revisión visible (presupuesto, excepción como
  campo de datos, cadena de PRs).
- **G3.3 — Aprobador nombrado.** El operador que aprueba issues y revisa PRs
  queda nombrado en la gobernanza instalada; un pipeline sin aprobador named
  detiene cada issue en la puerta.

Previene C3 en los tres gates: el formato abreviado de issue costó reruns, el
nombre de rama ilegal violó el validador del destino y las superficies con
prosa entre paréntesis fueron rechazadas por el contrato del worker.

El PR de sustitución retira o sustituye TODOS los artefactos inventariados en
la fase 0 en el mismo PR: nada del gobierno original sobrevive, ni en parte ni
«en paralelo» (HR-46). El gate de fase 3 no acepta un fichero de evidencia:
ejecuta la comprobación sobre el árbol real y sobre las instantáneas del host.

1. Genere las instantáneas del host con GETs de solo lectura (la misma
   mecánica de `assets/host-readback/`): protección de la rama y rulesets en
   JSON, y declárelas en `host_snapshots` del contrato. Si el repo es
   privado sin plan, guarde el cuerpo real del `403` («Upgrade to GitHub
   Pro…») como instantánea y declare esa área en `host_capabilities` del
   contrato del host como `unavailable`; el readback lo registra como
   evidencia de la capacidad, no como error.
2. Copie el contrato del host del patrón en `.github/host-contract.json`:
   los rulesets del patrón se leen de ahí, no de una declaración propia en
   el contrato de adopción (la identidad del patrón no es autodeclarada).
3. Declare en `pattern_paths` los artefactos que el patrón instala. El gate
   verifica EN PROCESO el manifiesto de gobierno `.governance-manifest.json`
   (existencia y sha256 de cada fichero, con las mismas funciones de
   `ci-pattern verify`): todo `pattern_path` debe estar cubierto por el
   manifiesto y todo artefacto de gobernanza del árbol o de las
   instantáneas — workflows, CODEOWNERS, dependabot, `.pre-commit-config.yaml`,
   `.husky/`, `lefthook.yml`, `.github/*.json` — que no esté en el manifiesto
   o inventariado es residual y el gate lo nombra uno a uno.
4. Registre la evidencia del PR de sustitución en
   `docs/adoption/governance-replacement-pr.md` y declárela en `evidence.path`
   de la fase: el gate también la comprueba.

Cierre: `ci-pattern adoption check --phase 3 <repo>` = 0.

## Fase 4 — Contratos cableados

- **G4.1 — Issue-spec equivalente** instalado y validando el formato canónico
  de issue en el destino.
- **G4.2 — Presupuesto de tamaño** con excepción declarada como campo de datos
  (`size-exception-reason:`) y cadena de PRs para lo que no cabe.
- **G4.3 — Preflight ejecutable** en local con un solo comando, paridad con el
  job de lint (HR-4).
- **G4.4 — CI en la rama base**, net-new si el destino no tenía.
- **G4.5 — Binding evidencia↔deploy solo si existe deploy.** El estado de
  commit por SHA (HR-10) se instala únicamente cuando la fase 0 inventarió un
  mecanismo de deploy; sin deploy no hay evidencia que registrar y el gate solo
  añade ruido. Verificación de salida de fase: el evaluador resuelve la
  revisión previa desde la fuente desplegada declarada (HR-40) y su política
  de exenciones bootstrap es simétrica entre los contextos requeridos
  (HR-41) — nunca por heurísticas de listas de runs ni por exención ad-hoc
  de un solo contexto.

Previene C3 en G4.1-G4.3: la receta de trabajo mencionaba una ruta de tests
inexistente (`tests/test_required_jobs.py`) y el repo equivocado para la wave
de #1160 (dijo cadete, los ficheros eran de APAP_WEB). Verifique en vivo cada
ruta, rama y repo que una prescripción mencione antes de delegar (HR-25).

Gate de salida de fase: cada contrato ejecuta al menos una vez contra datos
reales del destino (una issue canónica valida, un PR de prueba mide, el
preflight corre completo en local).

Cierre: `ci-pattern adoption check --phase 4 <repo>` = 0 (con los gates
declarados en `gates`, cubriendo exactamente el conjunto del patrón).

### Camino runner — repos privados sin protección de rama

Un repo privado con plan gratuito no aplica protección de rama, rulesets ni
checks requeridos (la API responde `403` «Upgrade to GitHub Pro…»). En ese
host el patrón se adopta con el camino runner (HR-34 con su clase
`runner-enforced`, HR-48 a HR-52):

1. Declare en `.github/host-contract.json` las áreas del host en
   `host_capabilities` como `unavailable`, la etiqueta del runner en
   `runner.label` —la que incluye el `runs-on`, p. ej. `cadete`— y las
   ramas gobernadas en `protected_branches`.
2. Capture las instantáneas con GETs de solo lectura: para
   `branch-protection` y `rulesets` guarde el cuerpo real del `403`, y para
   `actions-runners` la respuesta de `repos/<owner>/<repo>/actions/runners`.
3. Instale los tres controles (`assets/runner-controls/`) y cablee un
   workflow por control en el runner declarado: detective post-push (HR-49),
   merge gobernado (HR-50) y guard de force-push (HR-51), cada uno en un job
   cuyo `runs-on` incluya `runner.label`.
4. Añada el gate `runner-binding` a `gates` de la fase 4 del contrato de
   adopción con `--contract @.github/host-contract.json` y
   `--workflows @.github/workflows`; la fase 4 lo exige cuando el host
   declara camino runner (HR-52) y verifica además, contra la instantánea
   `actions-runners`, que el runner declarado está `online`.
5. Complete las fases 3 y 4 del camino normal: las relecturas `403` valen
   como evidencia de `unavailable` gracias al detector compartido
   (`assets/host_plan.py`), nunca como error.

Cierre: `ci-pattern adoption check --phase 4 <repo>` = 0 — con el binding
en verde y el runner online; si el runner se cae, la adopción queda
bloqueada nombrando `runner.label`.

## Fase 5 — Documentación operativa generada en destino

- **G5.1 — Documento GENERADO, nunca escrito a mano.** El consumer recibe su
  flujo operativo (issues → ramas → commits → PR → checks → merge → release →
  fricciones) generado con `ci-pattern adoption generate-doc <repo> --out
  docs/ci-pattern-flow.md` a partir de SUS datos: `ci-pattern.yaml`,
  `.github/host-contract.json`, la política de required-jobs y los formularios
  de issue. Un doc a mano contradice la configuración real en la primera
  deriva.
- **G5.2 — Enlazado desde `AGENTS.md`.** La ruta del doc aparece enlazada en
  `AGENTS.md` (o el fichero declarado como `linked_from`): una IA que llega al
  repo encuentra el flujo sin descubrirlo por el camino.
- **G5.3 — Gate de deriva.** `ci-pattern adoption check --phase 5` regenera el
  doc en el mismo proceso y lo compara por secciones: etiqueta documentada
  distinta del host-contract, check requerido distinto de la política o regex
  de rama distinta de la real son hallazgos que nombran la discrepancia. El
  mismo gate corre en el CI del consumer después de la adopción.

- **G5.4 — Bloque canónico en `AGENTS.md` (#305).** Copie al final de
  `AGENTS.md` el bloque `agents-block` del doc generado. COPIA byte a byte,
  nunca redacción propia: una copia alterada es hallazgo del gate.
- **G5.5 — Estructura del `AGENTS.md` (HR-55, #341).** Tras la propagación, el
  `AGENTS.md` queda como título + bloque `personal-skills:slice:*` (sha idéntico
  a `slice_block_sha256` del `.team-skills.yaml`) + como máximo UNA sección
  local. Lo escrito a mano que hoy vive fuera de esas piezas se mueve a
  `docs/` por materia (hechos, tablas, walkthroughs) o a la sección local
  (reglas y enlaces del repo); nada queda «entre marcas» que la propagación no
  haya puesto. El gate se activa con la declaración explícita
  `agents_md: "propagated"` en el contrato de adopción (sin ella informa y no
  exige, sin romper la adopción previa) y las fases 3 y 5 lo verifican con
  `_check_agents_structure`; el doctor de flota exige además el bloque a todo
  consumer con `governance: ci-pattern`, aunque no lo declare.

Previene la fricción de la sesión de gobierno del 2026-10-05: los agentes
dedujeron sobre la marcha las reglas de cadenas (`Closes #N` en prosa cerró
#194), etiquetas (se creó una etiqueta inexistente) y el impacto de la matriz.

Gate de salida de fase: el doc operativo generado, enlazado y en verde en el
gate de deriva sobre el árbol real del consumer.

## Fase 6 — Primer PR real como test de aceptación

- **G6.1 — El primer PR real recorre el ciclo completo.** Issue canónica,
  worktree dedicado, preflight completo, PR con etiquetas en el comando de
  creación, CI verde, merge. La adopción no se declara terminada antes de ese
  PR; ningún ensayo con PRs de prueba lo sustituye.
- **G6.2 — El patrón completo viaja en la skill antes de la adopción.** Todo
  sub-patrón que el destino necesite (gates dormibles incluidos) debe estar en
  la skill antes de empezar. Previene C8: el patrón dormible (R15) no estaba en
  la skill cuando Cadete lo necesitó y llegó a mitad de vuelo desde el benchmark
  de gentle-ai.

Gate de salida de fase: el primer PR real mergeado bajo el pipeline nuevo y el
veredicto de ese ciclo registrado como evidencia de la adopción (fase 6, con
número, SHA de 40 hex y `conclusion: success`).

Cierre: `ci-pattern adoption check --all <repo>` = 0.

## Trazabilidad fase → incidente

| Fase | Incidente que la fase cierra |
|---|---|
| 0 | C1 (batería desfasada), C2 (runner equivocado), C7 (toolchain sin pinear) |
| 1 | C4 (premisa de propagación falsa) |
| 2 | C1 (worktree de revisión), C5 (colisión de workers), C6 (.env), C7 (baseline) |
| 3 | C3 (prescripciones rotas de gobernanza) |
| 4 | C3 (rutas y repos no verificados) |
| 5 | descubrimiento sobre la marcha del flujo (sesión de gobierno 2026-10-05) |
| 6 | C8 (patrón dormible ausente) y la cadena completa como test de aceptación |
