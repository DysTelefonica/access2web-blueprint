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
  ausentes —una línea por ausente, `falta <cli> en el runner <nombre>`— y
  falla si la lista no existe, no es legible o está vacía). La identidad del
  runner es la opción opt-in `--runner <nombre>`: sin ella el preflight
  nombra la lista de ausentes como antes; con ella nombra cada ausente, en
  qué runner falta y cómo se instala. Desde #1172 el preflight distingue dos
  fallos que antes compartían salida: **binario ausente en el runner → AVISO**
  (sale 0: el fix aterriza en la definición del servicio del runner, no en el
  PR que se está abriendo) y **contrato roto** —lista vacía, ilegible o
  ausente, sin `--binaries`— → **bloqueo** (sale 2). El cableado por job es
  la sección «Cableado del preflight de tooling por job (G0.5)» de esta misma
  fase.
- **G0.5b — El job `gates` en sombra (fase 1 de #1179).** Un job que ejecuta
  los gates declarados del repo una vez y emite un **veredicto único** —cuántos
  gates están en verde y, para cada uno que falla, su causa— en vez de N checks
  que se ponen en rojo por el mismo defecto. Va **en sombra**: publica su check
  y NADIE lo exige; no cambia `required_jobs` ni la protección de rama. Su
  veredicto se compara con el rollup actual para decidir la fase 2 (criterio C6
  de #1179). El asset es `assets/gates/run_gates.py` y la lista de gates es un
  fichero versionado en el repo del consumer (`.github/gates.txt`, una línea por
  gate: `nombre: comando`).

  ```yaml
  jobs:
    gates:
      name: gates (sombra, no exigido)
      # Único evento que evalúa el PR (HR-30).
      if: github.event_name == 'pull_request'
      runs-on: [self-hosted, Linux, mi-consumer]
      # El techo del job cubre el gate MÁS LENTO (los gates corren en paralelo).
      timeout-minutes: 20
      permissions:
        contents: read
      steps:
        - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
          with:
            persist-credentials: false

        - name: Veredicto único de los gates (sombra)
          run: |
            python3 .agents/skills/ci-pattern/assets/gates/run_gates.py \
              --gates .github/gates.txt
  ```

  La lista de gates:

  ```text
  # .github/gates.txt — los gates del repo, uno por línea: `nombre: comando`.
  unit: bash testing/run-unit-suites.sh
  lint: bash scripts/lint.sh
  ```

  Tres reglas: el job **no** entra en `required_jobs` ni en el `needs:` del
  agregador (si entró, dejaría de ser sombra); su nombre de check lleva
  «(sombra)» para que nadie lo confunda con un check exigido; y **duplica el
  trabajo de los gates**, que es el precio de la sombra y el motivo por el que
  es temporal.
- **G0.6 — Estado de release.** Qué está sin liberar en el repo. Nunca adoptar
  a mitad de un fix sin liberar: el ruido del cambio de gobernanza se mezcla
  con el fix pendiente y ambos quedan sin dueño claro.
- **G0.7 — Clones locales frente a workspaces de CI.** Identifique los
  workspaces de CI del repo y declárelos fuera de límites para cualquier
  edición. Previene C5: los workspaces de CI pertenecen al runner; editar en
  ellos compite con el runner sobre el mismo working tree.

- **G0.6 — La página «dónde trabajo» (#1201).** El consumer escribe
  `docs/codebase/where-to-work.md` con una fila por tarea canónica en la forma
  «lee → toca → prueba» y una fila `local` con su índice y sus skills
  obligatorias. El asset `assets/where-to-work/check_where_to_work.py` mide
  **presencia y deriva**: que la página exista, que tenga las filas `form`,
  `ci`, `schema` y `local`, que cada fila declare las tres columnas y que cada
  ruta citada exista en el repo. No juzga la prosa. La plantilla está en
  `assets/where-to-work/templates/where-to-work.md`.

  ```yaml
  - name: La página dice dónde trabajar (#1201)
    run: >
      python3 .agents/skills/ci-pattern/assets/where-to-work/check_where_to_work.py .
  ```

### Cableado del preflight de tooling por job (G0.5)

El censo de la fase 0 no se queda en un inventario: se cablea. TODO job
migrado —TODOS, no solo el primero— lleva el preflight como primer paso
ejecutable después del checkout, con la lista de CLI del job versionada en el
repo del consumer y la identidad del runner que lo ejecuta:

```yaml
jobs:
  unit:
    runs-on: [self-hosted, Linux, ARM64, apap, oracle, coolify, noble, deploy]
    steps:
      - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
        with:
          persist-credentials: false

      # Primer paso EJECUTABLE del job: el contrato de tooling se prueba en el
      # runner que ejecuta este job (regla 16 del runner-skill), no en el
      # runner histórico.
      - name: Preflight del contrato de tooling del job
        run: |
          sh .agents/skills/ci-pattern/assets/tooling-preflight/tooling-preflight.sh \
            --binaries .github/tooling/unit.txt \
            --runner "${{ runner.name }}"

      - name: Suite del job
        run: make test
```

Y la lista del job, versionada en el repo del consumer: una CLI por línea,
`#` abre comentario y la línea en blanco se ignora. Una entrada puede
declarar **cómo se instala** con el comentario `# instalar: <comando>` (o
`# install:`): el aviso del binario ausente lo nombra, y sin esa declaración
el aviso remite a la definición del servicio del runner (#1172).

```text
# .github/tooling/unit.txt — CLI que el RUNNER debe aportar a este job.
python3
make
ruff # instalar: pip install ruff==0.15.21
```

Cuatro reglas del cableado:

1. **La ruta es la del consumer propagado**:
   `.agents/skills/ci-pattern/assets/tooling-preflight/tooling-preflight.sh`.
   Es la copia que la propagación deja en el repo del consumer; el asset del
   catálogo (`personal/ardelperal/ci-pattern/assets/tooling-preflight/`) no
   se cablea en un workflow.
2. **El cableado canónico pasa siempre `--runner "${{ runner.name }}"`.** Sin
   la identidad del runner, «falta <cli>» no dice dónde mirar: el fix aterriza
   en la definición del servicio del runner concreto, nunca en una
   instalación manual del host. Con identidad, el preflight nombra cada
   ausente, uno por línea: `falta <cli> en el runner <nombre>`. La opción es
   opt-in (una invocación heredada `--binaries FICHERO` a secas sigue siendo
   válida y conserva su texto anterior), pero el cableado nuevo la declara
   siempre: es lo que cierra el hueco.
3. **La lista declara lo que el runner aporta, no lo que instala el job.** La
   lista son las CLI que el host del runner debe tener por definición de
   servicio. Una CLI que el job instala con un `setup-*` (node, python) la
   aporta el propio job y falla sola si el `setup-*` no la deja; no va en la
   lista. Si un consumer quiere verificarla también, pone el paso del
   preflight justo después del `setup-*` —sigue siendo el primer paso
   ejecutable del job— y la incluye.
4. **La lista es dato, no script.** Se versiona junto al workflow del job
   (`.github/tooling/<job>.txt`); el preflight la lee con `command -v` y
   nombra los ausentes por stderr. Con `--runner` cada ausente sale con su
   instrucción de instalación (la que declara la entrada con
   `# instalar: <comando>`, o la definición del servicio del runner): el
   aviso es accionable sin que quien abre el PR tenga que buscar el fix.

APAP (`ardelperal/APAP_WEB`), BP (`DysTelefonica/access2web-blueprint`) y
cadete (`DysTelefonica/cadete`) copian este patrón en SUS jobs, con SU lista
versionada: este catálogo publica el patrón y no edita ningún
`.github/workflows` —el workflow de cada consumer es de su propietario—, así
que cada consumer abre su propio PR con su cableado.

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

> **Artefactos de gobierno en `phases[0].inventory` — recordatorio (oleada 8).**
> El catálogo declara como artefactos de gobierno todos los caminos gobernados
> del repo: workflows, plantillas de issue/PR, `AGENTS.md`, `ci-pattern-flow.md`,
> `CODEBASE-GUIDE.md`, etc. `CONTRIBUTING.md` también: cualquier consumer que
> lo lleve debe declararlo en `phases[0].inventory` con
> `{artifact: "CONTRIBUTING.md", disposition: "adopted", target: "CONTRIBUTING.md"}`
> para silenciar el hallazgo `fase 0: artefacto de gobierno sin entrada de
> inventario: CONTRIBUTING.md`. **HR-71 (fase 5) no basta sola**: la fase 0 lee
> solo del inventario del contrato, no del manifiesto de docs. Lo mismo aplica
> a cualquier artefacto nuevo que el consumer incorpore y que el catálogo
> reconozca como gobernado (consulte `assets/adoption/pattern-catalog.json`).

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
- **G2.6 — Lo que la fase 2 mide lo declara la base, si la hay.** Con
  `--seal-baseline <dir>` la fase 2 exige que el allowlist y el árbol que
  escanea el gate HR-28 sean los que mide la base (operador de identidad, ver
  la fase 3): el PR no puede apartar una fuga del gate cambiando lo que el
  gate lee. Si la ruta de la base ya no existe en la cabeza, solo se admite un
  movimiento puro: la ruta nueva tiene el mismo contenido que la base, byte a
  byte. Un movimiento con edición en el mismo PR da hallazgo. Sin la opción,
  la fase 2 se mide como siempre.
- **G2.7 — Con la misma ruta, el contenido tampoco mide menos (#1341).** La
  identidad de la ruta no dice nada de lo que hay dentro: un PR que reescribe
  el allowlist en su sitio (una variable de entorno nueva, una extensión
  excluida nueva) se lleva la medición sin que nada lo detecte. Con
  `--seal-baseline`, el allowlist no puede autorizar de más y la instantánea
  de etiquetas del host no puede dejar de cubrir las que cubría. Medir más se
  admite. Una ampliación del allowlist se admite si el allowlist la declara
  como **exención con identidad** (HR-33): la superficie, el elemento y quién
  la autoriza (`actor` o `issue`). Sin identidad la exención no exime de
  nada: nadie puede responder por ella.

### Camino Windows/WSL — worktree y `git`

Un worktree creado desde Windows guarda en su `.git` (un fichero, no un
directorio) la ruta absoluta del repositorio, con la forma `gitdir: C:/...`.
Esa ruta no existe dentro de WSL, donde el mismo sitio aparece como `/mnt/c/...`,
y `git` en WSL falla al leer ese worktree.

1. Elija un lado y no lo mezcle: o el worktree y el preflight son de Windows
   (Git for Windows y Python de Windows), o son de WSL (un worktree creado
   dentro del sistema de ficheros de WSL, p. ej. bajo `/home/<usuario>/`).
2. Cree el worktree desde ese lado con su propio comando de Git. Nunca copie
   el `.git` de un lado a otro.
3. Ejecute el preflight desde la raíz de ese worktree con el intérprete del
   mismo lado.

**Limitación conocida:** el comportamiento de `git` en WSL con un worktree
creado en Windows no está verificado por este patrón. Si el preflight falla en
WSL por la ruta de `.git`, la salida práctica es el punto 1, no un parche manual
de ese fichero.

Gate de salida de fase: cada actor con su worktree, la suite verde sin el
`.env` local y la toolchain de medición pineada y documentada, con la evidencia
declarada en el contrato (fase 2).

Cierre: `ci-pattern adoption check --phase 2 <repo>` = 0.

### Control local del checkout atrasado (#914)

El control no es un job de CI ni un hook de push: el CI siempre hace checkout fresco, y un bloqueo en `pre-push` pararía cada push de rama de PR cada vez que `main` avanza. Vive en el arranque del partial `web`, justo después de `git pull --ff-only`:

1. **Ejecutar.** `ci-pattern checkout-sync --repo <checkout>`. Códigos: 0 sincronizado; 1 atrasado o rama equivocada (seguir la instrucción que imprime); 2 no puede medir (`git fetch origin` y repetir).
2. **Linux y macOS.** El script es POSIX `sh`. La fecha de la última sincronización se lee con `stat -c` (GNU) o `stat -f` (BSD); el script prueba ambos.
3. **Windows.** Ejecutar desde Git Bash o desde WSL sobre el mismo checkout. `FETCH_HEAD` es por worktree: un worktree que nunca ha hecho `git fetch origin` sale con 2 aunque sus refs estén al día.

## Fase 3 — Reemplazo de gobernanza

- **G3.1 — Reemplazo sin hueco.** La gobernanza vieja se retira y la nueva se
  instala en el mismo PR: nunca un PR que solo retira y otro posterior que
  instala, porque el hueco entre ambos es una ventana sin gates.
- **G3.2 — Plantillas canónicas.** La plantilla de PR del patrón (`.github/PULL_REQUEST_TEMPLATE.md` de `DysTelefonica/team-skills`) incluye la sección **Chain Context** (HR-54) además de «Tests que prueban el cierre» (HR-53): cópiela tal cual, porque el gate de contrato de PR (`assets/pr-contract/check_pr_contract.py`) la lee como dato; el formulario de issue sale de `.github/ISSUE_TEMPLATE/`.  Issue y PR con las secciones exactas que los
  gates leen y con el contrato de revisión visible (presupuesto, excepción como
  campo de datos, cadena de PRs).
  La sección HR-53 tiene dos formas y el gate acepta las dos: con test, el
  fichero o el caso que demuestra el cierre; sin test, el encabezado estable
  «No aplica» y, debajo, la ruta versionada en git de la evidencia que cubre el
  cierre (un ADR, un registro de decisión, la página de documentación). La
  evidencia no tiene que ser un fichero de test. «No aplica» solo vale si todos
  los ficheros cambiados son no ejecutables (por defecto documentación:
  `**/*.md`, `docs/**`, `.github/ISSUE_TEMPLATE/**`,
  `.github/PULL_REQUEST_TEMPLATE.md`); un workflow, un script o cualquier
  código exige test, y el hallazgo nombra el fichero. Un consumer que quiera
  otros patrones los declara en `closure_not_applicable_paths` de su política.
  Ninguna etiqueta exonera la sección: `type:docs` no exonera nada (#911, #918).
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
   (existencia y sha256 gobernado de cada fichero, con las mismas funciones de
   `ci-pattern verify`): todo `pattern_path` debe estar cubierto por el
   manifiesto y todo artefacto de gobernanza del árbol o de las
   instantáneas — workflows, CODEOWNERS, dependabot, `.pre-commit-config.yaml`,
   `.husky/`, `lefthook.yml`, `.github/*.json` — que no esté en el manifiesto
   o inventariado es residual y el gate lo nombra uno a uno. Los workflows se
   sellan NORMALIZADOS (`schema_version: 3`, #1199 y #1361): el hash es el del
   contenido con cada pin de `uses:` y cada `image:` con digest a un marcador
   fijo, de modo que un bump de Dependabot no los desalinea y un cambio de lógica
   sí. Lo que no se normaliza (un `docker run` dentro de `run:`, una `image:`
   entre comillas, un `container:` en forma corta) falla cerrado: su bump exige
   re-sellar. El hash se calcula con el algoritmo del formato que declara el
   manifiesto: el canon vigente mide el formato 2 y el 3, cada uno con el suyo
   (#1371), porque la base del consumer mide el PR con su canon y su formato
   puede ser el anterior.
4. Registre la evidencia del PR de sustitución en
   `docs/adoption/governance-replacement-pr.md` y declárela en `evidence.path`
   de la fase: el gate también la comprueba.
5. **Cablee la fase 3 dentro del PR (HR-72).** Sin este paso la fase 3 se mide
   fuera del pull request y un fichero gobernado puede llegar a `main` sin
   re-sellarse. Declare en el contrato de adopción:

   ```json
   "governance_seals": { "job": "governance-seals" }
   ```

   y añada el job al `ci.yml` del consumer, con esta forma:

   ```yaml
   name: pr-gates

   on:
     pull_request_target:
       types: [opened, edited, labeled, unlabeled, synchronize, reopened]

   permissions:
     contents: read
     pull-requests: read
     checks: write

   concurrency:
     group: pr-gates-${{ github.event.pull_request.number }}
     # Solo un commit nuevo cancela una corrida (#978).
     cancel-in-progress: ${{ github.event.action == 'synchronize' }}

   jobs:
     canon:
       name: canon
       # El UNICO job que puede leer el token del catalogo (#1270). No ejecuta
       # codigo del PR: el checkout es de la BASE, de donde salen la accion y el
       # sello, asi que el PR no puede tocar la una ni elegir el commit que se
       # mide. El patron se publica como artefacto para los jobs del gate.
       runs-on: [self-hosted, Linux, mi-consumer]
       timeout-minutes: 10
       steps:
         - name: Check out the base (trusted action and seal)
           uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4
           with:
             ref: ${{ github.event.pull_request.base.sha || github.sha }}
             persist-credentials: false

         # La accion y el sello son los de la base. El token vive SOLO en este paso.
         - name: Resolve the canon at the seal
           uses: ./.github/actions/pattern-at-pin
           with:
             token: ${{ secrets.TEAM_SKILLS_READ_TOKEN }}

         # El canon viaja como tarball: un artifact no conserva el bit de
         # ejecucion, y el canon lo necesita (`assets/bin/ci-pattern`).
         - name: Pack the canon
           run: |
             set -euo pipefail
             tar -czf "$RUNNER_TEMP/canon.tar.gz" -C "$PATTERN_DIR" .

         - name: Publish the canon
           uses: actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a # v7.0.1
           with:
             name: canon
             path: ${{ runner.temp }}/canon.tar.gz
             if-no-files-found: error

     governance-seals:
       name: governance-seals
       # HR-72 (#1027). El gate evalua el PR, asi que corre en `pull_request_target`,
       # que ejecuta el workflow de la base (#1270).
       needs: canon
       runs-on: [self-hosted, Linux, mi-consumer]
       timeout-minutes: 15
       permissions:
         contents: read
         checks: write
       env:
         # El canon publicado por el job `canon`.
         PATTERN_DIR: .canon
       steps:
         - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
           with:
             ref: ${{ github.event.pull_request.head.sha || github.sha }}
             path: pr-head
             fetch-depth: 0
             persist-credentials: false

         # La CONFIGURACION que gobierna el gate sale de la BASE (#1270): el
         # CONTRATO DE ADOPCION se lee de ahi, sin credenciales. Un PR no puede
         # reescribir el contrato que lo mide. El estado (el manifiesto de
         # gobernanza, los ficheros gobernados y las instantaneas del host)
         # sigue siendo el del PR (#1351).
         - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.2.2
           with:
             ref: ${{ github.event.pull_request.base.sha }}
             path: base
             persist-credentials: false

         - name: Get the canon
           uses: actions/download-artifact@37930b1c2abaa49bbe596cd826c3c89aef350131 # v7.0.0
           with:
             name: canon
             path: .canon

         - name: Unpack the canon
           run: |
             set -euo pipefail
             tar -xzf .canon/canon.tar.gz -C .canon

         # UNA sola forma de medir (#1339): el gate corre una vez y la base es
         # el suelo. El PR puede AÑADIR (el artefacto de gobierno que acaba de
         # crear, con su entrada de inventario y su sello) y no puede QUITAR
         # (lo que la base sellaba no se reduce mientras el fichero siga en el
         # arbol). Antes el fragmento llevaba DOS invocaciones y su
         # combinacion ponia en rojo un PR legitimo (E9).
         - name: Governance identity (HR-46, HR-72)
           run: python3 "$PATTERN_DIR/assets/bin/ci-pattern" adoption check --phase 3 pr-head --config-from base

         # Un run de `pull_request_target` se atribuye a la BASE: su check run no
         # aparece en el PR ni satisface un check requerido. El veredicto se publica
         # como CHECK RUN sobre la CABEZA del PR, que es lo que bloquea el merge
         # (`checks: write`).
         - name: Publish the verdict as a check run on the PR head
           if: always()
           env:
             HEAD_SHA: ${{ github.event.pull_request.head.sha }}
             GH_TOKEN: ${{ github.token }}
             CONTEXT: governance-seals
           run: |
             set -euo pipefail
             if [ "${{ job.status }}" = "success" ]; then conclusion=success; else conclusion=failure; fi
             printf '{"head_sha": "%s", "name": "%s", "status": "completed", "conclusion": "%s"}' \
               "$HEAD_SHA" "$CONTEXT" "$conclusion" > "$RUNNER_TEMP/check-run.json"
             gh api --method POST "repos/$GITHUB_REPOSITORY/check-runs" --input "$RUNNER_TEMP/check-run.json"

         # #1172 (decisión del dueño, 2026-10-24): cuando el sello no coincide,
         # el mensaje del gate da el hash esperado y el comando de re-sellar.
         # El job deja además el manifiesto RE-SELLADO como artefacto, para que
         # el fix sea un fichero y no una instrucción. Sin push automático: el
         # manifiesto re-sellado no se commitea ni se empuja desde el job.
         # `if: failure()` corre solo cuando el gate falló, y el paso no usa
         # `|| true` ni `continue-on-error`: el veredicto del gate no se traga.
         - name: Manifiesto re-sellado (sin push)
           if: failure()
           env:
             PATTERN_DIR: pattern
           run: |
             python3 "$PATTERN_DIR/assets/bin/ci-pattern" manifest seal base \
               --out "$RUNNER_TEMP/governance-manifest.resealed.json"

         - name: Publicar el manifiesto re-sellado como artefacto
           if: failure()
           uses: actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a # v7.0.1
           with:
             name: governance-manifest-resealed
             path: ${{ runner.temp }}/governance-manifest.resealed.json
             if-no-files-found: warn
```

   El commit del patrón **no se写在**: se lee del sello `catalog @ <sha40>`
   de la primera línea de `.team-skills.yaml`, que el propagador reescribe en
   cada corrida (#1062). Un `pin:` literal queda viejo en la siguiente
   propagación y el `diff` que produce no dice nada.

   El fragmento es un **workflow completo** en `pull_request_target` (#1270): el
   job `canon` resuelve el patrón y lo publica como artefacto, y los jobs del
   gate lo consumen sin ningún secreto. El checkout del job `canon` es de la
   BASE, así que la acción y el sello son los de la base: el PR no puede tocar
   la una ni apuntar el sello a un commit más antiguo del catálogo. El patrón
   viaja como tarball porque un artefacto no conserva el bit de ejecución.

   El veredicto se publica como check run sobre la cabeza del PR: un run
   de `pull_request_target` se atribuye a la base y su check run no bloquea el
   merge.

   La CONFIGURACION que gobierna el gate sale de la base (`--config-from base`):
   el **contrato de adopción** se lee del checkout de la base, sin
   credenciales. Un PR no puede reescribir el contrato que lo mide: si lo
   cambia, el gate sigue midiendo con el de la base, y el cambio solo surte
   efecto al llegar a `main`. El ESTADO —el manifiesto de gobernanza, los
   ficheros gobernados y las instantáneas del host— sigue siendo el del PR.

   El `host-contract` y el `ci-pattern.yaml` NO salen de la base (#1351): son la
   declaración del consumer sobre su host y sus parámetros, y miden contra las
   instantáneas que el PR trae. Medir el contrato del host con el de la base
   haría que un PR legítimo quedara rojo para siempre: el que declara un
   ruleset propio del host —el único canal que el canon ofrece para el
   gobierno del host que el consumer conserva— y el que corrige una lectura
   obsoleta del plan del host. Su cambio surtiría efecto al llegar a `main`,
   con el veredicto, y el PR quedaría bloqueado.

   Lo que el PR no puede es apartar la medición: el INVENTARIO del patrón (que
   con `--config-from` sale de la base) manda sobre la lista de exenciones del
   `host-contract`. Un ruleset que el inventario declara retirado no se exime
   por declararlo en el contrato del host: la sustitución del patrón tiene que
   verse en el host, y el gate lo nombra mientras siga vivo.

   Con la base como suelo, el PR puede AÑADIR y no puede QUITAR (#1339). El
   fragmento mide el gate UNA vez: antes llevaba dos invocaciones
   (`--config-from base` y `--seal-baseline base`) y su combinación ponía en
   rojo un PR legítimo, porque la que gobierna con el contrato de la base no
   puede conocer el artefacto de gobierno que el PR acaba de crear (E9).
   Medir dos veces con dos reglas no es medir más: es medir con dos
   configuraciones, y el veredicto es el peor de los dos.

   - **AÑADIR**: lo que el PR declara de más se mide con la unión de la base y
     la cabeza. El inventario de la fase 0 y las rutas gobernadas declaradas se
     unen, y los `pattern_paths` de la fase 3 son los que declara el PR: el
     artefacto de gobierno que el PR acaba de crear, con su entrada de
     inventario y su sello, es un PR legítimo y sale en verde (E9). Retirar un
     workflow también se permite: si el fichero ya no existe, no hay nada que
     sellar.
   - **QUITAR**: la base manda en el conflicto. La entrada de inventario que
     declara la base sigue siendo la que mide (un PR no puede reapuntarla
     para que un residuo deje de nombrarse) y lo que la base sellaba no se
     reduce en un PR mientras el fichero siga en el árbol (E8b): el PR que
     deja de declararlo recibe un hallazgo, y el que lo retira de verdad
     (borrando el fichero) sale en verde.

   `--seal-baseline <dir>` sigue siendo la opción para el consumer que quiere
   que gobierne el contrato de SU CABEZA, con el de la base como suelo por
   fases. Es una alternativa a `--config-from`, no un añadido: el canon mide
   con `--config-from` y no con las dos.

   La BASE DE SELLADO no se reduce en un PR (#1317). Con `--seal-baseline <dir>`
   la fase 3 mide el conjunto que hay que sellar como la UNIÓN de los
   `pattern_paths` de la base y los de la cabeza: lo que la base sellaba no
   puede desaparecer de un PR mientras el fichero siga en el árbol, y lo nuevo
   de la cabeza también se sella. Retirar un workflow sí se permite: si el
   fichero ya no existe, no hay nada que sellar y la base deja de exigirlo. Sin
   la opción el gate se mide como siempre, y una base sin contrato legible
   falla cerrado (exit 2).

   La misma opción mide las fases 0 y 2 contra la base (#1319), con el
   operador de IDENTIDAD: la instantánea de etiquetas del host
   (`host_labels`, fase 0) y el allowlist y el árbol que escanea el gate de
   aislamiento (`env_isolation.allowlist` y `env_isolation.root`, fase 2)
   tienen que ser los que mide la base. El motivo es el mismo que en la fase
   3: el PR no puede apartar el gate cambiando lo que el gate lee — con un
   allowlist propio o con un árbol más corto, una fuga deja de medirse y la
   fase sale en verde. Si la ruta de la base ya no existe en la cabeza, solo se
   admite un MOVIMIENTO PURO: la ruta nueva tiene el mismo contenido que la base,
   byte a byte (en un directorio, cada fichero de la base con el mismo contenido;
   la cabeza puede añadir ficheros, que solo añaden medición). Un movimiento
   con edición en el mismo PR da hallazgo: es el hueco que cerraba la excepción
   de retirada en la primera versión. Un operador por fase, la misma opción:

   | Fase | Declaración | Operador |
   |---|---|---|
   | 3 | `pattern_paths` | unión de conjuntos |
   | 2 | `env_isolation.allowlist`, `env_isolation.root` | identidad; contenido del allowlist |
   | 0 | `host_labels` | identidad; contenido |
   | 1 | `evidence.path` | identidad |
   | 3 | `evidence.path` | identidad (la evidencia se declara por fase) |
   | 4 | rutas `@` de `gates[].args` | identidad (el conjunto de gates lo fija el canon) |
   | 5 | `operating_doc.doc`, `operating_doc.linked_from` | identidad |
   | 6 | `prs` (número, sha y veredicto de cada entrada) | contención (la cadena crece) |

   Con la MISMA ruta, el contenido tampoco puede medir menos que el de la base
   (#1341): la identidad de la ruta no dice nada de lo que hay dentro, y un PR
   que reescribe el allowlist en su sitio se llevaría la medición sin que nada
   lo detectara. El allowlist no puede **autorizar de más** —una variable de
   entorno, una extensión excluida o un fichero de datos nuevo es una
   autorización que la base no daba— y la instantánea de etiquetas no puede
   **dejar de cubrir** las que cubría: una etiqueta que desaparece apaga la
   obligación de inventariarla. Medir más sí se admite: quitar una
   autorización o añadir una etiqueta endurece la medición. La ampliación
   legítima se declara como exención con identidad en el propio allowlist
   (`exemptions`), que nombra la superficie, el elemento y quién la autoriza:
   sin `actor` o `issue` la exención no exime de nada.

   El contenido del árbol que escanea la fase 2 (`env_isolation.root`) NO se
   compara con el de la base: ese árbol es el código que el PR cambia, y
   compararlo pondría en rojo todo PR legítimo. Lo que sí vigila la identidad
   es que la ruta sea la misma y que un movimiento sea puro.

   En la fase 6 la base solo obliga a conservar sus entradas tal como están:
   la cabeza puede añadir las suyas. Cambiar el sha o el veredicto de una
   entrada de la base es apartar una aceptación, y el gate lo rechaza.

   La evidencia es una declaración POR FASE (#1348): la fase 1 mide la
   evidencia de la propagación y la fase 3 la del reemplazo de gobernanza, y
   son documentos distintos. Cada fase se compara con la de SU fase en la
   base. Con una sola ruta para las dos, la fase 3 se comparaba contra la
   declaración de la fase 1 y el gate daba un hallazgo falso: con base y
   cabeza idénticas salía en rojo, y el rojo no dependía del PR.

   La fase 4 mide el conjunto de rutas que los gates declaran medir (el
   consumer elige las entradas de cada gate, no el subconjunto), así que su
   operador es de conjunto y no de ruta. La excepción de retirada es la misma
   que en las demás fases de identidad (#1344): si una ruta que mide la base
   ya no está en el árbol medido, solo se admite un movimiento puro — la
   cabeza mide una ruta nueva con el mismo contenido, byte a byte. Un
   movimiento con edición en el mismo PR da hallazgo, y retirar la ruta sin
   medir otra con su contenido también.

   El sello entra del PR como **dato**: el job hace checkout de un solo fichero
   (`.team-skills.yaml`) y lo copia al workspace, porque la acción lo lee de ahí.
   Del PR no entra nada más, y nada de lo que entra se ejecuta. Es lo que
   mantiene verdes las PRs de propagación, que cambian el sello y re-sellan el
   manifiesto. El residuo de esta forma es que un PR puede apuntar el sello a un
   commit más antiguo del catálogo: degrada el gate, no roba el token.

   Cuando el sello no coincide, el gate nombra el hash esperado, el hash
   encontrado y el comando `ci-pattern manifest seal <repo>` (#1172). El job
   publica la copia re-sellada como artefacto `governance-manifest-resealed`:
   descargarla y commitearla es de quien abre el PR; el job no empuja nada.

   Tres cosas que el gate no perdona:

   - El job va en el `needs:` del agregador `required` **y** en `required_jobs`
     de `.github/required-jobs-policy.json`. En un repositorio privado con plan
     Free es la única vía de enforcement: `host_capabilities` deja ahí
     `branch_protection`, `rulesets` y `required_checks` en `unavailable` y el
     único check exigido es el agregador del propio CI
     (`required_checks: [{"name": "required", "class": "runner-enforced"}]`).
     Sin esa entrada el job corre y nadie lo exige.
   - La CLI se invoca como **comando de primer nivel** del paso, nunca a través
     de otro gate. `verify` ejecutaba `adoption check` en proceso y el ciclo
     superó 1.300 procesos de python y tumbó el VPS por OOM el 2026-10-05
     (HR-46). La marca `CI_PATTERN_NESTED=1` hace fallar el anidamiento con
     exit 2, pero es una red, no un permiso para encadenar.
   - Sin declaración `governance_seals` la fase 3 avisa a `stderr` y sale 0: la
     adopción previa no se rompe. En cuanto la declara, el rojo aparece en el PR
     que declara, que es justo lo que se busca.

**Dónde se mide cada fase.** La tabla nombra el destino de la 0 a la 6 y por
qué; las que dependen de credenciales del host o de un merge ya ocurrido no son
gates de PR por construcción, y no se fuerzan:

| Fase | Qué mide | Dónde se mide | Por qué |
| --- | --- | --- | --- |
| 0 | Inventario exhaustivo y etiquetas del host | Fuera del PR (`operating-doc-drift`, host-readback) | Lee el host, no el árbol: necesita credencial y, en plan Free, el 403 que #330 solo acepta si el contrato declara el área `unavailable` |
| 1 | Evidencia de propagación | Fuera del PR por ahora | Solo existencia de ficheros; medible en un PR, pero cablearla es trabajo de un PR propio, no de la adopción de la 3 |
| 2 | Aislamiento | Job propio `env-isolation` (HR-28) | Ya cableado en los consumers que lo declararon; el job invoca el asset, no la fase |
| 3 | Sellos del manifiesto y cableado de `P23_required_jobs` | **Job de PR `governance-seals` (HR-72)** | Lee el árbol y el contrato; es la única cuya rojo puede invertir un PR antes del merge |
| 4 | Los cinco gates del patrón en verde | Fuera del PR | Mide nombres de job y checks ya publicados: no hay PR que evaluar |
| 5 | Deriva del documento operativo | Job programado o `docs-gates` en el PR | Medirlo en cada PR «gastaría minutos sin valor de revisión» (#858) |
| 6 | Cadena de aceptación | Fuera del PR | Necesita el PR fusionado: número, SHA de 40 hex y `conclusion: success` |

Cierre: `ci-pattern adoption check --phase 3 <repo>` = 0.

## Fase 4 — Contratos cableados

- **G4.1 — Issue-spec equivalente** instalado y validando el formato canónico
  de issue en el destino.
- **G4.2 — Presupuesto de tamaño** sin exclusiones (adiciones + eliminaciones del
  PR, 400). La única excepción es la etiqueta `size:exception`, aplicada solo
  por el mantenedor cuando partir la PR rompería `main`; el campo
  `size-exception-reason:` del cuerpo está retirado. Encadenar PRs es el camino
  normal cuando el trabajo no cabe. Si el consumer mide A1 con el gate de
  contrato de PR (`pr-contract`), su workflow `MUST` generar el mapa de
  orígenes y pasarlo con `--origins-file` (plantilla:
  `assets/templates/.github/workflows/pr-contract.yml.tmpl`; generador:
  `assets/bin/origins_file.py`): sin el mapa el árbol vendorizado cuenta como
  líneas de autoría y la primera PR de propagación revienta el presupuesto
  (#1133, #1149). La fase 4 lo mide (#1159, HR-74). Y el CI del consumer `MUST`
  ejecutarla en cada PR (`adoption check --phase 4 pr-head`), como ya hace con la
  fase 5: si solo la lanza quien la corre a mano, ningún gate ve la deriva del
  cableado (#1359).
- **G4.3 — Preflight ejecutable** en local con un solo comando, paridad con el
  job de lint (HR-4). La política de paridad declara lo que cada comprobación
  necesita del entorno: `requires` en cada exclusión y `local.capabilities` con
  lo que tiene la máquina local. Vocabulario cerrado: `docker`, `runner`,
  `network` y `credentials`. El gate falla si el entrypoint local ejecuta un
  comando que exige una capacidad que el entorno local no declara. Declarar
  `requires` no saca la comprobación de la comparación: sigue siendo una
  diferencia declarada, con `side` y `reason`. Limitación conocida: la capacidad
  se declara sobre las exclusiones; un comando comparado que necesita Docker
  todavía no tiene declaración propia.
  Para correr una sola comprobación sin Docker, ejecute el comando que figura
  para ella en el entrypoint local, desde la raíz del repo. Los gates de
  documentación de la fase 5 no necesitan Docker: `ci-pattern adoption check
  --phase 5 <repo>`. El contrato de PR en local (#912) tampoco lo necesita; esta
  guía no lo documenta hasta que esa implementación esté en main.
- **G4.4 — CI en la rama base**, net-new si el destino no tenía.
- **G4.5 — Binding evidencia↔deploy solo si existe deploy.** El estado de
  commit por SHA (HR-10) se instala únicamente cuando la fase 0 inventarió un
  mecanismo de deploy; sin deploy no hay evidencia que registrar y el gate solo
  añade ruido. Verificación de salida de fase: el evaluador resuelve la
  revisión previa desde la fuente desplegada declarada (HR-40) y su política
  de exenciones bootstrap es simétrica entre los contextos requeridos
  (HR-41) — nunca por heurísticas de listas de runs ni por exención ad-hoc
  de un solo contexto.

- **G4.6 — Imágenes de CI fijadas por digest (opt-in).** Declare
  `P56_ci_images_gate: on` en `ci-pattern.yaml` y ejecute el gate
  `assets/ci-images/check_ci_images_pinned.py` en un job de CI; las excepciones
  van en `P57_ci_image_exceptions` (referencia → motivo). Sin declaración, inerte (HR-78).

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
   Si su cadena de PRs usa el modelo **rama de feature con PR tracker en
   borrador** (HR-54, regla 7), declare además `tracker_branch` con esa rama:
   el detective post-push pasará a exigir que el PR mergeado en cada rama
   protegida tenga esa rama como cabeza. Sin la declaración el control no
   cambia, que es lo correcto para una cadena apilada hacia `main`.
2. Capture las instantáneas con GETs de solo lectura: para
   `branch-protection` y `rulesets` guarde el cuerpo real del `403`, y para
   `actions-runners` la respuesta de `repos/<owner>/<repo>/actions/runners`.
3. Instale los tres controles (`assets/runner-controls/`) y cablee un
   workflow por control en el runner declarado: detective post-push (HR-49),
   merge gobernado (HR-50) y guard de force-push (HR-51), cada uno en un job
   cuyo `runs-on` incluya `runner.label`. El job de merge gobernado declara
   `issues: write`: el merge con `GITHUB_TOKEN` no dispara el cierre
   implícito de las palabras de cierre del PR y el asset cierra esas issues
   él mismo (#800); sin el permiso termina en exit 3 («fusionado, issue sin
   cerrar»).
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

- **G5.6 — `CONTRIBUTING.md` del consumer (HR-71).** Para consumers
  `primary_type: web` que activen `docs_manifest`, la fase 5 exige
  `CONTRIBUTING.md` con las secciones mínimas del manifiesto. La activación
  se resuelve leyendo `project_type` del `.team-skills.yaml` del repo
  (`_active`, `_resolve_project_type` en
  `assets/docs-manifest/check_docs_manifest.py`); si el YAML falta, no se
  puede leer o no declara `project_type`, HR-71 falla cerrado (HR-3): emite
  el hallazgo y activa el doc para que la ausencia del fichero también se
  mida. El catálogo **nunca** sobrescribe ni sella `CONTRIBUTING.md` en
  `.governance-manifest.json`; el contenido es del consumer y el gate solo
  mide existencia y secciones. La plantilla canónica está en
  `assets/templates/docs/CONTRIBUTING.md.tmpl`; copiada byte a byte en
  títulos, el repo la reescribe a su gusto en contenido.

- **G5.7 — Las tablas de checks escritas a mano, declaradas (HR-82, #1382).**
  El consumer que tenga una tabla de jobs escrita a mano (la que explica qué
  hace cada check del CI) la declara en su contrato de adopción:

  ```json
  "checks_tables": {
    "tables": [
      {"path": "docs/codebase/testing-and-ci.md", "workflow": "ci.yml",
       "section": "Los jobs de CI"}
    ]
  }
  ```

  La fase 5 mide cada fila contra los workflows reales: un job que ningún
  workflow define (fabricado), un job que vive en otro workflow (la fila lo
  atribuye a otro) y un job requerido que no aparece en ninguna tabla
  declarada son rojos. La clave es **opcional**: sin ella el gate no se mide y
  el consumer no cambia de veredicto hasta que la declara. La página del gate
  es `assets/checks-table/README.md`.

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
