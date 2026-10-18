---
name: ci-pattern
description: "Trigger: CI pattern, deterministic CI gates, review budget, SHA evidence, chained PRs, CI adoption gate, porting guide, issue-spec. Distills the repository CI pattern (deterministic gates, review budget, SHA evidence, chained PRs, continuous-improvement protocol), governs its adoption in another repository (ADOPTION GATE STOP + references/porting-guide.md) and teaches how to audit it."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.20.0"
  last_verified: 2026-10-17
  based_on: "auditoría issue→merge de ardelperal/APAP_WEB (épica ardelperal/APAP_WEB#935, 2026-09-29/30)"
  scope: ['universal', 'ops']
  auto_invoke: ['adopt the CI pattern in another repo', 'apply the CI pattern in a new repo', 'audit this CI pattern', 'review budget and chained PRs', 'SHA deploy evidence', 'CI adoption gate', 'porting guide']
  tiers: ['universal', 'ops']
---

# Patrón de CI — gates deterministas y evidencia por SHA

El patrón combina siete piezas sobre un mismo repositorio:

1. **Gates deterministas** que leen datos estructurados (rama, etiquetas, campos de
   la API, campos del cuerpo del PR), nunca prosa ni gestos.
2. **Presupuesto de revisión** por PR (adiciones + eliminaciones), con la sola
   excepción de la etiqueta `size:exception` —puesta exclusivamente por el
   mantenedor— y PRs encadenados cuando el trabajo no cabe.
3. **Evidencia ligada a una revisión**: cada deploy y cada batería de producción
   registran su veredicto como estado de commit sobre el SHA desplegado.
4. **Protocolo de mejora continua**: toda fricción se registra con evidencia, se
   arregla por el pipeline y se destila en regla; la recurrencia dispara
   automatización.
5. **Release con evidencia**: releases como tags semver anotados, hotfix con
   bump de PATCH y deploy inmediato, y post-mortem blameless tras todo
   incidente de producción (HR-19 a HR-21).
6. **Gobernanza de orquestación**: toolchains pineadas para medir y suite
   aislada del entorno local (HR-26, HR-28).

El catálogo destilado y el veredicto de gates viven en `references/`; los
parámetros portables, en `assets/parameters.md`. La ubicación actual de los
scripts de implementación de referencia se declara en `assets/parameters.md`.

**Frontera de alcance.** Este patrón cubre el gobierno de CI del repositorio:
gates deterministas, presupuesto de revisión, evidencia por SHA, cadena de
PRs, release y post-mortem. La autoridad de review —desarrollo dirigido por
recibos (RDD): linajes, recibos, rondas de jueces, sobres de consentimiento,
presupuestos de corrección y acknowledgement— es capacidad del harness
`gentle-ai`; este patrón la asume instalada y no la re-implementa ni la
documenta como propia. La memoria persistente es capacidad de `engram`. Son
los únicos dos harnesses externos que el patrón presupone.

## §1 Activation

### GATE DE ADOPCIÓN (STOP)

Antes de aplicar este patrón en un repositorio NUEVO: completar
`references/porting-guide.md` fase por fase (inventario read-only → auditoría
del mecanismo de propagación REAL → aislamiento de entorno → gobierno viejo y
nuevo en el MISMO PR → contratos cableados → primer PR real como acceptance
test). El checklist es determinista: cada fase cierra con el gate
`ci-pattern adoption check --phase N` = 0 (HR-45, HR-46) y la IA no decide el
orden ni cuándo ha terminado. Sin ese checklist completado y verificado: NO se
lanza ningún worker, NO se toca el repo destino, NO se abre PR. Excepción:
ninguna.

Anclaje: la adopción en Cadete (2026-09-30/10-01) aplicó el patrón sin el
checklist y produjo 18 fallos falsos de batería, 3 diagnósticos erróneos
(issue ardelperal/APAP_WEB#1130 cerrada con corrección), 5 prescripciones
rotas, una premisa de propagación falsa y una colisión de shared-checkout
(post-mortem: `docs/postmortems/2026-09-30-ci-pattern-adoption-cadete.md` en
`ardelperal/APAP_WEB`). Este gate existe para que ese modo de fallo sea
estructuralmente imposible.

Cargue esta skill cuando:

- Vaya a **adoptar o auditar este patrón de CI en otro repositorio** (inventario
  de gates, medición, instalación, validación real).
- Abra, espere o revise **PRs bajo este patrón**: etiquetas al crear, presupuesto,
  cadena de PRs, cierre de issue, merge con evidencia.
- Deba **clasificar un rojo de CI** o decidir entre push, rerun o dispatch.
- Toque **evidencia de deploy**: estados por SHA, smoke de producción, baterías.
- Vaya a **destilar una fricción nueva** en regla del playbook.
- Deba **delegar u orquestar trabajo concurrente** entre actores (worktrees por
  actor, esperas de CI, verificación en vivo de prescripciones).
- Deba **cortar un release o un hotfix** (tag semver, notas de release,
  playbook de deploy por release).
- Toque un **incidente de producción** (issue, fix, deploy inmediato,
  post-mortem blameless, action items como issues).

No la cargue cuando:

- Escriba tests de aplicación o decida su capa (eso es la skill local de testing
  del consumer, parámetro 12).
- Toque auth, CSRF o secretos de la aplicación (eso es la skill local de
  seguridad del consumer, parámetro 12).
- Ejecute la batería e2e de producción (eso es el runbook e2e del consumer,
  parámetro 11; esta skill solo gobierna cómo se registra el veredicto).

Fuentes normativas:

- `references/fricciones.md` — catálogo destilado con antídoto.
- `references/porting-guide.md` — checklist de pre-vuelo de la adopción, fase
  por fase; su cumplimiento es lo que el GATE DE ADOPCIÓN exige.
- `references/gate-verdicts.md` — veredicto por gate con evidencia.
- `references/benchmark-gentle-ai.md` — ideas contrastadas de otro CI.
- `assets/parameters.md` — los parámetros que se extraen por repo.
- `references/incidents.md` — ejemplo destilado de hotfix y
  post-mortem blameless sobre un incidente real.

## Uso desde una IA — CLI `assets/bin/ci-pattern`

El adoptador/verificador determinista del patrón vive en `assets/bin/ci-pattern`
(relativo a la raíz de la skill; se resuelve igual en el catálogo canónico
`DysTelefonica/team-skills` y en cualquier mirror). Su contrato normativo es
`references/cli-spec.md`; su salida está diseñada para ser leída por un agente
(terse, estructurada, accionable). Frases de disparo del usuario: «actualizá el
sistema de gobernanza», «adoptá el patrón en este repo», «verificá el
cumplimiento».

Flujo de comandos (la adopción la gobierna `adoption check` sobre el contrato
del consumer; `adopt` y `update` están retirados):

```bash
# Forma portable (el binario lleva shebang y modo 100755):
python3 <skill>/assets/bin/ci-pattern status <repo>

# 1. Estado del repo destino: ¿hay manifiesto? ¿está limpio?
ci-pattern status <repo>

# 2. Validar los parámetros ANTES de escribir nada (cierra el hueco G3)
ci-pattern params validate ci-pattern.yaml

# 3. Adopción: documento operativo generado y gate por fases del contrato
ci-pattern adoption generate-doc <repo> --out docs/ci-pattern-flow.md
ci-pattern adoption check --phase 5 <repo>

# 4. Verificación de cumplimiento tras cualquier cambio
ci-pattern verify <repo> --json
```

Códigos de salida: 0 limpio · 1 hallazgos · 2 uso · 3 recurso ausente. La CLI nunca hace merge, push, ni toca producción, ni borra ficheros,
ni sobrescribe un fichero modificado localmente sin reportarlo como hallazgo;
los valores de los parámetros, los conflictos de contenido y la disposición de
los veredictos siguen siendo juicio humano/IA (§8 de `references/cli-spec.md`).

## §2 Hard Rules

- **HR-1 — Los gates `MUST` leer datos estructurados** (nombre de rama, etiquetas,
  campos de la API, campos del cuerpo del PR) y `MUST NOT` deducir nada de prosa,
  gestos ni etiquetas aplicadas tarde. (evidencia: R2; fricciones F-004, B4)
- **HR-2 retirada (2026-10-15)**
- **HR-3 — Un gate `MUST` fallar en voz alta cuando no ha medido** y `MUST NOT`
  imprimir «OK» sin ejecución real detrás. (evidencia: R4; fricción A9; gate:
  `assets/workflow-policy/check_workflow_policy.py`)
- **HR-4 — Todo job que gatea el merge `MUST` tener paridad local: cada comando
  que el CI ejecuta en ese job, con las mismas opciones/flags que el CI le
  pasa, se reproduce desde el punto de entrada local declarado en la política,
  con los comandos CI-only declarados como datos de excepción con `reason`
  (fail-closed: un comando CI-only sin excepción declarada, o ejecutado con
  opciones/flags distintas a las que el local declara, es violación)**; el
  autor `MUST` ejecutar el punto de entrada completo antes de cada push.
  (evidencia: R5; F-001, B3; friccion F1 #396; enmienda
  DysTelefonica/team-skills#156, #396; gate:
  `assets/local-ci-parity/check_local_ci_parity.py`)
- **HR-5 retirada (2026-10-15)**
- **HR-6 — Las etiquetas del PR `MUST` aplicarse en el comando de creación**
  (higiene del autor): aplicadas tarde, alguien tiene que reejecutar o empujar
  para que el gate lea el dato nuevo, y esa higiene `MUST NOT` tratarse como
  evaluación del gate — lo que se evaluó fue el PR sin las etiquetas (HR-31).
  (playbook regla 2)
- **HR-7 — Solo el PR punta de una cadena `MUST` llevar `Closes #<issue>`**; los
  intermedios llevan `Refs` más la etiqueta de cadena, la rama se nombra
  `<tipo>/<N>-<slug>`, el autor `MUST` verificar `closingIssuesReferences` tras
  crear el PR y `MUST NOT` escribir palabras de cierre en el título. (R2; B11; playbook reglas 3-4)
- **HR-8 retirada (2026-10-15)**
- **HR-9 retirada (2026-10-15)**
  (R14; vigías zombis del 2026-09-29)
- **HR-10 — La evidencia de deploy o batería `MUST` registrarse sobre el SHA de la
  revisión desplegada** (estado de commit por SHA) y `MUST NOT` colgar de una
  variable global ni de la rama por defecto. (R6; ardelperal/APAP_WEB#1082)
- **HR-11 — Toda fricción `MUST` registrarse con evidencia (PR, issue, run o
  `fichero:línea`)**, arreglarse por el pipeline normal y destilarse en regla con
  su ternario completo; la segunda ocurrencia `MUST` disparar automatización, no
  otro arreglo manual. Cuando lo que falla es la propia skill en un consumer,
  el registro viaja a `DysTelefonica/team-skills` con `report-friction`
  (clasificación explícita, saneado y deduplicación por huella; el comentario de
  ocurrencia lleva el recuento y la segunda ocurrencia marca la etiqueta no
  protegida `friction:recurrent`). (R14; protocolo de mejora continua)
- **HR-12 retirada (2026-10-15)**
- **HR-13 — Los mensajes de un gate `MUST` describir la causa real** y, cuando el
  fallo depende de un dato que el autor no puede corregir, `MUST` decirlo y ofrecer
  una vía manual auditable. (R10; B4, B11)
- **HR-14 — Los commits `MUST` ser conventional y sin atribución de IA**:
  regla del contrato raíz del repo destino (`AGENTS.md`); este patrón no la
  duplica — el gate `assets/pr-contract/check_pr_contract.py` la aplica. Un
  commit de merge no se juzga por el asunto conventional —la regla mide los
  commits de trabajo, no la sincronía—, porque la disciplina de cadenas manda
  sincronizar con `git merge origin/main` y nunca con rebase (#333).
  (COL3; regla del repo)
- **HR-15 — Los ratchets `MUST` ser shrink-only** (rechazan lo nuevo, aceptan el
  inventario actual, regeneran baseline con un comando). Toda entrada de gate
  que un PR puede editar (baseline, policy file, allowlist) `MUST` compararse
  contra la copia de la rama base, y toda relajación — una cifra que sube o
  una entrada añadida a un manifest declarado — `MUST` llevar un campo de
  datos explícito con motivo y referencia, validado por el propio gate, y
  `MUST NOT` aceptarse sin él; en cadenas, la baseline sube en el slice
  aditivo y baja en el que consume, el slice aditivo lleva el mismo campo de
  relajación, y el fixture que un slice necesita `MUST` viajar en ese slice.
  (T2; playbook regla 12; team-skills#135; fricciones D3; gate:
  `assets/ratchet/check_ratchet.py`)
- **HR-16 — Las huellas de secretos y las baselines de gates `MUST` anclarse a
  identificadores estables** (SHA de contenido o ruta), nunca a números de línea
  mutables que rompen el gate en el primer reordenado. (seguimiento
  ardelperal/APAP_WEB#1112; gate: `assets/ratchet/check_ratchet.py`)
- **HR-17 retirada (2026-10-15)**
- **HR-18 retirada (2026-10-15)**
- **HR-19 — Todo release `MUST` ser un tag semver anotado** (`vMAJOR.MINOR.PATCH`)
  con GitHub Release asociada y marcada `Latest`; los prereleases
  (`vX.Y.Z-rc.N`) `MUST` quedar excluidos de la estable (`v*-*`), y las notas de
  release `MUST` ser concisas (qué cambió + enlace al issue o post-mortem), sin
  enterrar nunca la causa raíz del incidente. (incidente de pérdida de datos;
  `references/incidents.md`; gate del tag:
  `assets/release-gate/check_release_tag.py`)
- **HR-20 — Todo hotfix `MUST` aterrizar primero en main** y salir como bump de
  PATCH con deploy inmediato tras el merge; el fix `MUST NOT` vivir solo en una
  rama paralela ni esperar al próximo release regular. (incidente de pérdida de
  datos; `references/incidents.md`; gate del ancestro:
  `assets/release-gate/check_release_tag.py`)
- **HR-21 — Todo incidente de producción `MUST` cerrar con post-mortem
  blameless** en la ruta de post-mortems del consumer (parámetro 9;
  secciones: Timeline UTC / Impact / Root cause / What worked / What failed),
  con causas de sistema y nunca de personas, y cada action item `MUST`
  abrirse como issue de GitHub con owner. La trazabilidad del incidente `MUST`
  repartirse en **cuatro niveles enlazados por identificadores** —ID de
  incidente, número de ticket, SHA, ejecución de CI, issues, PRs y nombres de
  sistemas; identificadores sí, contenido de datos personales no—: **(1) el
  post-mortem**, con el qué pasó y sus identificadores; **(2) el registro
  operativo**, con el detalle técnico que hace falta para seguir atendiendo
  los tickets abiertos (IPs, hostnames, comandos, logs, cronología fina y el
  siguiente paso de cada ticket) en la ruta que declara el consumer (parámetro
  24; por defecto, el mismo repositorio en
  `docs/incidents/<AAAA-MM-DD>-<slug>.md`, y un repositorio de operaciones
  aparte es opcional); **(3) los secretos** —credenciales y tokens— en el
  gestor de secretos declarado (parámetro 25), nunca en git, ni siquiera en un
  repositorio privado; y **(4) los datos personales y las conversaciones con
  terceros**, que se quedan en el sistema de tickets y se referencian por su
  número. La evidencia `MUST` citarse con su `sha256` y su ubicación de acceso
  controlado, y lo redactado `MUST` marcarse de forma explícita (`[IP interna
  redactada]`), nunca borrarse en silencio. El gate de la fase 5 `MUST`
  bloquear secretos y datos personales en el post-mortem y en el registro
  operativo, y la evidencia declarada sin hash ni ubicación; las IPs, los
  hostnames y los comandos internos están permitidos (el repositorio es
  privado y lo ven las mismas personas). (canon del sector, precedente GitLab
  2017; `references/incidents.md`; decisión del operador, 2026-10-06;
  DysTelefonica/team-skills#322; gate: `assets/bin/ci-pattern adoption check
  --phase 5`)
- **HR-22 retirada (2026-10-15)**
- **HR-23 retirada (2026-10-15)**
- **HR-24 retirada (2026-10-15)**
- **HR-25 retirada (2026-10-15)**
- **HR-26 — Todo gate que compara números `MUST` medir con la toolchain
  versionada que lo juzga en CI** (pin de versiones en la stack de medición o
  re-medición por la toolchain del juez) y `MUST NOT` comparar una medición
  local con un baseline producido por otra versión: cobertura y baselines no
  son comparables entre toolchains. (baseline local 11503 vs runner 11559 por
  xdebug sin pinear; gate: `assets/workflow-policy/check_workflow_policy.py`)
- **HR-27 retirada (2026-10-15)**
- **HR-28 — El entorno de la suite `MUST` aislar del `.env` local del
  desarrollador, y el gate de aislamiento `MUST` declarar su dominio y fallar
  fuera de él**: un `.env` gitignored en la raíz filtra sus variables
  (p. ej. `APAP_*`) a `Settings` y rompe tests solo en local; si el
  aislamiento no existe todavía, los rojos ambientales `MUST` documentarse
  como fallo de entorno, nunca como defecto de código. El gate declara
  además las extensiones que escanea (`DOMAIN_EXTENSIONS`: `.sh`, `.py`) y
  `MUST` fallar en voz alta, nombrándolos, si el árbol de suites contiene
  ficheros de CÓDIGO de otro lenguaje que el contrato no excluya
  explícitamente (`excluded_extensions` del allowlist versionado): un sujeto
  incompleto no es un pase, y ampliar el dominio con un escáner nuevo `MUST`
  hacerse ADEMÁS de ese fallo, nunca en su lugar. Las menciones literales de `.env` que son DATO (fixtures que escriben un
  `.env` en tmp_path, aserciones sobre `.gitignore`) se declaran por fichero
  en `dotenv_data_files` con motivo, y esa exención solo cubre la mención:
  un cargador o una lectura de entorno sigue siendo hallazgo. Una exención
  que ya no hace falta falla como `HR-28 STALE`. `templates.env` (acceso a
  atributo) y `postgresql://dotenv` no son lectura (DysTelefonica/team-skills#635).
  (5 rojos ambientales locales por el `.env` de la raíz, 2026-10-01;
  DysTelefonica/team-skills#315)
- **HR-29 — El agregador de jobs requeridos `MUST` verificar la paridad de los
  tres conjuntos que lo sostienen**: `jobs(workflow) − {agregador}`, el `needs`
  del agregador y el conjunto conocido por el evaluador `MUST` ser iguales; una
  clave de `needs` que el evaluador no conozca `MUST` contarse como violación y
  `MUST NOT` ignorarse, y un job excluido a propósito del agregado se declara
  como dato con su motivo, nunca por omisión. Al añadir, renombrar o eliminar un
  job, los tres conjuntos se actualizan en el mismo PR: un job cableado en
  `needs` y ausente del conjunto conocido, o añadido al workflow sin cablear,
  queda fuera de todo veredicto y el agregador sale limpio con un fallo real
  dentro. (team-skills#133; derivada de la lectura del código del agregador de
  origen, 2026-10-01; fricciones D1)
- **HR-30 — Un nombre de check requerido `MUST` publicarse solo desde un evento
  que evalúa el PR**; en cualquier otro evento (`push`, `workflow_dispatch`,
  `schedule`) el job `MUST` publicar bajo otro nombre de contexto o fallar, y
  `MUST NOT` emitir éxito sin haber evaluado el PR: un verde requerido que no
  evaluó es un falso verde indistinguible del bueno y anula la protección de
  rama. (team-skills#134; derivada de la lectura del código de origen,
  2026-10-01; fricciones D2; gate: `assets/workflow-policy/check_workflow_policy.py`)
- **HR-31 — Un gate que lee datos mutables del PR o de la issue enlazada
  (cuerpo del PR, etiquetas, estado y etiquetas de la issue) `MUST` reejecutarse
  cuando esos datos cambian**, o su verde `MUST` invalidarse reevaluando el gate
  en el momento del merge. Límite honesto: un cambio en la issue enlazada no
  dispara eventos del PR en el host, así que para ese dato la reevaluación en el
  merge no es opcional — la regla `MUST NOT` prometer un disparador que no
  existe. Un verde que sobrevive a la edición del cuerpo o de una etiqueta no
  certifica nada. (team-skills#134; derivada de la lectura del código de origen,
  2026-10-01)
- **HR-32 — Todo control —gate, auditoría, workflow programado o control
  compensatorio— `MUST` tener un test de comportamiento que ejecute su lógica
  contra un fixture que viola la regla y observe el veredicto de fallo**; un
  test que solo comprueba subcadenas del fuente de un workflow o de un script
  `MUST NOT` contarse como evidencia del control, y un control compensatorio
  de una carencia del host `MUST` cumplir el mismo baremo que un gate. Los
  tests de cableado de workflow no se prohíben: dejan de contar como prueba
  del control. Extiende a los controles fuera del harness la exigencia de
  `deterministic-quality-harness` de observar cada gate saliendo con `1` ante
  una violación real; no la sustituye ni la copia. (team-skills#136; derivada
  de la lectura del código de origen, 2026-10-01; fricciones D7; gate: el
  test de comportamiento del propio control — ver
  `assets/audit-issue-to-merge.md` bloque 2, fila HR-32)
- **HR-33 — Toda exención de gate `MUST` concederse por identidad verificable**
  — actor verificado por el host **y** origen de la rama en el mismo
  repositorio que el gate juzga — y `MUST NOT` concederse por un prefijo de
  texto del nombre de rama ni por cualquier otro dato que el autor elige:
  nombrar la rama con el prefijo exento no puede bastar para saltarse el gate.
  Las ramas generadas por la plataforma (reversión, bots de actualización de
  dependencias, ramas de propagación) son un caso de la regla, no una excepción:
  se admiten porque actor y origen constan en la lista de datos del parámetro
  16, y la exención que un prefijo concede sin mirar el actor es una brecha, no
  una exención. La regla no nombra prefijos, bots ni ecosistemas concretos: son
  datos del parámetro 16. (team-skills#137; derivada de la lectura del código
  de origen, 2026-10-01; fricciones D8; gate:
  `assets/exemptions/check_exemptions.py`)
- **HR-34 — Toda regla de gobierno declarada en la documentación —etiqueta,
  nombre de check requerido, método de merge, ruleset y flag de protección—
  `MUST` contrastarse contra la API del host mediante un drift check de solo
  lectura y `MUST` clasificarse como `host-enforced` o `documented-only`**;
  la capacidad del host por área —`host_capabilities` (`available` o
  `unavailable` para `branch_protection`, `rulesets` y `required_checks`) en
  el contrato del host— es dato verificado por la relectura en ambas
  direcciones: el `403` de plan de un repo privado sin Pro se registra
  `unavailable` (evidencia, no error) y una contradicción es drift. La
  tercera clase `runner-enforced` designa un control que aplica un workflow
  que corre en un runner propio declarado —`runner.label` en el contrato del
  host— y que, ante una violación, actúa y deja evidencia; solo se acepta en
  un área cuya capacidad sea `unavailable`, exige fixture violador y receta
  HR-32, y una regla `runner-enforced` en un área que el host sí aplica —
  etiquetas y ajustes de merge— es drift;
  la regla sin enforcement en el host es `documented-only` y `MUST NOT`
  describirse como gate. El check compara el contrato declarado con
  instantáneas JSON de las respuestas de la API (la red queda fuera del
  ejecutable), nunca muta el host ni prueba una escritura para descubrir
  permisos, y ante instantánea ausente o ilegible falla en voz alta (HR-3).
  Los principios de clasificación y verificación en vivo son de
  `repository-delivery-governance`; esta skill aporta el paso de adopción y
  el ejecutable (`assets/host-readback/`). (team-skills#138; API del host
  de origen, 2026-10-01; fricciones D9)
- **HR-35 retirada (2026-10-15)**
- **HR-36 retirada (2026-10-15)**
  el desajuste y se detiene, nunca se adapta en silencio. (repo
  equivocado en un encargo y superficies prescritas de memoria, 2026-10-02;
  hueco G10 del inventario de activos portable)
- **HR-37 retirada (2026-10-15)**
  presupuesto y su excepción es de `repository-delivery-governance` (HR-10)
  y se referencia, no se copia. (auditoría issue→merge del consumer de
  origen, solo lectura, 2026-10-01: de los últimos 60 PRs fusionados, 60 de
  60 con autor y quien fusiona en la misma cuenta, 3 de 60 con alguna
  revisión registrada y 0 aprobaciones requeridas;
  DysTelefonica/team-skills#139; fricciones D10)
Cada HR de esta lista figura en este SKILL.md y su clase de enforcement y la
evidencia actual están registradas en el bloque 2 de
`assets/audit-issue-to-merge.md` (matriz de enforcement con la clase
`documented-only` / `locally-validated` / `CI-enforced` / `host-enforced` /
`manual` y la ruta del activo que la sostiene; la antigua matriz
`references/hr-gate-matrix.json` y su meta-gate
`assets/hr-gate-matrix/check_hr_matrix.py` se retiraron por #513/#540).

Las HR-6 y HR-7 y HR-14 y HR-31 están gobernadas por el gate
`assets/pr-contract/check_pr_contract.py` (etiquetas obligatorias, Closes
solo en la punta de cadena, excepciones como dato, commits convencionales
sin atribución de IA, re-ejecución en edited/labeled/unlabeled). La clave opcional `checks` de la política (`{comprobación: true|false}` sobre `labels`, `closes`, `chain_context`, `issue_links`, `closure_tests`, `docs_evidence`, `budget`, `commits`) apaga comprobaciones sueltas para adoptar el contrato de forma incremental (DysTelefonica/team-skills#735): sin `checks` están todas activas; un nombre desconocido, un valor no booleano o las ocho apagadas salen 2 (HR-3), y las apagadas se listan por `stderr`. El gate no ejecuta `git`: el workflow genera los datos del rango y se los pasa por fichero (`--changed-files`, `--numstat-file`, `--commits-file`, `--tracked-files`); cada comprobación activa exige el suyo y su ausencia sale 2 (DysTelefonica/team-skills#753).

- **HR-38 retirada (2026-10-15)**
- **HR-39 — Un gate requerido por PR `MUST` ser función del diff y de entradas
  versionadas, y `MUST NOT` depender de un estado externo que cambia solo**
  (bases de avisos, registros, servicios): el gate cuyo veredicto depende de
  ese estado `MUST` ejecutarse de forma programada sobre la rama por defecto
  y, además, en los PRs que tocan los manifiestos relevantes —lista de rutas
  en datos, parámetro 22—; su rojo programado `MUST` abrir una issue con la
  evidencia y `MUST NOT` bloquear PRs que no tocan esos manifiestos.
  (gate: `assets/workflow-policy/check_workflow_policy.py` — la clave
  `state_dependent_gates` del policy file declara cada gate dependiente
  de estado con su `pr_workflow`, `pr_job`, `manifests`, `scheduled_workflow`,
  `command` y `alert_command`; gate: `assets/workflow-policy/tests/test_workflow_policy.py`;
  #396; cite DysTelefonica/team-skills#141; siete fallos del gate de avisos en
  cuatro ramas sin relación dentro de 80 minutos, 2026-09-29/10-01)
- **HR-40 — La entrada de revisión previa de todo gate de deploy o batería
  `MUST` resolverse desde lo realmente desplegado** (endpoint de salud con
  revisión, artefacto de release declarado o registro de deployment) y
  `MUST NOT` deducirse de heurísticas de listas de runs, índices de array o
  HEAD de la rama por defecto; si la fuente de resolución no está
  disponible, el gate `MUST` fallar en voz alta nombrándolo — nunca caer a
  un sustituto aproximado. (ardelperal / APAP_WEB#1221: el gate resolvió la
  revisión previa por head_sha `[0]` de la lista de runs / bloqueó todos
  los deploys; complementa HR-10 —dónde se registra el veredicto— sin
  duplicarla: esta fija cómo se resuelve la entrada)
- **HR-41 — Toda exención de bootstrap (revisión desplegada antes de que
  existiera el contexto) `MUST` declararse como dato de política simétrica
  para TODOS los contextos requeridos del mismo gate**, y la admisión
  ad-hoc de la exención en un contexto y no en otro `MUST` tratarse como
  violación de la política, no como configuración. (asimetría observada
  entre `release/e2e-production` y `release/smoke-production`,
  ardelperal/APAP_WEB#1221; encaja en la política de exemption declarada en `assets/parameters.md`)
- **HR-42 — Todo paso de runbook que dependa de estado externo (filas de
  BD, usuarios sembrados, flags, secretos, DNS) `MUST` crearlo de forma
  idempotente o verificarlo fail-loud antes de usarlo, con el comando de
  verificación escrito en el propio runbook**, y una prosa del tipo «is
  confirmed seeded» sin comando detrás es un defecto del runbook, no un
  estado de producción. (runbook del gate de release asumía la fila
  `e2e@apap.local` que producción había perdido —
  ardelperal / APAP_WEB#1223; gobierna el runbook / complementa HR-3 / HR-13
  sin duplicarlas)
- **HR-43 — La salida de un agregador verde `MUST` listar los jobs saltados
  con el marcador que los autoriza**, y un agregador que no distinga medido
  de saltado `MUST` fallar: un verde agregado sobre un árbol donde jobs
  requeridos no midieron (carril docs-only, marcador fail-closed) sin
  publicar qué miembros midieron y cuáles saltaron es un falso verde
  indistinguible del real. (auditoría del CI de
  `ardelperal / APAP_WEB#1219` / clase F1; complementa HR-3 sin duplicarla)
- **HR-44 — Toda edición material de una skill `MUST` añadir un fragmento
  de changelog válido en `changelog.d/` del propio skill** (`type: patch`
  para reglas y evidencia, `type: minor` para piezas de ola, más una
  entrada). El PR no fija `metadata.version` ni `CHANGELOG.md`: el PR de
  release consolida los fragmentos pendientes y sube la versión (#665). Un
  cambio sin fragmento válido falla cerrado en el gate
  `changelog-fragment`; solo el PR de release lo esquiva, y lo reconoce por su
  forma exacta: borra fragmentos, no añade ninguno, cambia en `SKILL.md` solo
  la línea `metadata.version` y en la skill nada más que un `CHANGELOG.md`. Y una sync de espejo `MUST` partir de un SHA resuelto de la
  fuente (`fetch` + `rev-parse`) verificando el readback contra ese SHA —
  nunca contra un working tree. (la v0.3 ganó HR-29 / 30 / 31 sin bump / el
  espejo del consumer quedó desincronizado; segunda deriva por sync desde
  checkout atrasado: `ardelperal / APAP_WEB#1211`)
- **HR-45 — La adopción del patrón se gobierna por contrato de datos y gate
  por fase**: el consumer declara su adopción en
  `.github/ci-pattern-adoption.json` (esquema versionado en `assets/adoption/`)
  y `ci-pattern adoption check (--phase N | --all) <repo>` verifica cada fase
  del porting-guide contra el árbol real; una adopción solo está completa
  cuando `--all` sale 0, ninguna fase pasa si la anterior no pasa, y `verify`
  la exige en verde cuando el contrato existe. `adopt` y `update` están
  retirados: la adopción la gobierna una IA con este contrato, no un
  scaffolder. (#276; gate: `assets/bin/ci-pattern adoption check`)
- **HR-46 — Todo artefacto de gobierno del destino que no venga del patrón
  `MUST` eliminarse en el PR de sustitución**: la adopción no conserva el
  gobierno original, ni en parte ni «en paralelo»; tras la fase 3 solo queda
  el gobierno del patrón. El gate de fase 3 ejecuta la comprobación —escanea
  el árbol con los globs de gobierno y las instantáneas del host
  (`branch-protection`, `rulesets`)— y falla nombrando cada artefacto
  residual o inventariado sin destino declarado. La identidad del patrón
  se verifica EN PROCESO contra el manifiesto de gobierno y sus rulesets
  salen de `.github/host-contract.json`: un gate de adopción NUNCA invoca
  `ci-pattern verify` ni `adoption check` como subproceso (verify ejecuta
  `adoption check` en proceso: el ciclo recursivo superó 1.300 procesos
  python y tumbó el VPS por OOM el 2026-10-05). Cualquier subproceso que
  esta CLI lance nace con `CI_PATTERN_NESTED=1`; si la marca ya está
  presente, `verify` y `adoption` fallan cerrados con exit 2.
  El smoke-run del meta-gate HR no es un gate anidado: el test declarado
  corre con la marca RETIRADA (puede llamar a `adoption` en proceso) y el
  encadenamiento se cierra por identidad —el test que ya se está ejecutando
  no se vuelve a smoke-ejecutar, `CI_PATTERN_SMOKE_TESTS` lleva sus rutas—,
  con `CI_PATTERN_SMOKE` como profundidad y tope trampa. Así ningún camino
  recursa y los tests declarados por los fixtures siguen juzgándose
  (DysTelefonica/team-skills#342).
  (decisión del operador,
  2026-10-05; DysTelefonica/team-skills#277; gate:
  `assets/bin/ci-pattern adoption check --phase 3`)
- **HR-47 — El ciclo de vida documentado se comprueba con una prueba de recorrido**: una IA que solo dispone del documento operativo generado completa el ciclo issue → CI verde sobre un consumer de fixture con un `gh` simulado de forma real; los objetos creados (issue, rama, PR) pasan los gates del consumer y una receta corrupta (regex, etiqueta inexistente) hace fallar el recorrido. Vive en el repo del patrón (`testing/suites/ci-pattern-docs-walkthrough`), no en la skill: la ejecuta el CI del patrón. (#305)
- **HR-48 — Toda la documentación de un proyecto `MUST` vivir en su propio repositorio**, bajo `docs/` por materia y con rutas relativas al repo: los post-mortems en la ruta del parámetro 9 (`docs/postmortems/<AAAA-MM-DD>-<slug>.md`). Fuera del repositorio, y sin copia en él, quedan solo los secretos, los backups, los datos en bruto y el material privado. Una instrucción que mande documentación a una ruta absoluta o externa al repo (otra unidad, un recurso de red o un directorio de usuario) `MUST NOT` propagarse ni adoptarse: el gate de las fases 3 y 5 de `adoption check` la rechaza nombrando fichero y línea, y la sección «Dónde vive la documentación» del documento operativo —y del bloque de `AGENTS.md`— la emite el generador. Las convenciones compartidas entre proyectos viven en el catálogo de team-skills y se propagan; la documentación de cada proyecto, no. (decisión del operador, 2026-10-06; `DysTelefonica/cadete` `AGENTS.md` «Documentation location», DysTelefonica/team-skills#319; gate: `assets/bin/ci-pattern adoption check --phase 5` y `--phase 3`)

- **HR-49 retirada (2026-10-15)**
- **HR-50 retirada (2026-10-15)**
- **HR-51 retirada (2026-10-15)**
- **HR-52 retirada (2026-10-15)**
- **HR-53 — Toda issue se triagea por causa raíz antes de tocar código y se cierra contra un test con nombre.** La clasificación `MUST` ser exactamente una de cinco clases —(A) cubierta por un cambio en curso, nombrando el cambio y su test; (B) duplicada de una clase conocida, nombrando la issue canónica; (C) bug nuevo, asignado a su clúster de causa raíz; (D) feature; (E) poco clara, se pregunta al autor sin suponer—, y el juicio de clasificación, la unicidad de raíz y la prueba de sobreingeniería son `documented-only`, sin gate (HR-34). Dos o más issues con la misma raíz se resuelven con UN solo fix en la raíz que las cierra todas, cada una con su test con nombre. Un PR que cierra una issue `MUST` nombrar el test que demuestra el comportamiento corregido, y ese test `MUST` existir en el árbol: un «debería estar arreglado» no cierra nada, y una issue que mezcla dos fallos de los que solo se resuelve uno no se cierra —se comenta qué se resolvió y qué sigue pendiente—. Si el fix añade un estado, un verbo, un flag, un gate o una representación paralela de un dato existente, el PR `MUST` justificar por qué no basta con borrar o relajar algo. Un gate publicado sin el cambio que permite cumplirlo bloquea a todos sus consumers y es incidente, no fricción (HR-21). El mecanismo que describe una issue es una hipótesis y el síntoma es la evidencia: `MUST` reproducirse el síntoma antes de implementar, y un test escrito contra el mecanismo descrito que pasa en main sin cambios invalida el diagnóstico, no la issue. (decisión del operador, 2026-10-06; DysTelefonica/team-skills#334; gate: `assets/pr-contract/check_pr_contract.py`)
- **HR-54 — La cadena de PRs se declara como dato y el enlace de la issue no admite mezclas.** Todo PR `MUST` declarar la sección **Chain Context** como dato: los que llevan `chain:partial` y **todo PR cuya base es la rama por defecto** —la punta de una cadena o un PR suelto, que el gate no distingue por el cuerpo y por eso exige la sección a los dos—; un PR suelto declara `position: 1/1`, `depends-on: none` y `follow-up: none`. La sección lleva los campos `chain`, `position` (N/M), `base`, `depends-on`, `follow-up`, `starts-at`, `ends-with` y `review-budget` (líneas añadidas + borradas / 400) y un diagrama que marca el PR actual con exactamente un `📍`: el gate verifica que `base` coincida con el `baseRefName` del PR, que `position` y `review-budget` tengan forma N/M y que haya un solo marcador. El cuerpo `SHOULD NOT` cerrar y referenciar la MISMA issue a la vez (`Closes #N` junto a `Refs #N`; `issue-spec` lo avisa y se aplica `Closes`, #787) y `MUST NOT` enlazar una issue de otro repositorio: el enlace apunta al repositorio del PR. Regla escrita (DysTelefonica/team-skills#633): `Refs`/`Closes` solo apuntan a issues del mismo repositorio, nunca a PRs; la PR de la que se depende va en `depends-on` de Chain Context; las referencias dentro de spans de código o bloques no cuentan como closing keywords, como en GitHub (gate: `assets/issue-spec/check_issue_spec.py`, con modo `local --body-file` sin payload, token ni red). La regla 7 —secciones obligatorias declaradas por el consumer en un fichero de configuración— se publica en el mismo gate (DysTelefonica/team-skills#644): opt-in con `--required-sections <path>` en los modos `local` y `pr-event`; sin configuración, el comportamiento del asset no cambia. Las secciones obligatorias son títulos Markdown (H1..H6) y no cuentan dentro de un span de código o de un bloque cercado, igual que las closing keywords; en `pr-event` la regla 7 se aplica al cuerpo de la issue consultada, no al cuerpo del PR. El presupuesto de revisión limita **cómo se corta** el trabajo, nunca el código: queda prohibido quitar comentarios, líneas en blanco, documentación o tests, y comprimir código, para caber; si ninguna pasada honesta de corte cabe, se entrega el mejor corte con `size-exception-reason:` (HR-8), y el borrado íntegro de un fichero es un caso válido de esa excepción. Cada commit es una unidad entregable —comportamiento con sus tests y su documentación— convertible en un PR de la cadena, y el PR declara su frontera de reversión (qué desaparece al revertirlo). Las ramas publicadas se sincronizan con `git merge origin/main`, **nunca** con rebase; un PR solo está listo con el CI en verde sobre su último commit, y si el consumer no ejecuta CI en PRs cuya base no es main la receta documenta cómo forzarlo (push, no retarget); una cadena en conflicto con main no permite declarar lista su punta. Los checks obligatorios se leen de la política real del host (`required_checks` y su clase, #303), no de la costumbre: si su obligatoriedad se desconoce, el PR no está listo para mergear. (gate: la sección Chain Context, el `base` contra `baseRefName`, el marcador único y el enlace a otro repositorio, en `assets/pr-contract/check_pr_contract.py`; **también es gate la regla 7 cuando el contrato la declara**: con `tracker_branch` en el contrato del host, el detective post-push (`assets / runner-controls / check_push_compliance.py` / HR-49) exige que el PR mergeado en una rama protegida tenga esa rama como cabeza —lo que llega a `main` viene del PR del tracker— y un `tracker_branch` vacío falla cerrado; documented-only, sin gate (HR-34): la elección entre los dos modelos de cadena y sus compromisos —**apilada hacia main** (cada PR con su base en el anterior; no exige contrato y mergea PR a PR, a cambio de que el orden y las bases solo se lean del cuerpo) o **rama de feature con PR tracker en borrador** (concentra la revisión en un PR y exige declarar su rama en el contrato del host, a cambio de un paso más de gobierno)—, la honestidad del corte, la unidad de commit con su frontera de reversión, la sincronía con main y la lectura de los checks obligatorios son juicio verificable en revisión, no en binario.) (decisión del operador, 2026-10-06; DysTelefonica/team-skills#333; gate: `assets/pr-contract/check_pr_contract.py`)
- **HR-55 — El `AGENTS.md` de un consumer propagado es título + bloque propagado + como máximo una sección local.** Tras la adopción (#341), el fichero `MUST` constar de: una línea de título, el bloque entre marcas `personal-skills:slice:<consumer> @ v…` cuyo cuerpo, byte a byte, tiene el sha256 `slice_block_sha256` declarado en `.team-skills.yaml`, y —como mucho— UNA sección local `## ` después del cierre. Cualquier texto antes del bloque (que no sea el título), entre el cierre y la sección, o una segunda sección `MUST` moverse a `docs/` o a esa única sección local; los enlaces de la sección local `MUST` existir en el repo (incluidos los skills del tier en `.agents/skills/`). La activación NO es autoselección: depende de la declaración explícita `agents_md: "propagated"` en el contrato de adopción. Sin declaración, el gate informa en `stderr` que la estructura todavía no se exige y sale 0 (la adopción previa no se rompe); con declaración aplica completo, aunque el AGENTS no lleve el bloque. El fail-open de que ningún consumer se declare lo cierra por fuera el doctor de flota, que exige el bloque a todo consumer con `governance: ci-pattern` en el registro, aunque no lo declare. (decisión del operador, 2026-10-06; DysTelefonica/team-skills#341; gate: `assets/bin/ci-pattern` (`_check_agents_structure`, `adoption check --phase 3` / `--phase 5`) con fixtures en `assets/adoption/tests/test_adoption_check.py`)
- **HR-56 — Lo que compone la propagación no lleva nombres propios, rutas personales ni texto de mantenimiento.** Partials, fragmentos y bloques generados nombran roles («cualquier mantenedor», «el autor del PR»), nunca logins ni personas, y citan rutas instaladas (`.agents/skills/<skill>/…`), nunca `personal/<usuario>/…` ni rutas de usuario como `C:\Users\…` o `/home/…`. Las instrucciones de mantenimiento del catálogo viven entre `<!-- maintenance:begin -->` y `<!-- maintenance:end -->` y el compositor las excluye: lo que no se propaga no está sujeto al control. El gate `MUST` fallar nombrando `fichero:línea` ante cualquier login de `maintainers` en `fleet/registry.json`, ruta `personal/<x>/` o ruta de usuario en la región propagable; sin esa lista declarada el control falla también, porque sin dato no hay medición (HR-3). (decisión del usuario, 2026-10-07; DysTelefonica/team-skills#358; gate: `assets/bin/ci-pattern` (`_check_slice_content`), exclusión en `scripts/lib/compose-slice.ps1`, fixtures en `assets/adoption/tests/test_adoption_check.py`)
- **HR-57 — Una migración versionada ya en la rama base no se renombra, borra ni edita.** En consumers con migraciones SQL versionadas el runner registra las aplicadas por nombre de fichero: renombrar, mover, borrar o editar una migración ya aplicada hace que el próximo deploy re-ejecute su DDL o la pierda (#431). El gate compara el árbol de migraciones de HEAD contra la rama base (`--base`, `origin/main` por defecto) sobre `--migrations-dir` y `--pattern` declarados: **rename**, **delete** y **edit** son rojos; las migraciones nuevas requieren prefijo numérico **estrictamente mayor** al máximo de la base; los prefijos ya duplicados en la base se toleran, los nuevos duplicados se rechazan. Exit codes: 0 limpio · 1 hallazgos · 2 fail-closed (no es git, dir ausente, patrón sin matches, base no resuelta). El gate se distribuye como asset copiado y `MUST NOT` invocar la CLI del patrón como subproceso (lección fork-bomb #342); el recursado se cierra por identidad con `CI_PATTERN_SMOKE_TESTS` (HR-45 / 46 / 48). (decisión del usuario, 2026-10-12; DysTelefonica/team-skills#431; gate: `assets/migrations-immutable/check_migrations_immutable.py`, fixtures en `assets/migrations-immutable/tests/test_migrations_immutable.py` y `test_migrations_immutable_new.py`)
- **HR-58 — Toda PR que toca una superficie de UI lleva su evidencia e2e o una exención justificada.** En consumers web, una PR que cambia UI sin flujo e2e degrada la promesa «cada cambio de UI trae su prueba de navegador» a una norma de revisión. El gate `MUST` detectar el cambio de UI por globs declarables (`--ui-glob`) menos una lista blanca de rutas no-UI (`--non-ui-prefix`, detección invertida), y exigir en el rango base..head un fichero de flujo e2e **añadido o actualizado** bajo `--e2e-dir` matcheando `--e2e-pattern`, o una sección de exención (`--exemption-section`) **no vacía** en el cuerpo de la PR (`--pr-body-file`): una sección vacía no exonera. Exit codes: 0 sin UI, con e2e o con exención justificada · 1 UI sin evidencia ni exención · 2 fail-closed (no es git, git falla, cuerpo inexistente). El gate se distribuye como asset copiado y `MUST NOT` invocar la CLI del patrón como subproceso (lección fork-bomb #342); es detectivo —vale en repos privados Free sin protección de rama ni rulesets— y no usa funciones de pago del host. La invariante propagada vive en el partial `web.md` (§Evidencia e2e por PR). (decisión del usuario, 2026-10-12; DysTelefonica/team-skills#425, piloto ardelperal/APAP_WEB#1487; gate: `assets/e2e-evidence/check_e2e_evidence.py`, fixtures en `assets/e2e-evidence/tests/test_e2e_evidence.py`)
- **HR-61 — El perfil de runners se declara, gatea y mantiene en ci-pattern.yaml.** El parámetro `P51_runners_profile` (`hosted` default | `self-hosted` para repos privados Free) gobierna el pool de runners: con `self-hosted`, todo job de `pull_request` o `push` debe correr en el pool de `P52_runners_pool_labels` o tener una excepción `P53_hosted_exceptions` con motivo no vacío. Sin perfil declarado, el gate es inocuo (retrocompatibilidad). Un job `runs-on` dinámico (`${{... }}`) o sin `runs-on:` es hallazgo, no silencio. El gate es fail-closed (HR-3): `ci-pattern.yaml` ausente o P51 inválido → exit 2. (gate: `assets/runner-profile/check_runner_profile.py`)
- **HR-62 — La adopción del perfil self-hosted incluye guía ARM64 y alerta de consumo.** No es ejecutable en CI: documenta la compatibilidad del pool ARM64 (setup-python, setup-node, Playwright/Chromium, Postgres) y el snippet de alerta temprana de consumo de minutos hosted. (decisión del operador; guía en `references/runner-profile.md` y Fase 4 de porting-guide.md)
- **HR-63 — Todo consumer que declare `docs_map` en su contrato de adopción `MUST` tener un mapa del código hub+spokes** para que cualquier agente sepa de un vistazo dónde hacer su trabajo: un hub (`docs/CODEBASE-GUIDE.md`: modelo mental, tabla de páginas y ruta de lectura) y spokes (`docs/codebase/*.md`, con «Si necesitas X, lee/toca Y» y la regla de colocación; las páginas negativas listan bajo «No existe en este repo» rutas que NO pueden existir y el gate falla si aparecen) con plantillas en `assets/templates/docs/`. La fase 5 de `adoption check` exige el hub, que cada enlace del hub exista y que cada spoke esté enlazado desde el hub, ejecuta `assets/docs-codeguide/check_docs_codeguide.py` en proceso (cada ruta citada existe y cada módulo de cada raíz declarada está descrito; las raíces se declaran en `docs_map.code_roots` como lista de `{root, glob}` o, para una sola, con `code_root` y `code_glob`, nunca las dos formas), exige que todo directorio de primer nivel con código (ocultos y dependencias no cuentan) caiga dentro de alguna de esas raíces o figure en `docs_map.exclude_dirs` (#805: el rojo nombra el directorio) y exige que un workflow defina el job declarado en `docs_map.job` y que `.github/required-jobs-policy.json` lo liste en `required_jobs` (check requerido por PR). La activación es por declaración explícita, como `agents_md` en HR-55: sin ella avisa por `stderr` y sale 0, para no poner en rojo a los consumers existentes; una declaración incompleta falla nombrando la clave (HR-3). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#709, varias raíces #727; gate: `assets/bin/ci-pattern` (`_check_docs_map`, `adoption check --phase 5`) con fixtures en `assets/adoption/tests/test_docs_map.py`)
- **HR-64 — Todo PR que toca código cubierto por una página del mapa del código `MUST` tocar esa página o declarar en el cuerpo una sección «Documentación» no vacía con el motivo de no cambiarla**; esa sección o el diff son la evidencia. Las checklists por tipo de cambio viven en la guía de mantenimiento del consumer (`docs/codebase/maintainer-playbook.md`, registrada en el manifiesto de HR-70) y el cuerpo del PR enlaza la de su tipo (`type:*`). El mapa «página → globs de código que cubre» y los enlaces a las checklists son dato declarado en `docs_evidence` de la política del gate, sin heurística; con la declaración presente pero inválida el gate sale 2 (HR-3), y sin declaración avisa por `stderr` y no falla. La sección vacía o solo con comentarios no exonera; que el motivo sea cierto y la página correcta es juicio de revisión, `documented-only` (HR-34). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#710; gate: `assets/pr-contract/check_pr_contract.py` (`check_docs_evidence`), fixtures en `assets/pr-contract/tests/test_docs_evidence.py`)
- **HR-65 — Los comandos que la documentación manda ejecutar `MUST` extraerse y probarse, no enumerarse.** Con la declaración `docs_commands` en el contrato de adopción (`cli_name`, `cli_command`, `job`, y opcionalmente `executed_verbs`, `dry_run_flag`, `exclude`), el job requerido `assets/docs-commands/run_docs_commands.py` extrae de `README.md`, `AGENTS.md` y `docs/**` los comandos que empiezan por el nombre de la CLI del consumer y los clasifica mecánicamente: `executed` (verbo seguro, sin placeholders ni shell compuesto, ejecutado en un directorio temporal), `parse-only` (con placeholders rellenos con un valor ficticio o verbo no seguro, ejecutado con `dry_run_flag`) y `presence-only` (shell compuesto o sin forma de probarlo sin efectos; contado para que el nivel sea visible). Un comando que el parser real rechaza es rojo con `fichero:línea`; extraer cero comandos falla (HR-3); `exclude` solo puede encoger (entrada obsoleta = `STALE`, HR-15). La fase 5 de `adoption check` valida la declaración y el job requerido pero **nunca ejecuta** los comandos ni la CLI del patrón como subproceso (lección #342); sin declaración avisa por `stderr` y sale 0. (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#711, a partir de `documented_invocation_test.go` de gentle-ai; gate: `assets/docs-commands/check_docs_commands.py`, fixtures en `assets/docs-commands/tests/test_docs_commands.py`)
- **HR-66 — Los enlaces locales de `README.md`, `AGENTS.md` y `docs/**` `MUST` resolver, y la navegación del mapa `MUST` ser bidireccional.** Con la declaración `docs_links` en el contrato de adopción, la fase 5 de `adoption check` ejecuta `assets/docs-links/check_docs_links.py` en proceso: cada enlace local (con su `#ancla` en destinos Markdown) resuelve contra el directorio de su fichero, cada spoke enlaza al hub y el hub a cada spoke, los ficheros y frases retirados no reaparecen y la deuda heredada se declara en `baseline`, que solo puede encoger (una entrada que ya no coincide con un enlace roto es `STALE`). No cuentan las URLs con esquema, las anclas de la misma página ni los bloques cercados. Recorrer cero ficheros falla (HR-3); el job `docs-links` es un check requerido por PR (`docs_links.job`, en `required_jobs`). Sin declaración avisa por `stderr` y sale 0 (como `agents_md` en HR-55). Las URLs remotas no se comprueban (sin red). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#712; gate: `assets/bin/ci-pattern` (`_check_docs_links`, `adoption check --phase 5`) con fixtures en `assets/docs-links/tests/test_docs_links.py`)
- **HR-67 — Todo consumer que declare `docs_decisions` en su contrato de adopción `MUST` llevar un registro de decisiones en `docs/decisions/`.** Un fichero por decisión (`AAAA-MM-DD-<slug>.md`) y un índice `README.md`, registrado en el manifiesto de HR-70. Cada decisión declara `estado` (`adoptada`, `rechazada`, `sustituida` o `aplazada`), `fecha` ISO, quién `decide` y el `motivo`; una `sustituida` enlaza con `sustituida-por` a otra que existe (sin ciclos ni autorreferencia) y una `aplazada` declara su `revision`. Toda cita `decisión: <slug>` de `README.md`, `AGENTS.md` o `docs/` apunta a una decisión que existe, y el índice y el directorio no pueden divergir. La fase 5 de `adoption check` ejecuta `assets/docs-decisions/check_docs_decisions.py` en proceso y, con `docs_decisions.job`, exige el check requerido por PR. Un directorio sin decisiones falla (HR-3); sin declaración avisa por `stderr` y sale 0 (como `agents_md` en HR-55). Qué merece registrarse como decisión es juicio de revisión, `documented-only` (HR-34). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#713, a partir del libro de cierre de gentle-ai; gate: `assets/docs-decisions/check_docs_decisions.py`, fixtures en `assets/docs-decisions/tests/test_docs_decisions.py`)
- **HR-68 — Todo consumer que declare `architecture` en su contrato de adopción `MUST` tener un documento de arquitectura que no pueda mentir** (`docs/architecture.md`, registrado en el manifiesto de HR-70): una tabla «Árbol de módulos» con la ruta y la responsabilidad de cada módulo de primer nivel, y un bloque cercado `architecture` (JSON) con las capas, los módulos de cada una y las dependencias permitidas; lo no permitido está prohibido. La fase 5 de `adoption check` ejecuta `assets/docs-architecture/check_docs_architecture.py` en proceso: cada módulo descrito existe bajo algún code-root, cada módulo existente está descrito, cada módulo pertenece a exactamente una capa y cada capa tiene su regla; el rojo nombra el módulo y el sentido. Un lado vacío falla cerrado (HR-3). Las reglas de dependencia las impone el linter del stack (`architecture.stack`: `python` con import-linter, `typescript` con dependency-cruiser) leyendo ese mismo bloque: la configuración versionada (`architecture.linter_config`) se deriva de él y el gate falla si diverge, si el linter no está instalado o no analizó nada (HR-3) y si un import viola una regla, con `fichero:línea` y la regla incumplida. Un consumer sin linter de imports (p. ej. PHP) declara `architecture.stack: none` y omite `linter_config`: el árbol y las capas del documento se miden igual, pero las dependencias entre capas quedan `documented-only` (HR-34); `none` con `linter_config` falla por contradictorio (DysTelefonica/team-skills#805). Las capas pueden declararse por módulo (`modules`) o por patrón de ruta (`paths`, primera coincidencia y `exempt`), y los patrones se expanden a los paquetes concretos para el linter (#751). Los code-roots se declaran como en HR-63 (`code_roots` o `code_root` y `code_glob`). El job declarado en `architecture.job` es un check requerido por PR (`required_jobs`). Sin declaración avisa por `stderr` y sale 0 (como `agents_md` en HR-55). Qué capas tiene cada consumer es decisión suya y la registra según HR-67; que la responsabilidad descrita sea cierta es juicio de revisión, `documented-only` (HR-34). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#714; gate: `assets/bin/ci-pattern` (`_check_architecture`, `adoption check --phase 5`) y `assets/docs-architecture/check_docs_architecture.py`, fixtures en `assets/adoption/tests/test_architecture.py`, `assets/docs-architecture/tests/` y el recorrido de PR `test_pr_walkthrough.py`: un módulo nuevo sin tocar `docs/architecture.md`, o sin capa, deja el job rojo nombrándolo)
- **HR-69 — `AGENTS.md` `MUST` enrutar la documentación del código, y la tabla se genera, no se escribe.** El bloque gobernado `ci-pattern-agents` lleva, cuando el consumer declara `docs_manifest`, la tabla «Dónde trabajar» (Si vas a… → lee…) hacia el hub y el spoke que toque, las páginas «No existe en este repo», la checklist del tipo de cambio, el registro de decisiones y los demás documentos activados con `route` en el manifiesto de HR-70. Las filas salen del manifiesto y del contrato del consumer (`assets/docs-manifest/check_docs_manifest.py`, `routing_rows`), solo para los documentos que tiene activados; lista, contenido y enrutado no pueden divergir. El propagador compone el bloque por consumer dentro del slice, de modo que `slice_block_sha256` y la estructura de HR-55 se mantienen; sin `docs_manifest` el bloque es byte a byte el canónico. La fase 5 de `adoption check` falla si el bloque de `AGENTS.md` deriva de su regeneración (`ci-pattern adoption render-agents-block <repo>`) o si una ruta enrutada no existe en el repo. (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#716; gate: `assets/bin/ci-pattern` (`_agents_block_for`, `_check_agents_block`, `adoption check --phase 5`), fixtures en `assets/adoption/tests/test_agents_routing.py`)
- **HR-70 — El Gobierno `MUST` tener una lista canónica de los documentos que debe haber y de lo que va dentro de cada uno.** `assets/docs-manifest/required-docs.json` declara por documento su ruta, para qué sirve, sus secciones obligatorias (títulos) y la HR que lo exige; las HR-63 a HR-69 registran en él su documento en vez de declarar el suyo aparte, y el enrutado de `AGENTS.md` se genera de este mismo manifiesto, de modo que lista, contenido y enrutado no pueden divergir. Con la declaración `docs_manifest` en el contrato de adopción, la fase 5 de `adoption check` exige que cada documento activado exista y que cada sección obligatoria exista como título Markdown y no esté vacía, y falla nombrando fichero y sección; un documento condicionado (`activated_by`) solo se exige si su declaración está activa. El consumer puede añadir documentos y secciones con `docs_manifest.extend` y nunca quitar los del canon (cualquier otra clave del contrato es hallazgo). Sin declaración avisa por `stderr` y sale 0 (como `agents_md` en HR-55); un manifiesto ilegible o inválido falla cerrado (HR-3). La calidad del texto es juicio de revisión, `documented-only` (HR-34). (decisión del dueño, 2026-10-17; DysTelefonica/team-skills#719; gate: `assets/bin/ci-pattern` (`_check_docs_manifest`, `adoption check --phase 5`) y `assets/docs-manifest/check_docs_manifest.py`, fixtures en `assets/docs-manifest/tests/test_docs_manifest.py`)

## §3 Decision Gates

| Condition | Action |
|---|---|
| Va a adoptar el patrón en un repo nuevo | GATE DE ADOPCIÓN del §1 (STOP): complete `references/porting-guide.md` fase por fase y verifique cada gate de salida antes de lanzar workers, tocar el repo destino o abrir PR. Sin checklist: nada. |
| El diff supera el presupuesto de líneas | Parta por unidad de trabajo y encadene PRs; sin etiqueta `size:exception` del mantenedor no hay excepción. |
| Va a abrir un PR intermedio de una cadena (`chain:partial`) | Palabras de cierre (`Closes`, `Fixes`, `Resolves`) NUNCA en el título ni en el cuerpo del intermedio: solo el PR punta cierra la issue. Intermedios con `Refs #<N>` + etiqueta de cadena; verifique `closingIssuesReferences` tras crear (HR-7). |
| GitHub no registró `closingIssuesReferences` tras crear el PR | Etiqueta de cadena más excepción declarada en el cuerpo; cierre la issue a mano tras el merge con comentario que lo documente. |
| El rojo exige un push de corrección | Empuje y deje que la CI se dispare sola; `rerun --failed` solo para transitorios; `workflow_dispatch` solo cuando el workflow cambió: si publica checks requeridos, resuelva el PR desde el SHA y evalúelo completo. |
| Va a disparar un evento manual o programado sobre la rama de un PR | Solo si el job evalúa el PR completo; si no, publique bajo otro nombre de contexto o falle — nunca un check requerido en verde sin evaluación (HR-30). |
| El cuerpo del PR, sus etiquetas o la issue enlazada cambian después del verde | Reejecute el gate si el host dispara algún evento para ese dato; si no lo dispara (cambio en la issue enlazada), reevalúe en el momento del merge antes de dar el verde por válido (HR-31). |
| El read-back7 muestra `allow_auto_merge` descartado (plan gratuito) y el verde del PR depende de un gate requerido | Auto-merge no disponible: aplique la jerarquía sin él —required checks bloquean, workflow programado recuerda, script versionado con deadline como último recurso— y haga el merge cuando el CI esté verde; un verde por auto-merge sin verificación no existe (HR-39). |
| Un rojo de avisos de dependencias aparece en un PR que no toca los manifiestos (parámetro 22) | Clasifíquelo como rojo externo: abra una issue con la evidencia (IDs de ejecución) y no lo corrija en ese PR; el gate corre programado sobre la rama por defecto y en los PRs que tocan los manifiestos (HR-39). |
| Un gate rojo cuya causa posible es un prerrequisito del runbook (fila, usuario, flag, secreto) | Verifique el prerrequisito con el comando del runbook antes de diagnosticar código; una premisa sin comando detrás es un defecto del runbook (HR-42). |
| Va a declarar un control —gate, auditoría, workflow programado o compensatorio | Primero el fixture violador: sin test de comportamiento que observe su veredicto de fallo no se declara como control; se clasifica `documented-only` hasta tenerlo (HR-32). |
| Llega un PR de reversión creado por la plataforma (el botón Revert genera su propia rama) | Admítalo por identidad verificable — actor y origen en el mismo repositorio, según el parámetro 16 — nunca por el patrón de su nombre; su trazabilidad es el PR que revierte, no una issue nueva: el PR de reversión referencia ese PR (HR-33). |
| Llega una rama de propagación o de bot (actualización de dependencias, propagación del catálogo) | Admítala por identidad verificable — actor exento Y rama del mismo repositorio, según el parámetro 16; su trazabilidad es el manifiesto o registro que la rama actualiza — y exija que esa identidad se pueda verificar en el consumer; el patrón del nombre por sí solo no admite ni rechaza (HR-33). |
| Falla un paso del job de lint | Ejecute el preflight completo antes de empujar, no solo el paso roto. |
| Va a añadir, renombrar o eliminar un job del workflow | Actualice los tres conjuntos en el mismo PR — jobs del workflow, `needs` del agregador y conjunto conocido por el evaluador; una clave de `needs` desconocida es violación, no se ignora (HR-29). |
| Aparece una fricción que ninguna regla cubre | Aplique el protocolo de HR-11: registrar con evidencia, arreglar por el pipeline, destilar en regla. |
| El PR sube una baseline o una entrada a un manifest declarado | El gate compara la entrada contra la copia de la rama base; la relajación solo pasa con su campo de datos de motivo y referencia, validado por el propio gate (HR-15). |
| Va a guardar un registro acumulativo entre sesiones | Sin clave de upsert: una observación nueva por entrada. |
| Debe reiniciar la aplicación de producción | Use el procedimiento del runbook de deploy; el reinicio directo redespliega el HEAD de la rama por defecto sin gate de evidencia. |
| Hereda un proceso vivo (vigía) de otra sesión | Verifique qué hace leyendo su script; cumplido su propósito, termínelo; inesperado, deténgase y reporte su contenido. |
| Se declara un incidente de producción | Hotfix: issue `type:bug` → fix en main → deploy inmediato → release PATCH → post-mortem blameless con action items como issues (HR-20, HR-21). |
| Va a esperar un resultado de CI o de otro actor | Aplique el modelo gentle-ai: required checks en la PR, sin strict; push a main más un nightly diario como red de seguridad; nunca un vigía IA. |
| Va a delegar una tarea que correrá en paralelo con otra | Asígnele worktree y rama propios en el encargo ; nunca dos actores sobre el mismo working tree. |
| Va a emitir un encargo de delegación | Llene `references/delegation-template.md` bloque a bloque desde lectura en vivo, con comando, salida y `verified_at`; una prescripción sin comando es hipótesis, no hecho. |
| Recibe SHA, conteos, rutas o superficies prescritos por el orquestador | Verifíquelos en vivo antes de ejecutar y reutilice lo que exista. |
| Va a comparar una medición local contra un baseline o gate de CI | Mida con la toolchain pineada del juez o re-mida con ella; sin pin no hay comparación válida (HR-26). |
| Un PATCH de settings respondió 200 pero el read-back no muestra el valor | Lea de vuelta dos veces con delay y verifique el plan de la org: el campo puede ser paywalled y descartarse en silencio. |
| Tests rojos solo en local y verdes en CI | Sospeche del `.env` local filtrando variables a `Settings`; aísle el entorno de tests o registre el rojo como ambiental (HR-28). |
| Va a cortar un release | Tag semver anotado + GitHub Release `Latest` + notas concisas + playbook `RELEASE-<TAG>.md` con imagen de rollback capturada antes de desplegar (HR-19). |
| El diff toca rutas de alto riesgo (parámetro 18) y no hay segundo revisor | El cuerpo del PR lleva el campo de evidencia de revisión (parámetro 19) con lente, veredicto y referencia, validado por el gate; la prosa libre no cuenta como evidencia. |

## §4 Execution Steps

### Adopción del patrón en un repo

1. **Audite.** Inventaríe cada control del CI —gate, auditoría, workflow
   programado o compensatorio— con su veredicto (se queda, se refuerza, se
   duerme, se retira), la evidencia que lo sostiene y el test que lo ve
   fallar; un control sin fixture violador se registra como
   `documented-only`, no como control (HR-32). Contraste cada afirmación de
   la documentación contra los workflows reales; corrija los desvíos en la
   misma sesión. Modelo de salida: `references/gate-verdicts.md`. El
   procedimiento reproducible del patrón web —recorrido, matriz de
   enforcement, consultas de solo lectura, preguntas del patrón y salida—
   vive en `assets/audit-issue-to-merge.md`. Registre además cada artefacto
   de gobierno del destino en el contrato de adopción con su destino —
   `adopted` hacia el artefacto del patrón que lo sustituye, o `retired`—
   (HR-46): la fase 3 ejecuta esa sustitución sobre el árbol y las
   instantáneas del host y no acepta un fichero de evidencia.
2. **Lea de vuelta el host.** Capture con GETs de solo lectura las
   respuestas de la API —etiquetas, settings de merge, protección de la
   rama, rulesets— en instantáneas JSON y contrástelas con el contrato
   declarado mediante `assets/host-readback/check_host_drift.py`: cada
   regla queda clasificada `host-enforced` o `documented-only`, y la que
   no tiene enforcement en el host `MUST NOT` describirse como gate
   (HR-34). El check nunca muta el host ni prueba una escritura para
   descubrir permisos; instantánea ausente o ilegible es un rojo, no un
   skip (HR-3).
3. **Mida.** Presupuesto real por PR, duración de la CI, falsos verdes conocidos,
   pasos de CI que nadie ejecuta en local. Sin cifra no hay decisión de gates.
4. **Instale.** Extraiga los parámetros del repo (`assets/parameters.md`),
   adapte los scripts de referencia (ubicación declarada en `assets/parameters.md`),
   configure la protección de rama (checks requeridos, sin force-push) y el
   preflight canónico que lee los pasos del job de lint del propio workflow.
5. **Valide en real.** Ejecute de verdad cada gate que toque producción  y
   complete una cadena de PRs encadenados de extremo a extremo con CI en cada
   tramo antes de declarar el patrón adoptado.

### Operación por unidad de trabajo

1. Issue con el contrato canónico completo y etiqueta de aprobación aplicada
   **al crearla** — por cualquier mantenedor con permisos de administración
   del repositorio (rol, nunca un login concreto; ese mismo rol es quien
   puede lanzar el merge gobernado)—: el gate lee la issue remota, no su
   copia local. Las seis
   secciones canónicas son contrato del CUERPO DE LA ISSUE también, no solo
   del lado PR, con estos nombres exactos: `Problema y contexto`, `Evidencia
   verificable`, `Alcance y no objetivos`, `Criterios de aceptación`, `Plan
   de validación` y `Dependencias y riesgos`; una issue aprobada sin ellas
   falla issue-spec hasta aumentarse a mano (la contraparte issue-side es
   documental, no de gate; la documentación operativa que se genera en
   destino se gobierna en DysTelefonica/team-skills#279).
2. Worktree dedicado y rama propia generada según el ruleset del repo
   (parámetro 1). Cuando la tarea corre en paralelo con otro actor, el
   worktree es propio y exclusivo. El encargo al worker se llena
   con `references/delegation-template.md`.
3. Preflight completo en local antes de cada push (HR-4).
4. PR con todas las etiquetas en el comando de creación; `Closes #<issue>` solo
   en la punta de la cadena (HR-6, HR-7).
5. Verifique `closingIssuesReferences` tras crear el PR; si no aparece, declare
   la excepción en el cuerpo y cierre a mano tras el merge (HR-7).
6. CI en vuelta: clasifique rojos solo tras identificar el paso que falla vía
   API; empuje los fixes; rerun solo para transitorios (gates del §3).
7. Merge con auto-merge —si el read-back7 confirma
   `allow_auto_merge`; si no, merge manual tras el verde por la fila de §3— y
   commit de fusión, sin borrar la rama remota; con `strict`, actualice la
   rama con la secuencia determinista de HR-17.
8. Deploy: registre el veredicto de la batería sobre la URL de salud del
   parámetro 6 de la revisión desplegada (HR-10); apague cualquier flag de
   prueba al terminar, también si la batería falla.

**Carril docs-only.**

Los PRs que solo documentan pueden usar un carril barato del consumer, con
regla estricta y fail-closed: TODOS los ficheros cambiados `MUST` ser de
documentación (la lista de rutas que cuentan como docs es dato declarado del
consumer). El marcador del carril es fail-closed: un PR marcado docs-only
que toque un fichero fuera de las rutas declaradas es una violación que el
agregador rechaza, nunca un atajo silencioso. Los eventos de release son
inmunes: el corte de release y el hotfix no usan el carril (HR-19, HR-20).
Un carril sin esa verificación del agregador es `documented-only` (HR-34),
no gate. (evidencia: ardelperal/APAP_WEB#1196 y #1202)

**Ramas generadas por la plataforma.**

El botón Revert crea su propia rama de reversión, y cada ecosistema de
actualización de dependencias habilitado en el repositorio abre las suyas;
ninguna casa con el patrón canónico (parámetro 1), y un prefijo de texto no
puede conceder la exención (HR-33). Para cada fuente de ramas de plataforma,
con los actores y patrones concretos como datos del parámetro 16:

1. **Admisión por identidad.** El gate de nombre y el gate de trazabilidad
   admiten la rama solo cuando el actor figura en la lista de actores exentos
   y la rama pertenece al mismo repositorio que el gate juzga; un prefijo
   exento con actor no exento, o un actor exento con rama de otro repositorio,
   es una violación ordinaria. Las ramas de propagación del catálogo
   (`skill-fleet/<consumer>`) se rigen igual: su identidad debe poder
   verificarse en el consumer, no presumirse del nombre.
2. **Trazabilidad.** El PR de reversión no abre issue: su trazabilidad es el
   PR que revierte, al que referencia. Las ramas de un ecosistema de
   actualización trazan contra el manifiesto o registro que la rama actualiza,
   que es el cambio verificable de la unidad de trabajo; el gate de issue-spec
   evalúa eso en lugar del contrato issue-first, nunca lo omite en silencio.

### Release, hotfix y post-mortem

**Corte de release (semver, estilo gentle-ai).**

1. Confirme que main está verde contra la base actual y sin rebase pendiente;
   el verde contra una base obsoleta no autoriza el release.
2. Cree el tag anotado `vMAJOR.MINOR.PATCH` y la GitHub Release asociada
   marcada `Latest`; los prereleases usan `vX.Y.Z-rc.N` y quedan excluidos de
   la estable con el patrón `v*-*` (HR-19).
3. Escriba notas concisas: qué cambió y enlace al issue o post-mortem; la
   causa raíz completa vive en su doc dedicado, nunca dentro de las notas (HR-21).
4. Ejecute el playbook de deploy: `RELEASE-<TAG>.md` con build/push si aplica,
   apply, rollout, verificación y rollback; capture la imagen de rollback
   antes de desplegar.

**Hotfix (incidente de producción).**

1. Abra la issue canónica (`type:bug`) con la evidencia del incidente.
2. Aterrice el fix en main por el pipeline normal (issue, worktree, PR, CI,
   merge); sin rutas paralelas ni fixes que vivan solo en una rama (HR-20).
3. Despliegue inmediatamente tras el merge y registre el veredicto sobre el
   SHA desplegado (HR-10).
4. Corte el release PATCH con notas que enlacen al issue (HR-19).
5. Escriba el post-mortem blameless en la ruta de post-mortems del consumer
   (parámetro 9) con secciones Timeline (UTC) / Impact / Root cause / What
   worked / What failed; causas de sistema, nunca personas (HR-21).
6. Abra cada action item como issue de GitHub con owner asignado y verifique
   que no queden solo en el doc (HR-21).

### Destilación de fricciones (protocolo de mejora continua)

1. Detecte con evidencia: sin incidente citado no es fricción, es opinión.
2. Regístrela inmediatamente y distinga el destino: fricción local del
   consumer → su tracker; defecto atribuible a la skill →
   `DysTelefonica/team-skills` vía `report-friction` (clasificación,
   saneado y deduplicación por huella; ver `references/cli-spec.md` §8).
   Clasifique como bloqueante o de seguimiento; DysTelefonica/team-skills#279 define el
   ciclo evidencia → triage → corrección (documentación operativa generada en destino)
   y remite a este comando para el tránsito hacia la skill.
3. Arregle por el pipeline normal (issue, worktree, PR, CI, merge), nunca
   mezclada con otro cambio.
4. Destile en regla con su ternario completo: regla, rojo que la creó,
   procedimiento correcto.
5. Vigile la recurrencia: la segunda ocurrencia se automatiza (gate, workflow
   programado o cambio de diseño), no se arregla a mano otra vez.
6. Abra la contraparte en la skill: todo fix de CI genera DOS issues — una
   en el repo afectado (arregla el CI) y otra en la skill canónica (arregla
   el patrón) — (directiva permanente del operador, 2026-10-02, registro
   Engram #8138). Una fricción rastreada solo en el repo deja el patrón sin
   corregir y el siguiente repo o IA la vuelve a pagar.
7. Audite periódicamente: una pasada que compare cada afirmación de la doc
   contra la realidad de los workflows y declare qué claims quedaron verificados.

## §5 Output Contract

El contrato se declara por modo de ejecución: la salida declara `mode` y solo
lleva las keys aplicables a ese modo; las keys de otros modos van ausentes,
no vacías.

| Key | Type | Description |
|---|---|---|
| `mode` | `"adoption" \| "operation" \| "audit"` | Modo declarado; determina qué keys aplican. |
| `adoption_gate` | `"checklist_complete" \| "blocked"` | Solo `adoption`: estado del GATE DE ADOPCIÓN del §1 — checklist fase por fase verificado, o adopción bloqueada con la fase pendiente. |
| `pattern_params` | object | `adoption`, `audit`: los parámetros del repo resueltos según `assets/parameters.md`. |
| `gate_verdicts` | array | Solo `audit`: `{gate, verdict, evidence}` por cada gate auditado. |
| `frictions_registered` | string[] | `operation`, `audit`: fricciones registradas con su evidencia (issue, PR, run). |
| `frictions_distilled` | string[] | Solo `operation`: reglas destiladas con su ternario completo. |
| `evidence_shas` | array | Solo `operation`: `{sha, context, state}` — estados de commit registrados por revisión. |

| `preflight_command` | string | Solo `operation`: comando canónico ejecutado antes del último push. |
| `hr_traceability` | array | Solo `audit`: pares `{rule, evidence}` que trazan cada HR-N aplicada a su fricción de origen. |
| `risks` | string[] | Todos: riesgos abiertos (gates sin medir, contradicciones doc-workflow pendientes). |
| `next_recommended` | `"fix_violations" \| "retry_checklist" \| "register_friction" \| "none"` | Todos: siguiente acción — enum propio de esta skill (hallazgos de auditoría → `fix_violations`; GATE DE ADOPCIÓN incompleto → `retry_checklist`; fricción sin registrar → `register_friction`; nada → `none`). |

## §6 Anti-patterns

| Symptom | Fix |
|---|---|
| Etiquetas aplicadas después de crear el PR y el gate queda rojo | Etiquete en `gh pr create --label`; si ya creó el PR, use la API REST y espere rerun o push. |
| `Closes` en el título o en un tramo intermedio cierra la issue antes de tiempo | Palabras de cierre solo en el cuerpo del PR punta; intermedios con `Refs` y etiqueta de cadena (ardelperal/APAP_WEB#1138). |
| Verde en local y rojo en CI por un paso de lint que el runner de tests local no ejecuta (parámetro 13) | Preflight canónico que reproduce el job completo; ejecútelo antes de cada push. |
| Gate que imprime «`OK (0 hallazgos)`» sin haber corrido el herramienta | Fail-loud: el gate falla si el herramienta no está instalado o no midió (HR-3). |
| Un rojo de raíz enmascarado como N fallos en el agregador | El agregador separa causa raíz de skips en cascada y agrupa las consecuencias. |
| Job cableado en `needs` y ausente del conjunto conocido falla sin que el agregador lo note | Paridad three-way verificada: `jobs(workflow) − {agregador}` = `needs` = conjunto conocido; una clave de `needs` desconocida cuenta como violación y no se ignora (HR-29). |
| Un check requerido sale en verde desde un evento manual o programado que nunca evaluó el PR | Nombres requeridos solo desde el evento que evalúa; en cualquier otro, otro nombre de contexto o fallo — nunca éxito sin evaluación (HR-30). |
| El verde sobrevive a la edición del cuerpo del PR o a un cambio de etiquetas de la issue enlazada | Disparador para el dato mutable donde exista y reevaluación en el merge donde no; un gate sin relectura del dato no certifica nada (HR-31). |
| El mismo rojo de avisos aparece a la vez en ramas sin relación —siete ejecuciones del gate de avisos en cuatro ramas dentro de 80 minutos (2026-09-29/10-01; ejecuciones 36737166885, 36739875318, 36741326730, 36747068968)— | El veredicto depende de la base de avisos del momento, no del diff: clasifíquelo como rojo externo, abra issue y re-agende el gate a programado más PRs que tocan manifiestos (HR-39). |
| Test que «prueba» el control afirmando subcadenas del fuente del script o del workflow | Ejecute la lógica del control contra un fixture que viola la regla y observe el veredicto de fallo; la prueba de subcadenas no cuenta como evidencia (HR-32). |
| Exención de gate concedida por un prefijo del nombre de rama: renombrar la rama basta para saltarse el gate | Exención solo por identidad verificable — actor exento Y origen en el mismo repositorio, como datos del parámetro 16; el prefijo del nombre no es identidad y la exención que concede solo es una brecha (HR-33). |
| PR de reversión de la plataforma bloqueado por el gate de nombre en el momento de más prisa | El botón Revert genera su propia rama: admítala por identidad verificable según el parámetro 16, con el PR revertido como trazabilidad; exigirle el patrón canónico bloquea el revert sin proteger nada (HR-33). |
| Control compensatorio documentado y nunca visto fallar | Mismo baremo que un gate: fixture violador y veredicto de fallo observado; sin eso es `documented-only`, no control (HR-32). |
| Etiqueta documentada que no existe en el host: la doc ordena aplicarla y ningún gate la encontrará jamás | Contraste con la API del host mediante el drift check de solo lectura (`assets/host-readback/`); la etiqueta que el host no tiene se crea o se borra de la doc en la misma sesión (HR-34). |
| Bucle `--watch`, watcher de sesión o vigía IA esperando CI | Mecanismo primero: auto-merge armado más `allow_update_branch` —con disponibilidad verificada por el read-back7—; donde no hay mecanismo, script versionado con deadline y fallback; sin procesos vivos de sondeo, ni siquiera como vigía que sobrevive a la sesión. |
| Variable global que aprueba o bloquea deploys para siempre | Estado de commit por SHA; la evidencia viaja con la revisión (HR-10). |
| Fix de fricción aplicado a mano sin registrar | El rojo vuelve con la próxima sesión; registre con evidencia y destile en regla. |
| Registro acumulativo guardado con upsert por clave | Una observación nueva por entrada; el upsert sustituye el contenido. |
| Rerun de un run cuyo SHA ya está corregido en local | Empuje; el rerun reproduce el SHA original y quema un ciclo de espera. |
| Batería ejecutada con los tests de la rama por defecto contra un deploy anterior | Worktree en la revisión desplegada; los fallos por desfase se clasifican, no se registran como fallo del deploy. |
| Baseline de un ratchet editada a mano en cada reducción | Comando de regeneración de baseline; el gate solo rechaza entradas nuevas (HR-15). |
| El mismo PR que el gate juzga sube la baseline o añade una entrada a la allowlist | El gate compara cada entrada editable contra la copia de la rama base; la relajación exige su campo de datos con motivo y referencia (HR-15). |
| Causa raíz enterrada en las notas de release o en el cuerpo del PR | Notas concisas con enlace; el post-mortem vive en su doc dedicado (HR-19, HR-21). |
| Fix de incidente viviendo solo en una rama paralela esperando el release regular | Hotfix aterriza en main, deploy inmediato y release PATCH (HR-20). |
| Post-mortem que nombra personas como causa | Blameless: causas de sistema; falló el proceso, no la persona (HR-21). |
| Deploy sin imagen de rollback capturada | Capture la imagen de rollback en el playbook antes de aplicar. |
| Bucle `--watch`, watcher de sesión o vigía IA esperando CI | Mecanismo primero: auto-merge armado más `allow_update_branch` —con disponibilidad verificada por el read-back7—; donde no hay mecanismo, script versionado con deadline y fallback. |
| Dos workers sobre el mismo working tree | Un worktree por actor concurrente; el síntoma es el commit sobre la rama ajena a mitad de vuelo. |
| Prescripción del orquestador ejecutada sin verificarla | El snapshot puede ser stale: verifique SHA, conteos y rutas en vivo y reutilice lo existente. |
| Encargo que nombra repo, superficies o SHAs sin comando de verificación | Use `references/delegation-template.md`: cada dato con comando y salida en vivo; el worker reporta el desajuste en vez de adaptarse. |
| Medición local comparada contra el juez de CI sin toolchain pineada | Pin de versiones en la stack de medición o re-medición por la toolchain del juez (HR-26). |
| PATCH de settings dado por bueno por su 200 OK | Read-back doble con delay y verificación del plan: paywalled se descarta sin error. |
| Suite que solo pasa con el `.env` del desarrollador delante | Aísle el entorno de tests del `.env` local y documente los rojos ambientales conocidos (HR-28). |
| Rojo de CI diagnosticado grepeando logs del runner en vez de pedir la evidencia por paso | `gh api repos/<org>/<repo>/actions/jobs/<id>` separa el paso que falla; el log del runner equivocado (hosted frente a self-hosted) fabricó el diagnóstico «Docker daemon» ×3 (C2 del porting-guide). |
| Premisa de propagación o de gobernanza del destino tomada de su documentación | Audite el mecanismo real en vivo (hooks, reconciliador, markers, manifest) antes de depender de él; la premisa falsa dejó espejos stale medio tramo (C4 del porting-guide). |
| Adopción del patrón lanzada sin el checklist de pre-vuelo | GATE DE ADOPCIÓN del §1: sin `references/porting-guide.md` completado no hay workers, ni toques al repo destino, ni PR (anclaje: Cadete, 2026-09-30/10-01). |
| Re-implementar autoridad de review en el patrón (linajes, recibos, rondas de jueces, consentimientos como reglas de CI) | Es capacidad del harness (`gentle-ai`); el patrón la asume instalada, no la duplica ni la documenta como propia. |
| Lente de revisión exigida solo en prosa (ningún gate la lee) | Para diffs de alto riesgo, campo estructurado de evidencia (lente, veredicto, referencia) en el cuerpo del PR, validado por un gate; la prosa libre no es evidencia. |

## §7 Companion skills

| Skill | Qué posee | Cargar junto cuando |
|---|---|---|
| `repository-delivery-governance` | Auditoría genérica del repositorio y matriz de enforcement (policy documentado, validado en local, CI, host, manual) | Vaya a auditar o cambiar la protección de rama, checks requeridos o política de merge. |
| `deterministic-quality-harness` | Ratchets, identidad estable de hallazgos y assets de gates (p. ej. `check_pr_size.py`, policy files) | Vaya a instalar o ajustar el motor de un gate. El qué y cuándo de la operación es esta skill. |
| `ci-pattern` (esta skill) | Operación por unidad de trabajo, evidencia por SHA, cadena de PRs y release | — |
| `branch-pr` | Ciclo de apertura y fusión del PR | Vaya a abrir o fusionar el PR de una unidad de trabajo bajo este patrón. |
| `chained-pr` | Encadenado de PRs cuando el diff supera el presupuesto | El diff supera el presupuesto y va a encadenar PRs. |
| `work-unit-commits` | Partición de commits por unidad de trabajo | Vaya a planificar los commits del PR por unidad de trabajo. |
| `documentation-alan-style` | Documentación que refleja el código real | Vaya a actualizar la documentación del CI para que refleje el código. |
| `oracle-vps-github-runners` | Provisión y operación de runners self-hosted en VPS Oracle: restricciones de arquitectura y sistema del pool (la réplica minio es amd64-only, no arranca en arm64; los jobs Windows nunca son movibles — su HR-15) y receta de siete pasos de migración de repo | Vaya a mover jobs del patrón a runners self-hosted o a auditar por qué un job no es movible entre runners del pool (ardelperal/APAP_WEB#1204, #1214). |

Frontera con `deterministic-quality-harness`: su asset `check_pr_size.py` y el de
este patrón (`assets/pr-size/`) aplican la misma regla: la única excepción al
presupuesto es la etiqueta `size:exception` puesta por el mantenedor (decisión
del dueño, #806; `slices/partials/web.md` del catálogo `DysTelefonica/team-skills`),
y el campo `size-exception-reason:` del cuerpo no existe ni se lee. Tampoco hay
exclusiones por lockfiles o ficheros generados.

## §8 References

- `assets/branch-name/` — gate del nombre de rama del PR (parámetro P01): la regex
  es dato (`--pattern`, por defecto `<type>/<slug>` sin número de issue), exige el
  nombre completo y exime la rama por defecto, `archive/*`, `skill-fleet/*` y las
  ramas de Dependabot solo si `--actor` es el bot; lo anterior a la convención se
  lista con `--allow`. Salida 2 ante regex ilegible o rama vacía. Es el control
  compensatorio de un host sin rulesets (#806).
- `assets/release-gate/` — gate de release sobre el tag que dispara el release
  (HR-19 y HR-20): exige tag semver estricto (`vMAJOR.MINOR.PATCH`, sin ceros
  a la izquierda ni sufijo), anotado, con versión mayor que la última release
  y commit alcanzable desde la rama por defecto; corre en local con `git` y
  sin red. team-skills no publica releases: el dogfooding se sustituye por el
  test sobre repositorio git temporal, exención documentada en la matriz.
- `assets/ratchet/` — gate de ratchet shrink-only sobre hallazgos normalizados
  (HR-15 y HR-16): identidades estables (`rule`, `path`, `fingerprint`
  independiente del número de línea) contra una baseline versionada con
  `target`/`target_date`; rechaza lo nuevo, exige encoger la baseline para
  fijar lo eliminado, compara contra la copia de la rama base (toda entrada
  añadida exige campo `relaxation` con `reason`/`reference`) y trata la fecha
  vencida como hallazgo; sin prueba de vida de la medición sale por 2
  (liveness, HR-3).
- `assets/exemptions/` — gate de identidad de exenciones (HR-33): escanea
  documentos JSON (baselines, políticas, matrices) y exige que toda exención
  declarada como dato — `relaxation`, `exemption` — lleve una identidad
  verificable: `actor` no vacío o `issue` positivo; una exención sin
  responsable auditable no es gobierno.
- `assets/hr-gate-matrix/` — directorio con el gate HR-1
  (`check_structured_inputs.py`) y su suite; el nombre del directorio es
  histórico y se conserva hasta una eventual reubicación del gate HR-1
  (la matriz HR→gate y su meta-gate se retiraron por #513/#540; el bloque 2
  de `assets/audit-issue-to-merge.md` registra la nueva forma de la matriz
  de enforcement).
- `assets/required-jobs/` — primer asset ejecutable de la skill: agregador
  fail-closed de jobs requeridos con política externa (sin nombres de jobs,
  eventos ni skips en el código), política de ejemplo y suite con test de
  paridad workflow↔política; el `README.md` del asset documenta el destino
  en el consumer y el cableado de `toJSON(needs)`.
- `assets/workflow-policy/` — gate de política de workflows (HR-3 y HR-26 y HR-30): lector YAML estructural fail-closed, solo stdlib y sin red, que
  verifica `uses:` por SHA/digest, toolchain exacta, fail-loud en los jobs
  requeridos y publicación de checks requeridos solo desde eventos de PR;
  rechaza las claves duplicadas (salida 2) y exige en todo job
  `timeout-minutes`, un grupo `concurrency` (de job o de workflow), ningún
  puerto de host fijo en los servicios y un `timeout … docker info` antes del
  primer `docker run` (#806);
  la lista de jobs requeridos viene de la política de `required-jobs` y un
  job requerido ausente de todo workflow analizado es hallazgo.
- `assets/audit-issue-to-merge.md` — procedimiento reproducible de la auditoría issue→merge del patrón web (HR-10 y protocolo de mejora continua): recorrido ordenado, matriz de enforcement con el vocabulario de las cinco capas, consultas de solo lectura al host, preguntas del patrón con su HR y salida con evidencia `fichero:línea`; Markdown puro, sin script (el check de deriva ejecutable es el gate de host-readback).
- `assets/host-readback/` — asset de readback del host: contrato de
  ejemplo (`host-contract.example.json`), drift check de solo lectura
  (`check_host_drift.py`, stdlib, sin red) y suite de tests; clasifica
  cada regla `host-enforced` o `documented-only` (HR-34) contrastando el
  contrato declarado con instantáneas JSON de las respuestas de la API.
- `assets/parameters.md` — los parámetros que se extraen por repo, con los
  valores del consumer de origen como ejemplo y la ubicación de los scripts
  de referencia.
- `references/delegation-template.md` — plantilla canónica de encargo de
  delegación: repo verificado en vivo, base y tip, worktree y rama
  generada, superficies editables verificadas, hechos en vivo a confirmar por
  el worker, condiciones de parada y plazo con reporte por transición.
- `references/fricciones.md` — catálogo destilado: fricción, antídoto y evidencia.
- `references/porting-guide.md` — checklist de pre-vuelo de la adopción fase
  por fase; cada gate cita el incidente real de Cadete (2026-09-30/10-01) que
  lo justifica. Exigido por el GATE DE ADOPCIÓN del §1.
- `references/gate-verdicts.md` — veredicto de cada gate del inventario con su base.
- `references/benchmark-gentle-ai.md` — ideas transferibles y rechazadas de otro CI.
- `references/incidents.md` — ejemplo destilado de hotfix, post-mortem y
  release sobre un incidente real de pérdida de datos.
- `references/cli-spec.md` — contrato normativo de la CLI `assets/bin/ci-pattern`:
  comandos, códigos de salida, manifiesto, idempotencia y frontera honesta.
- `references/asset-inventory.md` — inventario de extracción de los activos y
  parámetros del patrón (62 activos, P01-P48, huecos G1-G12) tomado del origen
  verificado; fuente de las citas de `assets/parameters.schema.json`
  y `assets/templates/`.
- Implementación de referencia del resto del patrón: versionada hoy en
  `ardelperal/APAP_WEB` (`scripts/preflight.py`, `scripts/check_issue_specs.py`,
  `scripts/check_release_evidence.py`, `scripts/check_release_e2e_required.py`,
  `scripts/production_smoke.py` y `.github/release-e2e-paths.txt`), pendiente
  de publicarse como asset portátil de esta skill; el agregador de jobs
  requeridos ya vive aquí (`assets/required-jobs/`). Detalles en
