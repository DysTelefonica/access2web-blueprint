# Flujo operativo del repositorio

Documento GENERADO por `ci-pattern adoption generate-doc` a partir de los datos del
repositorio (ci-pattern.yaml, .github/host-contract.json, .github/required-jobs-policy.json,
formularios de issue). NO lo edite a mano: regenérelo con el mismo comando. La deriva la
detecta el gate de documentación operativa de `ci-pattern adoption check`.

<!-- ci-pattern-doc:begin issues -->
## 1. Issues

- Cree la issue con un formulario de `.github/ISSUE_TEMPLATE/`:
  - Etiquetas de tipo declaradas por el contrato del host (`.github/host-contract.json`), las que el gate acepta: `type:bug`, `type:chore`, `type:docs`, `type:feature`, `type:refactor`.
- Formularios disponibles en el destino: config.yml, issue-canonical.yml.
- Antes de crear la issue, busque duplicados en issues abiertas y cerradas.
- La implementación solo arranca con la etiqueta `status:approved` en la issue.

<!-- ci-pattern-doc:end issues -->

<!-- ci-pattern-doc:begin branch -->
## 2. Ramas

- Genere el nombre de rama; nunca lo escriba a mano (HR-35): `<tipo>/<número>-<slug>`, con el slug del título de la issue en minúsculas y guiones:
  `git switch -c <tipo>/<número>-<slug>`
- Regex obligatoria: `^(feat|fix|refactor|docs|ci|test|chore)/\d+-[a-z0-9]+(-[a-z0-9]+)*$`
- El generador que construye y valida el nombre (HR-35) es un activo de la skill del patrón: no se instala en el consumer, así que aquí se compone el nombre y se valida contra la regex.

<!-- ci-pattern-doc:end branch -->

<!-- ci-pattern-doc:begin commits -->
## 3. Commits

- Conventional Commits: `tipo[(scope)]: descripción` con tipo en feat|fix|docs|chore|refactor|test|ci|perf.
- Sin atribución de IA: sin `Co-Authored-By` ni pies de página generados.

<!-- ci-pattern-doc:end commits -->

<!-- ci-pattern-doc:begin pull-request -->
## 4. Pull requests

- Presupuesto: 400 líneas (añadidas + borradas, tests y docs incluidos). Si no cabe, cadena de PRs.
- Slice intermedio de cadena: etiqueta `chain:partial` y `Refs #N` en el cuerpo; CERO palabras de cierre.
- Con `chain:partial` o en la punta de cadena, el cuerpo lleva la sección Chain Context (HR-54): posición N/M, base igual a la del PR, dependencias, presupuesto y un único `📍`.
- Solo el PR punta de cadena lleva `Closes #N`.
- Antes de un `Closes #N`, compruebe la matriz HR→gate de la skill adoptada: `rg '"issue": <número>' <skill>/references/hr-gate-matrix.json`; una entrada `manual` apuntando a la issue exige reclasificarla en el mismo PR.
- Exactamente una etiqueta `type:*` por PR: type:bug, type:chore, type:docs, type:feature, type:refactor.
- Toda edición material de una skill sube su `metadata.version` (HR-44).

<!-- ci-pattern-doc:end pull-request -->

<!-- ci-pattern-doc:begin required-checks -->
## 5. Checks requeridos

- Jobs requeridos del agregador: codeql, quality, required, review-budget, security. Tabla generada 1:1 desde ellos: ningún check puede fallar por una regla que no esté aquí.
- Contextos requeridos declarados: required.
- Checks del contrato de host: required.

| `codeql` | verificación del agregador del patrón | `bash scripts/local-preflight.sh` | → ejecute el preflight (HR-4), lea el fallo y corrija la causa responsable; si la regla no está en este documento, deténgase y pregunte |
| `quality` | verificación del agregador del patrón | `bash scripts/local-preflight.sh` | → ejecute el preflight (HR-4), lea el fallo y corrija la causa responsable; si la regla no está en este documento, deténgase y pregunte |
| `required` | verificación del agregador del patrón | `bash scripts/local-preflight.sh` | → ejecute el preflight (HR-4), lea el fallo y corrija la causa responsable; si la regla no está en este documento, deténgase y pregunte |
| `review-budget` | verificación del agregador del patrón | `bash scripts/local-preflight.sh` | → ejecute el preflight (HR-4), lea el fallo y corrija la causa responsable; si la regla no está en este documento, deténgase y pregunte |
| `security` | escaneo de seguridad de superficie | `bash scripts/local-preflight.sh` | → atienda el hallazgo o justifíquelo en el PR; nunca silencie el scanner sin causa responsable |

<!-- ci-pattern-doc:end required-checks -->

<!-- ci-pattern-doc:begin merge -->
## 6. Merge

- El merge es decisión humana con CI en verde; ningún agente mergea.
- Métodos de merge permitidos (host-contract): merge commit, rebase, squash.

<!-- ci-pattern-doc:end merge -->

<!-- ci-pattern-doc:begin release-hotfix -->
## 7. Release y hotfix

- Todo release es un tag semver anotado (HR-19).
- Todo hotfix aterriza primero en main (HR-20).
- Contextos de producción: e2e `no aplica: el repositorio no despliega en producción`, smoke `no aplica: el repositorio no despliega en producción`; runbook: `docs/11-releases.md`.

<!-- ci-pattern-doc:end release-hotfix -->

<!-- ci-pattern-doc:begin incidents -->
## 8. Registrar un incidente

Un incidente de producción deja cuatro niveles, enlazados por identificadores (ID de incidente, número de ticket, SHA, ejecución de CI, issues y PRs); el CONTENIDO de los datos personales nunca entra en el repo.
- **1. Post-mortem** (parámetro 9): `docs/postmortems/<AAAA-MM-DD>-<slug>.md`, blameless, con causas de sistema y cada action item como issue con owner. Aquí sí van identificadores y nombres de sistemas: qué pasó, por qué y qué lo evita.
- **2. Registro operativo**: `docs/incidents/<AAAA-MM-DD>-<slug>.md` —todo el detalle necesario para seguir atendiendo los tickets abiertos (IPs, hostnames, comandos, logs, cronología fina y el estado con el siguiente paso de cada ticket)—. Con el parámetro por defecto vive en este mismo repositorio; el post-mortem lo enlaza por ruta o por identificador.
- **3. Secretos**: credenciales y tokens viven en GitHub Actions repository secrets, NUNCA en git, ni siquiera en un repositorio privado. El post-mortem y el registro los citan por su identificador en el gestor.
- **4. Datos personales y conversaciones con terceros**: se quedan en el sistema de tickets y se referencian por su número; nunca se copian al repositorio.
- La evidencia se cita con su `sha256` y su ubicación de acceso controlado. Lo que se redacta se marca de forma explícita (`[IP interna redactada]`): nunca se borra en silencio.
- El gate de esta fase rechaza secretos y datos personales en el post-mortem y en el registro operativo, y la evidencia declarada sin hash ni ubicación. Las IPs, los hostnames y los comandos internos están permitidos.

<!-- ci-pattern-doc:end incidents -->

<!-- ci-pattern-doc:begin friction -->
## 9. Fricciones

- Registre toda fricción con evidencia (HR-11); en la segunda ocurrencia, etiqueta `friction:recurrent`.
- El ciclo evidencia → triage → corrección de la skill sustituye a toda receta de operación local del consumer.

<!-- ci-pattern-doc:end friction -->

<!-- ci-pattern-doc:begin lifecycle -->
## 10. Ciclo de vida de un cambio

Cada etapa dice qué hacer, con qué comando y cómo saber que está hecha. Jobs requeridos que debe ver en verde antes del merge: codeql, quality, required, review-budget, security.
- **1. Issue.** Con el formulario del consumer (sección 1), busque duplicados y cree la issue:
  `gh issue list --search "<términos>" --json number,title,state --state all --limit 20`
  `gh issue create --title "<título>" --body-file <cuerpo-de-la-issue.md>`
  Hecho: la issue existe y lleva su etiqueta de tipo.
- **2. Aprobación.** No empiece sin ella; consulte la issue:
  `gh issue view <número> --json labels --jq '.labels[].name'`
  Hecho: la lista contiene `status:approved`.
- **3. Worktree.** Desde la rama base actualizada, y con el nombre de rama compuesto aquí, nunca escrito a mano (regex `^(feat|fix|refactor|docs|ci|test|chore)/\d+-[a-z0-9]+(-[a-z0-9]+)*$`):
  `git fetch origin && git switch <rama-base> && git pull --ff-only`
  `<rama-nueva>` = `<tipo>/<número>-<slug>`, con el `<slug>` del título de la issue en minúsculas y guiones
  `git worktree add ../<repo>-<número>-<tema> -b <rama-nueva>`
  Hecho: `git worktree list` muestra la ruta nueva.
- **4. Commits.** Por unidad de trabajo: código, tests y docs juntos; Conventional Commits (sección 3); sin atribución de IA.
  Hecho: `git log --oneline <rama-base>..HEAD` muestra una entrada por unidad.
- **5. Integración de la base y preflight local idéntico al CI (HR-4).** Sincronice la rama de feature con la base remota ANTES de publicar: tráigala con `git merge` —nunca rebase, `--amend` ni force-push sobre la rama publicada—, resuelva los conflictos si los hay y ejecute el preflight sobre el resultado integrado:
  `git fetch origin && git merge origin/<rama-base>`
  `bash scripts/local-preflight.sh`
  Hecho: exit 0 sobre el árbol integrado.
- **6. Pull request.** Contra la rama base de la etapa 3; cuerpo con `Refs #N` (o `Closes #N` solo en la punta de cadena), exactamente una etiqueta `type:*` (sección 4) y presupuesto de 400 líneas; si no cabe, cadena:
  Con `chain:partial` o en la punta de cadena (base = rama por defecto) el cuerpo lleva ADEMÁS la sección **Chain Context** (HR-54) como dato: `chain`, `position` N/M, `base`, `depends-on`, `follow-up`, `starts-at`, `ends-with` y `review-budget`, con `base` igual a la base del PR y un único `📍` en el diagrama.
  `gh pr create --base <rama-base> --head <rama-nueva> --title "<tipo>(<ámbito>): <resumen>" --body-file <cuerpo-del-pr.md>`
  Hecho: `gh pr view <número> --json number,url` devuelve número y URL.
- **7. Checks.** Consulte la tabla de checks requeridos (sección 5) y el estado del PR:
  `gh pr checks <número>`
  Hecho: ningún check en fail.
- **8. CI en verde.** Si un check falla, aplique su receta de recuperación (sección 5), reejecute el preflight de la etapa 5 y empuje de nuevo. Si la base avanzó durante la revisión, repita la integración de la etapa 5 (merge, nunca rebase) y vuelva a comprobar: el verde de una cabeza anterior a la integración no es evidencia del resultado nuevo.
  Hecho: todos los checks en pass sobre la cabeza y la base actuales.
- **9. Merge gobernado.** El merge lo decide un humano con CI en verde; ningún agente mergea. Métodos permitidos (host-contract): merge commit, rebase, squash.
  Hecho: `gh pr view <número> --json state --jq .state` devuelve MERGED.
- **10. Limpieza.** Solo después del merge verificado y sin cambios pendientes, retire el worktree local; la rama remota se conserva SIEMPRE (nunca `git push origin --delete`, `--delete-branch` ni borrado de ramas):
  `git worktree remove ../<repo>-<número>-<tema>`
  Hecho: `git worktree list` ya no muestra la ruta y la rama sigue en `origin`.

<!-- ci-pattern-doc:end lifecycle -->

<!-- ci-pattern-doc:begin docs-location -->
## 11. Dónde vive la documentación

- **Toda la documentación de este proyecto vive en el repositorio**, bajo `docs/`, por materia (p. ej. `docs/11-releases.md`, el runbook de release): la leen el CI, el runner y la IA de cualquier miembro del equipo (HR-48).
- **Post-mortems** en `docs/postmortems/<AAAA-MM-DD>-<slug>.md` (parámetro 9): un incidente y sus mitigaciones quedan junto al código que los justifica.
- **Fuera del repositorio, y sin copia en él, solo**: secretos, backups, datos en bruto y material privado. Nada más.
- **Las convenciones compartidas entre proyectos** viven en el catálogo de team-skills y se propagan; la documentación de este proyecto no se propaga.
- Toda ruta de documentación es **relativa al repositorio**: una ruta absoluta o externa la rechaza el gate de documentación operativa de esta fase (HR-48).

<!-- ci-pattern-doc:end docs-location -->

<!-- ci-pattern-doc:begin agents-block -->
## 12. Bloque canónico para AGENTS.md

<!-- ci-pattern-agents:begin -->
## Flujo operativo y mandato de no invención

Bloque gobernado: no lo edite a mano; se regenera desde el patrón.
Antes de actuar, lea la parte del flujo operativo de la situación. Si un paso,
comando, etiqueta o dato que necesita no está en este documento, NO lo invente:
deténgase y pregunte en la issue o al mantenedor.

| Situación | Qué leer antes de actuar |
| --- | --- |
| Crear una issue | Sección 1 (formulario, duplicados, aprobación) |
| Crear una rama | Sección 2 (regex y forma del nombre) |
| Sincronizar la rama de feature con la base | Sección 10, etapa 5 (integración por merge, nunca rebase) |
| Escribir commits | Sección 3 (convención y unidades de trabajo) |
| Abrir un pull request | Sección 4 (cuerpo, etiqueta, presupuesto) |
| Encadenar PRs | Sección 4 (cadena: `chain:partial`, `Refs #N`, punta) |
| Consultar checks | Sección 5 (tabla de checks y recetas de recuperación) |
| Pedir un merge | Sección 6 y la etapa 9 del ciclo de vida |
| Registrar un incidente de producción | Sección 8 (los cuatro niveles: post-mortem, registro operativo, secretos, datos personales) |
| Reportar una fricción | Sección 9 |
| Cualquier paso del ciclo | Sección 10 (ciclo de vida completo) |

**Mandato de no invención:** si un paso, comando, etiqueta o dato que necesita
no está documentado en el flujo operativo, no lo invente: deténgase y
pregunte en la issue o al mantenedor.

## Dónde vive la documentación

Toda la documentación de este proyecto vive en ESTE repositorio, bajo `docs/`
por materia; los post-mortems en `docs/postmortems/<AAAA-MM-DD>-<slug>.md`.
Fuera del repositorio, y sin copia en él, quedan solo los secretos, los
backups, los datos en bruto y el material privado. Nunca ubique documentación
en una ruta absoluta o externa al repositorio (HR-48): el gate de la fase 5 la
rechaza. Detalle: la sección «Dónde vive la documentación» del flujo operativo.
<!-- ci-pattern-agents:end -->

<!-- ci-pattern-doc:end agents-block -->
