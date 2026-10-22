# Recuperación del backlog no fusionado (#794, #785, Lanzadera, propagate)

## Objetivo y alcance

El usuario pidió fusionar todo lo que no está en `main` si el CI está verde. La
auditoría previa mostró que las cuatro unidades no están en estado mergeable y
que el merge está reservado al mantenedor por la hard rule de `AGENTS.md`
(línea 204: «Ningún agente mergea ni aplica `status:approved`»). El usuario
autorizó entonces el trabajo preparatorio: sincronizar #823, arreglar la cadena
B, y preparar las dos ramas sin PR.

Alcance: dejar cada unidad realmente mergeable y verificada, sin mergear, sin
aplicar `status:approved` y sin aprobar PRs. El merge lo pulsa el mantenedor.

## Hallazgos y fuentes

Auditoría sobre `origin/main` (tras `git fetch --prune`), no sobre `HEAD` local:
`main` local estaba 32 commits atrás y filtrar contra él produce falsos
positivos.

Cadena A (issue #794):

- #823 `fix/794-fastapi-starlette` → `main`, estado `BEHIND`. `required` en
  success pero contra una base ya avanzada: verde obsoleto. 2 commits.
- #824 `fix/794-pytest9` → `fix/794-fastapi-starlette`, `CLEAN` y verde.

Cadena B (issue #785):

- #799 `fix/785-mutation-cosmic-ray` → `main`, `DIRTY` (conflictos).
- #816 `ci/785-mutation-matriz` → `fix/785-mutation-cosmic-ray`, `UNSTABLE`:
  `required` fail y `review-budget` fail.
- #825 `ci/785-mutation-linea-base` → `ci/785-mutation-matriz`, `CLEAN` y verde.

El orden de integración es forzoso y no conmutable:

```
#823 → main     #824 → #823
#799 → main     #816 → #799     #825 → #816
```

Sin PR:

- `docs/lanzadera-nueva-resumen-2026-09-29`: 1 commit,
  `lanzadera-nueva-resumen-2026-09-29.html` +440 líneas, publicada en `origin`.
- `chore/propagate-personal-skills-at-7eca6b8`: 1 commit, 265 archivos y
  +37.715 líneas, 94× el presupuesto de 400. No está en `origin`.

Dependabot: 11 PRs abiertos hacia `main`, todos `mergeable=UNKNOWN`, con
colisión interna (#768 y #769 son ambos `codeql-action-4.38.2`). Fuera del
alcance autorizado.

## Tareas

- [x] T1 — Sincronizar `fix/794-fastapi-starlette` con `origin/main` por merge
      (nunca rebase), publicar el merge y observar el CI nuevo. Hecho: #823 pasó
      de `BEHIND` a `MERGEABLE` con `required` en verde contra `b5ac6ba`.
- [x] T2 — Diagnosticar los conflictos de #799 contra `main` y registrar el
      delta real antes de resolver nada. Hecho: conflictos en
      `.github/workflows/ci.yml` y `docs/10-runners.md`, resueltos a favor de
      `main` porque #800 (`34cbcb3`) ya había hecho el mismo cambio de runner con
      más alcance. #799 quedó `MERGEABLE` con el CI completo en verde contra
      `e55bae6`.
- [x] T3 — Reparar #816 hasta `required` y `review-budget` en verde. Hecho por
      corte, con autorización explícita: el gate medía 443 contra 400 y no había
      prefijo contiguo en la historia (el commit inicial `c4aa844` mezclaba las
      dos unidades). #816 se cerró con su rama remota conservada y lo reemplazan
      #843 (82 líneas) y #844 (376 líneas); #825 se retargeteó al tramo 3/4 y
      quedó con su delta propio de 119 líneas. Los tres en verde.
- [ ] T4 — Issue y PR para el resumen de Lanzadera. El archivo son 440 líneas en
      un solo HTML, así que supera el techo de 400 y necesita `size:exception`
      con su motivo, o una decisión distinta del mantenedor.
- [ ] T5 — Analizar el troceado de `chore/propagate-personal-skills-at-7eca6b8`
      y proponer la división antes de crear PRs.

## Restricciones

- Un único escritor activo. Los worktrees se crean con el layout canónico de
  `worktree-reorg-per-project` v2.0: contenedor hermano, nombre plano.
- Sin merge a `main`, sin `status:approved`, sin aprobar PRs, sin borrar ramas
  remotas. Publicación no forzada únicamente sobre las ramas autorizadas.
- Sin rebase y sin `--amend` sobre ramas publicadas: `docs/ci-pattern-flow.md`
  etapa 5 lo prohíbe de forma explícita. La integración con la base es por merge.
- No ocultar un rojo cambiando políticas, runners, workflows o scripts. Si un
  check falla por su causa real, se corrige la causa o se reporta el bloqueo.
- T5 no se ejecuta como 95 PRs de 400 líneas: el número es absurdo y hay que
  elevarlo como decisión antes de publicar.

## Comprobaciones

T1: `git merge` sin conflictos o conflicto reportado; el PR deja de estar
`BEHIND`; `required` en verde contra el SHA nuevo. Un verde anterior no cuenta.

T3: `gh pr checks` de #843, #844 y #825 con `required` y `review-budget` en
success contra el head nuevo; sin `skipping` no justificado.

T4: `scripts/check_pr_size.py` sobre la rama y el PR enlazado a su issue.

T5: conteo real de archivos y líneas por unidad propuesta, con el presupuesto
de 400 líneas de `CONTRIBUTING.md` como techo.

## Estado operativo

Auditoría completada: 17 worktrees ya fusionados eliminados (2,1 G → 651 M) y
`odd/tasks/a2w-phase5-operating-doc.md` preservado desde un worktree de fase 5.

Autorización explícita del usuario en esta sesión: sincronizar #823 con
`origin/main` y re-disparar CI; resolver los conflictos de #799 y reparar #816;
partir #816 en dos PRs encadenados aceptando cerrar el PR original porque el
force-push sobre una rama publicada está prohibido; preparar issue y PR de la
rama de Lanzadera; y trocear la rama de propagación de skills personales. El
merge sigue siendo del mantenedor.

### Cadena de #785 renumerada a cuatro tramos

```
#799 (1/4) fix(ci): cosmic-ray y config al dominio        base main
#843 (2/4) ci(mutation): gate multi-sesión, 82 líneas     base #799
#844 (3/4) ci(mutation): matriz + ratchet, 376 líneas     base #843
#825 (4/4) test(mutation): línea base, 119 líneas         base #844
#816       cerrado; rama ci/785-mutation-matriz conservada
```

Los cuatro PRs abiertos tienen `required`, `quality`, `codeql`, `review-budget`
y los tres `security` en verde contra su cabeza actual.

Defecto destapado por el corte: los commits de la matriz se escribieron contra
la base anterior a #800, así que reemplazaban sus comentarios en vez de
convivir con ellos —la cabecera del job `mutation` perdía la explicación de la
etiqueta con cero runners registrados—. Reconciliado en `6b58da6` dentro de
#844, y #825 hereda la reconciliación por el merge de su base. Eso mismo habría
ocurrido al sincronizar #816.

Pendiente heredado, fuera de alcance: el comentario del paso `disable
sparse-checkout cone mode` del job `quality` sigue diciendo «on the self-hosted
a2w runner», que dejó de ser cierto con #800.

## Siguiente paso

T4: la issue #845 está creada y el PR queda bloqueado hasta que un mantenedor
aplique `status:approved`: el agente no aplica esa etiqueta.

T5: el corte está definido y respaldado por el precedente del repositorio, pero
su publicación está bloqueada por un hallazgo de duplicación de autoridad que
decide el mantenedor.

### T4 — resumen de Lanzadera

Issue #845 con las seis secciones canónicas del formulario, creada desde
`issue-canonical.yml` (que declara `labels: []`, así que nace sin etiquetas) y
con read-back verificado. El PR no se abre sin `status:approved`.

Hallazgo que la issue eleva: `lanzadera-nueva-resumen-2026-09-29.html` no sigue
el patrón `estado-planificacion-*.html` de los cuatro HTML que ya viven en la
raíz de `main`. O es el tablero vivo del plan y se renombra, o es un resumen de
una sola vez y la decisión que lo exceptúa queda escrita. Son 440 líneas en un
solo documento, así que el PR necesita `size:exception` con su motivo.

### T5 — propagación de personal-skills

Corte propuesto, con el precedente de #830 y #840:

```
tramo 1/2  AGENTS.md (42) + .team-skills.yaml (55) = 97 líneas, revisable
tramo 2/2  .agents/skills/**  263 archivos / 37.623 líneas, atómico
           con size:exception: «importe vendorizado del catálogo, atómico;
           la parte revisable vive en el tramo 1»
```

La unión de los dos tramos reproduce el árbol de la rama exactamente: su diff
contra `main` son 264 adiciones y 1 modificación, y el corte es por ruta. La
rama no está publicada, así que no hay force-push ni reescritura de nada
publicado: los dos tramos son ramas nuevas.

**Hallazgo que bloquea la publicación.** De los 263 archivos bajo
`.agents/skills/`: 233 son byte a byte idénticos a su gemelo en `skills/`, 29
divergen y 1 no tiene gemelo. `main` ya arrastra 60 archivos duplicados entre los
dos árboles, así que la propagación llevaría la cifra a unos 293 e introduciría
**29 skills con dos versiones distintas en el mismo repositorio**. Ninguna de las
43 skills que se añaden es nueva: las 43 ya existen en `skills/`.

Eso es autoridad duplicada: el `AGENTS.md` de `main` apunta la tabla de skills a
`skills/…`, de modo que un agente cargaría la copia vieja mientras el catálogo
ya materializó la nueva. Se decide antes de publicar: retirar la copia legacy en
el tramo 2/2 (como declara #840 en su cuerpo) o aceptar explícitamente que los
dos árboles conviven.

Segundo punto: los siete PRs de propagación mergeados (#755, #761,
#767, #770, #830, #838–#840) **no referencian ninguna issue**, lo que
contradice la regla issue-first.

**Decisión del usuario**: retirar la copia legacy en el tramo 2/2, e issue
paraguas primero, con los tramos después de la aprobación.

**Tercer hallazgo, y es el más grave.** El merge de la rama contra `main`
**conflictúa** en `AGENTS.md` y `.team-skills.yaml`
(`git merge-tree --write-tree --name-only`). El merge-base es `169820e` (#767) y
desde ahí `main` reescribió el tramo superior del `AGENTS.md`
(`@@ -1,200 +1,34 @@` en coordenadas del merge-base): la diferencia real
`main` → rama es de 194 inserciones y 111 borrados. Los marcadores del slice
difieren (`@ v1844dd2` en `main`, `@ veaf74c8` en la rama) y su dirección no
está establecida. Publicar la rama sin reconciliar **re-inyecta el slice viejo
sobre la gobernanza que `main` ya adoptó** — el mismo tipo de defecto que en la
cadena de #785 borraba el rationale de #800 en silencio.

La retirada, además, rompe 46 referencias del repositorio a
`skills/<nombre>/SKILL.md`, 77 de ellas a las 43 skills que se retirarían.

Issue paraguas creada: **#846** (`[XCUT] chore(fleet): vía canónica para las
propagaciones de flota y fin de la autoridad duplicada en skills/`), con las seis
secciones canónicas y read-back verificado. Los dos tramos no se abren sin
`status:approved`.

### T5 — ejecutada (defectos filados + retirada por PR)

Instrucción del propietario: sólo evidencia y especificación en los defectos
(**modo auditoría, sin tocar el propagador**), y retirar por PR las 3 skills
duplicadas. La propagación real queda pendiente de confirmación.

**Issues en `DysTelefonica/team-skills`:**

| Issue | Contenido | Estado |
|---|---|---|
| **#421** | `-DryRun` no es de sólo lectura: el bloque *Step 2: git operations* (504 `fetch`, 514/521 `checkout`, 546 `reset --hard`, 554 `merge --ff-only`) corre sin ningún `if (-not $DryRun)` — cero menciones de `DryRun` entre 498 y 560 | creada, etiqueta `friction:recurrent`, enlazada desde **#377** (comentario verificado: `issue_url` → issue 377) |
| **#422** | La cabecera del paso 4 (35–38) promete retirar toda copia legacy; el código sólo retira las skills que **salieron** del catálogo (684 → 690 → 731 → `foreach` 732) | creada, `triage_class` C |

Ambas con read-back verificado (título y cuerpo exactos). Ninguna toca código.

**Comparación previa a la retirada (3 vías: `skills/` vs `.agents/skills/` vs
catálogo clonado completo, 920 commits):**

| Skill | ficheros | resultado |
|---|---:|---|
| `repository-delivery-governance` | 7 | idéntica a su gemela y al catálogo |
| `worktree-reorg-per-project` | 1 | idéntica a su gemela y al catálogo |
| `deterministic-quality-harness` | 52 | contenido idéntico; **24 diferencias = sólo CRLF** (`diff --strip-trailing-cr` → 0) |

Ninguna tenía cambios propios → **no se abre issue de subida al catálogo**.
Cuidado registrado: el test por SHA de blob dio 24/24 «únicos» por la
normalización de finales de línea; es un falso positivo que sólo `--strip-trailing-cr`
desmonta.

**PR de retirada:** `access2web-blueprint#847`, rama
`chore/846-retirar-skills-duplicadas`, `Refs #846`, `type:chore`, **0/-7620**
en 60 ficheros. Preflight local exit 0, `git diff --check` limpio, gate de
tamaño `OK … overridden`. Rama pasó `check_branch_name.py`. En verde:
`required`, `quality`, `review-budget`, `codeql` y los tres `security`.

Dos hallazgos operativos dejados por escrito:

- El gate de presupuesto se cumple con el **cuerpo** (`size:exception` +
  `size-exception-reason:`), no con la etiqueta: `check_pr_size.py` no mira
  etiquetas. Me faltó el marcador en el primer intento (`FAIL 7620`).
- `pull_request.body` vive en el payload del evento: el workflow no escucha
  `edited` y `gh run rerun` reutiliza el payload original (verificado: relbió
  el cuerpo viejo). Sólo un `synchronize` —commit `ea7bfbd` con 0 ficheros—
  relee el cuerpo corregido.

No se tocó `.skills-fleet-manifest.json` (artefacto del propagador anterior;
el test de contrato del catálogo exige que el actual no lo escriba y ningún
gete del consumer lo lee): fuera de alcance, documentado en el PR.

### Estado final de las tareas

| Tarea | Estado |
|---|---|
| T1 #823 | hecho y verificado |
| T2 #799 | hecho y verificado |
| T3 #816 partido | hecho y verificado; **cadena entera mergeada** (`10b6b00`, `93e8330`, `6504640`, `7bb55fd`, `daa4726`) |
| T4 resumen de Lanzadera | issue #845 con `status:approved`; PR aún no abierto |
| T5 defectos del propagador | **#421 y #422 en team-skills**, modo auditoría |
| T5 retirada de las 3 skills | **PR #847 abierto y en verde** (falta la etiqueta protegida `size:exception`) |
| T5 propagación real | **no ejecutada** — espera confirmación explícita |
