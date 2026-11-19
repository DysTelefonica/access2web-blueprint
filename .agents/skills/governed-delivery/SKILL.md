---
name: governed-delivery
description: "Trigger: actualizar mi rama, abrir un pull request, encadenar PRs, pedir un merge, aplicar etiquetas, partir un cambio, conservar una rama remota. Reglas operativas del día a día en repos con governance."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.7.0"
  last_verified: 2026-11-19
  scope: ['universal', 'ops']
  auto_invoke: ['actualizar rama', 'abrir un PR', 'encadenar PRs', 'pedir merge', 'etiquetar un PR']
  tiers: ['universal', 'ops']
---

# Entrega gobernada — día a día

Skills pequeñas de trabajo corriente: qué hacer, y por qué en una línea.
Para ADOPTAR o AUDITAR el patrón (gates, contratos, HRs, fases) cargue
`ci-pattern`; ésta no lo sustituye, la precede.

## Mapa rápido

| Situación | Regla |
| --- | --- |
| Mi rama está desfasada de main | 1 — `git merge origin/main` |
| El PR está en verde y hay que integrarlo | 2 — lo lanza el agente con el perfil de el administrador que lleva la sesión, camino gobernado y CI en verde |
| El cambio no cabe en un PR | 3 y 5 — cadena de unidades, punta cierra |
| Añadir o elegir etiquetas | 4 — `type:*` orientativa, del contrato del host |
| El diff pasa de 400 líneas | 5 — partir; la etiqueta `size:exception` la decide el administrador que lleva la sesión, nunca recortar |
| Después del merge | 6 — la rama remota se conserva |
| Empujé algo que la política prohíbe | 7 — el detective/guard abre el incidente |
| El consumer decide la estrategia de merge | 8 — `--squash` o `--no-ff`, ambos válidos; la rama remota sobrevive |

Estructura de una cadena, por posición:

```
PR-1 (chain:partial, Refs #N) ──> PR-2 (chain:partial, Refs #N) ──> PR-N (Closes #N)
        📍 actual
```

## Las nueve reglas

Antes de `gh pr create`, ejecute `ci-pattern pr preflight --draft --pr-body <cuerpo.md> --base <rama-base> --label <etiqueta>`: comprueba cuerpo, etiquetas y tamaño sin API.

1. **Actualice con `git merge origin/main`.** Nunca rebase de una rama ya
   publicada, ni `--amend`, ni force-push: la historia publicada es de
   todos y reescribirla rompe los enlaces, los PRs y el trabajo de quien
   venía revisando. — *referencia: `ci-pattern`, HR-54.*

   Las historias se unen con un `git merge origin/<base>` normal,
   resolviendo los conflictos. En una rama que va a un PR están prohibidos
   `git merge -s ours` y `-X ours`: `-s ours` registra la base como
   fusionada pero conserva el árbol de la rama y descarta los cambios de la
   base, y el merge gobernado puede revertir después en silencio el trabajo
   de la base (incidente: ardelperal/APAP_WEB, 2026-11).

2. **El merge lo lanza el agente que lleva la issue, con el perfil de
   el administrador que lleva la sesión, por el camino gobernado**
   (`gh workflow run <workflow> -f pr=<N>`), cuando la CI está en verde. No
   hace falta una orden aparte: una orden permanente registrada en el
   `AGENTS.md` local del repositorio cuenta como orden registrada.
   `status:approved` no es puerta: la puede poner cualquiera, agentes
   incluidos (decisión del dueño del 2026-10-23, 17:24, en los comentarios
   de #1093). Antes de actuar, el agente comprueba que la cuenta de `gh` en
   uso administra el repositorio:
   `gh api repos/{owner}/{repo} --jq '.permissions.maintain or .permissions.admin'`.

   Tras la acción, el agente relee el estado del merge con un solo intento;
   si el resultado es incierto, se detiene sin reintentar. El merge exige CI
   en verde y el camino gobernado; nunca `--admin`. El verde se comprueba con
   `python .agents/skills/ci-pattern/assets/runner-controls/wait_ci.py
   --repo <owner/repo> --pr <n> --check <agregador>` (p. ej. `required`): espera a que los checks
   aparezcan y terminen y nunca da verde con checks ausentes o pendientes
   (`gh pr checks --watch` sale con 1 justo después del push, #1502). —
   *referencia: `ci-pattern`, flujo §«Pedir un merge».*

3. **Las cadenas**: los intermedios llevan `chain:partial` y `Refs #N`;
   solo la punta lleva `Closes #N`. El merge avanza de abajo arriba: se
   re-apunta el PR siguiente sobre main. Re-apuntar no lanza la CI (dispara
   `edited`, que el CI no escucha). Bajo el perfil pre-release (#1627,
   regla 5), un re-apuntado con el head sin cambiar no necesita run: el
   merge gobernado da por bueno el run verde que ya existe sobre ese SHA.
   Cuando el eslabón sí necesite un run, la única vía es un commit vacío
   (`git commit --allow-empty`); nunca fusionando main en la rama
   (§«Disciplina local», push). Cerrar y reabrir el PR no es la vía
   de relanzar. Un run cancelado o caído por un transitorio lo relanza el
   agente que lleva la issue, con el perfil de el administrador que lleva la sesión, con `rerun_run.py`. —
   *referencia: `ci-pattern`, HR-7 y HR-54.*

4. **Etiqueta `type:*` orientativa, de las que declara el contrato del
   host.** Ningún gate la exige (team-skills#1093, R1). Una etiqueta que el
   host no tiene es ruido que ningún gate verá jamás: se documenta o se usa
   la existente, nunca se inventa. — *referencia: `ci-pattern`, HR-34.*

5. **400 líneas por PR.** Corte por unidad de trabajo —cada commit, una
   pieza entregable con sus tests y su doc—. La etiqueta `size:exception` la
   decide el administrador que lleva la sesión, solo cuando el corte honesto no cabe; nunca se recortan
   comentarios, documentación ni tests para entrar en el presupuesto. —
   *referencia: `ci-pattern`, HR-8 y HR-54.*

   El presupuesto se planifica **antes** de empezar, no se descubre al final
   (regla del dueño, 2026-11-02). Quien encarga el trabajo estima su tamaño y
   lo parte en eslabones de una cadena ya en el encargo. Quien lo hace mira el
   tamaño de su diff mientras trabaja y corta el eslabón hacia las 350 líneas.
   El chequeo de tamaño del preflight local es la última red, y la CI nunca
   debería ser la primera en descubrir que una PR se pasa.

   Si un eslabón no se puede partir, el administrador que lleva la sesión decide antes del push si
   de verdad no se puede (un fichero generado, una migración atómica o un
   renombrado mecánico) o si falta el esfuerzo de partirlo, y lo registra con
   un comentario en la PR. El único mecanismo de la flota es la etiqueta
   `size:exception` que lee el gate `pr-size` (`ci-pattern` HR-8), con las
   cuatro respuestas en el cuerpo de la PR.

6. **La rama remota se conserva tras el merge** (nunca `git push origin
   --delete <rama>`): permite seguir o revertir con exactitud algo que se
   hizo, y su aparente ruido en la lista de ramas no justifica perder esa
   trazabilidad. — *referencia: `ci-pattern`, flujo §«Worktree y rama
   remota».*

7. **El detective post-push y el guard de force-push abren un incidente**
   cuando un empujón viola la política —push a la rama protegida o
   reescritura de historia—: se repara con post-mortem y registro de la
   causa, nunca borrando el rastro. — *referencia: `ci-pattern`, HR-49
   (detective) y HR-51 (guard).*

8. **La estrategia de merge la decide el consumer.** `--squash` y `--no-ff`
   son ambas válidas; la invariante del catálogo es que la rama remota
   sobreviva al merge, no la forma del commit en la base activa
   (Refs DysTelefonica/team-skills#528 §3.13 D1; `slices/partials/web.md`
   §Divergencias documentadas). Si el partial del consumer prescribe una
   concreta, gana el partial; en ausencia de prescripción, el agente
   respeta el método configurado en el host o pregunta a el administrador que lleva la sesión.

9. **Antes del merge, quien lleva la issue comprueba la conformidad con la
   issue, no la CI.** La PR hace lo que pide la issue (alcance y criterios de
   aceptación), y cada criterio tiene su evidencia en la PR, una línea por
   criterio: test con su salida en verde, comando con su salida real, captura
   para la UI, consulta o log para los datos. Un criterio sin evidencia no
   está hecho. Que la CI pase es parte de hacer la PR (`ci-pattern` HR-86).

## Disciplina local

Reglas sin actor: valen igual con roles o sin ellos.

- **Scratch.** El trabajo temporal va en `d=$(mktemp -d)` y se queda allí; si hay que
  rehacerlo, otro `mktemp -d`. Al terminar se listan las rutas de scratch y de
  worktrees creadas, para decidir qué limpiar.
- **Push.** Cada push encola la CI entera de la rama. Se juntan los commits de una
  tanda y se empuja una vez; una rama que aún no se fusiona no se pone al día; con
  `merge_freshness` en `off` no se fusiona `main` en la rama salvo conflicto (`DIRTY`)
  o el mismo fichero de `.github/workflows/` cambiado en las dos (#783); en una
  cadena se empuja solo el eslabón siguiente; si un push deja obsoleta la ejecución
  anterior de la propia rama, se cancela, nunca la de una rama ajena.
- **Preflight.** La CI confirma, no descubre (`ci-pattern` HR-86): antes de cada push
  pasa el preflight local del repositorio, con tamaño, contrato de PR y commits,
  nombre de rama, lint, tipos, tests unitarios y secretos. Una PR en rojo en uno de
  esos gates es un defecto de proceso.

## Anti-patrones

- Rebase o `--amend` sobre una rama que ya se empujó: la historia de
  otros deja de cuadrar y el review ya emitido se invalida.
- `git merge -s ours` o `-X ours` en una rama que va a un PR: la base
  figura como fusionada y sus cambios se pierden en el merge.
- Force-push a la rama por defecto: el guard lo detecta y abre el
  incidente.
- Declarar un PR listo con el CI en rojo, o encadenar sin re-apuntar y
  empujar el siguiente tramo.
- Una etiqueta inventada fuera del contrato del host.
- Recortar documentación, comentarios o tests para caber en las 400
  líneas: se parte el cambio, no se adelgaza el diff.
- Que un agente ejecute el merge con la CI en rojo, saltándose el camino
  gobernado o con una cuenta que no administra el repositorio.

## Qué cargar junto a esto

| Tarea | Skill |
| --- | --- |
| Cualquier cambio (obligatorio) | `ai-slop-discipline` |
| Adoptar o auditar el patrón, HRs, gates, fases | `ci-pattern` |
| Clasificar issues, evidencia de PR, políticas de rama | `repository-delivery-governance` |
| Preflight local y ratchets de calidad | `deterministic-quality-harness` |

## Antes de actuar

Antes de cualquier cambio cargue `ai-slop-discipline` y pase su
comprobación; ésta es la capa de «cómo se entrega», aquella es la capa de
«qué se puede escribir».
