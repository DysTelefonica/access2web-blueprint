---
name: governed-delivery
description: "Trigger: actualizar mi rama, abrir un pull request, encadenar PRs, pedir un merge, aplicar etiquetas, partir un cambio, conservar una rama remota. Reglas operativas del día a día en repos con governance."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-10-07
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
| El PR está en verde y hay que integrarlo | 2 — lo lanza un mantenedor, camino gobernado |
| El cambio no cabe en un PR | 3 y 5 — cadena de unidades, punta cierra |
| Añadir o elegir etiquetas | 4 — solo `type:*` del contrato del host |
| El diff pasa de 400 líneas | 5 — partir o `size-exception-reason:`, nunca recortar |
| Después del merge | 6 — la rama remota se conserva |
| Empujé algo que la política prohíbe | 7 — el detective/guard abre el incidente |

Estructura de una cadena, por posición:

```
PR-1 (chain:partial, Refs #N) ──> PR-2 (chain:partial, Refs #N) ──> PR-N (Closes #N)
        📍 actual
```

## Las siete reglas

1. **Actualice con `git merge origin/main`.** Nunca rebase de una rama ya
   publicada, ni `--amend`, ni force-push: la historia publicada es de
   todos y reescribirla rompe los enlaces, los PRs y el trabajo de quien
   venía revisando. — *referencia: `ci-pattern`, HR-54.*

2. **El merge lo lanza un mantenedor, por el camino gobernado**
   (`gh workflow run <workflow> -f pr=<N>`). Ningún agente mergea ni aplica
   `status:approved`: la autoridad de integración es humana y verificable,
   no se delega. Un mantenedor puede aprobar y mergear su propio cambio si
   la protección de la rama lo admite. — *referencia: `ci-pattern`, flujo
   §«Pedir un merge».*

3. **Las cadenas**: los intermedios llevan `chain:partial` y `Refs #N`;
   solo la punta lleva `Closes #N`. El merge avanza de abajo arriba —
   re-apunte del PR siguiente sobre main y empujón, porque re-apuntar por
   sí solo no lanza el CI. — *referencia: `ci-pattern`, HR-7 y HR-54.*

4. **Una sola etiqueta `type:*`, y de las que declara el contrato del
   host.** Una etiqueta que el host no tiene es ruido que ningún gate
   verá jamás: se documenta o se usa la existente, nunca se inventa. —
   *referencia: `ci-pattern`, HR-34.*

5. **400 líneas por PR.** Corte por unidad de trabajo —cada commit, una
   pieza entregable con sus tests y su doc—; `size-exception-reason:` solo
   cuando el corte honesto no cabe, y nunca recortar comentarios,
   documentación ni tests para entrar en el presupuesto. — *referencia:
   `ci-pattern`, HR-8 y HR-54.*

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

## Anti-patrones

- Rebase o `--amend` sobre una rama que ya se empujó: la historia de
  otros deja de cuadrar y el review ya emitido se invalida.
- Force-push a la rama por defecto: el guard lo detecta y abre el
  incidente.
- Declarar un PR listo con el CI en rojo, o encadenar sin re-apuntar y
  empujar el siguiente tramo.
- Dos etiquetas `type:*`, o una etiqueta inventada fuera del contrato
  del host.
- Recortar documentación, comentarios o tests para caber en las 400
  líneas: se parte el cambio, no se adelgaza el diff.
- Que un agente ejecute el merge o aplique `status:approved`.

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
