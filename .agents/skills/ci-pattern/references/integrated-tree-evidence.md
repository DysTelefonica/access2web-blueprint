# Evidencia del árbol integrado

Qué mide el merge gobernado después de fusionar, dónde queda registrado y
qué hacer cuando sale en rojo (#1265).

## Qué mide

Con `merge_freshness: "required"` en `.github/host-contract.json`, el merge
gobernado garantiza que el árbol integrado es el árbol que revisó el CI: un
PR atrasado respecto de su base se bloquea antes de fusionar
(`freshness_findings`, `behind_by > 0` → exit 1) y el merge se pide con el
SHA verificado del head, de modo que si la cabeza cambia entre la
comprobación y el PUT la API responde 409 y el fallo es cerrado.

Por defecto (`merge_freshness` ausente u `off`, #1397) el CI corre una vez,
en la PR: un PR atrasado se fusiona y el árbol integrado incluye commits de
la base que el PR no traía. Ese árbol lo verifica el CI que despacha
`post_merge_dispatch` sobre la base, más el nightly; sin `post_merge_dispatch`
el atraso sigue bloqueando, porque nada verificaría el árbol fusionado.
Con `post_merge_verify` (#1440) la verificación se separa de los detectores:
sus workflows se despachan solo cuando la PR iba atrasada, tienen que correr
sus jobs requeridos en `workflow_dispatch`, y una lista no vacía también basta para que
el atraso no bloquee.

Esa garantía era una conclusión, no un registro. Desde #1265, el asset
`assets/runner-controls/check_governed_merge.py` la mide y la publica: tras
el merge compara el **árbol** del commit de merge con el **árbol** del head
verificado y escribe el veredicto en el resumen del job.

Se comparan árboles y no commits a propósito. Con la base contenida en el
head, el árbol resultante es el del head bajo merge, squash y rebase, así
que la comparación no depende del método de merge por defecto del host —que
el host puede cambiar sin avisar— y tampoco del hecho de que el árbol del
commit de merge tenga dos padres.

El árbol se lee de `commit.tree.sha` en
`GET /repos/<owner>/<repo>/commits/<sha>`: la forma real de la API. La
lectura del árbol es un helper compartido con el detective post-push, para
que no haya dos implementaciones de "el árbol de un commit".

## Verificación solo cuando hace falta (`post_merge_verify`, #1440)

La
verificación del árbol fusionado y los detectores son cosas distintas, y
el contrato las declara en listas distintas. Los detectores
(`push-compliance`, `force-push-guard`) van en `post_merge_dispatch` y se
despachan tras cada merge. Los workflows que verifican el árbol integrado
(el `ci` de cadete) van en la clave opcional `post_merge_verify` y se
despachan **solo cuando la PR iba atrasada** respecto a su base: si iba
al día, el árbol integrado es el que ya verificó su CI y no se lanza nada
(el log lo dice con `NOT DISPATCHED: <workflow> (post_merge_verify)`). El
dato es el `behind_by` que el gate ya lee para la frescura. La validación
previa al merge es la de `post_merge_dispatch` (el workflow existe y
declara `workflow_dispatch`) y una más: ningún job **requerido** de un
workflow de verificación puede saltarse en `workflow_dispatch`. Los jobs
requeridos son los de `required_jobs` en `.github/required-jobs-policy.json`
de la base, el mismo fichero que lee `check_required_jobs.py`; si no se
puede leer, el gate falla cerrado con exit 2 nombrándolo. La regla es
conservadora: un `if:` de job que filtra por `github.event_name` tiene que
admitir `workflow_dispatch` (nombrarlo sin negarlo); si no lo admite
(`!= 'workflow_dispatch'`, `== 'pull_request'`, `!(... 'workflow_dispatch')`),
el merge se bloquea con exit 1 nombrando el workflow y el job, porque la
verificación saldría verde sin haber corrido. También se acepta el `if:` en el
que una disyunción de nivel superior es solo un filtro de evento que admite el
dispatch, como `github.event_name != 'pull_request' || (...)` o
`!= 'schedule'`: esa disyunción ya hace verdadero el `if:` en el dispatch
(#1454). Un `&&` o una expresión que el gate no sabe evaluar siguen la regla
conservadora. Un job requerido que de verdad solo corre en otro evento se
declara en `events.workflow_dispatch.accepted_skips` de la misma política,
la lista que ya respeta el agregador, y el gate no lo exige. Una lista mal
formada falla cerrado con exit 2. Un `if:` que no filtra por
evento se acepta, y los jobs que no son requeridos no se miran. Un mismo workflow no puede
estar en las dos listas. Con `merge_freshness: off`, una
`post_merge_verify` no vacía basta para que una PR atrasada se fusione,
igual que `post_merge_dispatch`. El ref, el rango `before`/`after` y el
exit 4 ante un 4xx/5xx del dispatch son los de `post_merge_dispatch`. Un
consumer que no declara la clave sigue como hoy. Migración de cadete:
mover `ci` de `post_merge_dispatch` a `post_merge_verify` y quitar
`github.event_name != 'workflow_dispatch'` del `if:` de sus jobs
requeridos (o hacer que su filtro de eventos admita `workflow_dispatch`).

## Dónde queda

El veredicto se imprime en la salida del job y se escribe en
`$GITHUB_STEP_SUMMARY`, con los dos SHA de commit y los dos SHA de árbol:

```
## Integrated tree evidence

- merge commit: `<sha del commit de merge>`
- merge tree: `<sha del arbol>`
- verified head: `<sha del head verificado>`
- head tree: `<sha del arbol>`
- verdict: the integrated tree is the tree CI reviewed
```

En el caso de drift, la última línea dice que el árbol integrado **no** es
el árbol que revisó el CI. Con la frescura apagada, el drift no es rojo: la
última línea dice `the integrated tree includes base commits the PR did not
carry; the post-merge CI on '<base>' verifies it` y el código es 0.

## Códigos de salida

| Código | Significado |
|---|---|
| 0 | Merge hecho y el árbol integrado es el del head verificado o, con la frescura apagada, incluye commits de la base que verifica el CI post-merge. |
| 1 | El contrato no se cumplió: no se fusiona. |
| 2 | Fallo cerrado: el entorno o la medición no son salvables. |
| 3 | Merge hecho, las issues enlazadas no se cerraron. |
| 4 | Merge hecho, el dispatch post-merge falló. |
| 5 | Merge hecho, **el árbol integrado no es el árbol que revisó el CI** (solo con `merge_freshness: "required"`). |

El 5 es post-merge: el merge no se deshace, porque no se puede. El rojo no
significa "el merge falló" sino "la revisión integrada no tiene la
evidencia funcional del PR". El resumen del job dice cuál de las dos cosas
ocurrió.

Si la comparación no se puede medir —la API falla o el commit no trae la
forma esperada—, el asset falla cerrado con 2 y nombra la causa. Nunca
informa de una coincidencia que no ha medido.

## Qué no hace

- **No re-corre la suite.** La evidencia es la comparación de árboles, no una
  segunda ejecución de PHPUnit ni de Playwright. En un consumer donde las
  suites funcionales solo corren en `pull_request`, el push a la rama
  protegida y el dispatch posterior al merge las saltan: con esta medición
  el árbol integrado queda ligado al árbol probado sin ocupar runner.
- **No publica ningún check en el PR.** El veredicto vive en el job del
  merge, cuyo check-run cuelga del ref del disparo, nunca de la cabeza del
  PR. Nada en el PR aparenta una evaluación distinta de la que se hizo.

## Relación con el detective post-push

`check_tree_integrity` (en `check_push_compliance.py`) también compara
árboles, pero cubre otro caso: audita el rango pushado en una rama
protegida, incluidos los merges que **no** pasaron por el camino gobernado.
Este control confirma el resultado del merge que el propio asset acaba de
hacer, sobre el head que el propio asset acaba de verificar.

Tres diferencias que conviene no confundir:

- El detective **falla abierto**: unos datos ausentes o de forma inesperada
  no bloquean. Este control **falla cerrado**.
- El detective **no aplica en un rebase-merge**, porque deduce la base del
  primer padre del commit de merge en lugar de comparar los dos árboles
  directamente. Este control sí lo cubre.
- El detective **no emite veredicto**: su rojo es una incidencia con
  evidencia. Este control publica el veredicto con los cuatro SHA.

Los dos miden, cada uno desde donde puede. No son dos detectores de la
misma causa: el detective mira el rango pushado, este mira su propio merge.

## Qué hacer cuando sale en rojo

1. Leer el resumen del job: los cuatro SHA y el veredicto.
2. Si los árboles difieren, comparar el commit de merge con el head
   verificado (`git diff <head> <merge>`) y ver qué se coló.
3. Comprobar si la base se movió entre la comprobación y el merge: es la
   causa esperable cuando alguien fusionó otra cosa a la vez.
4. La revisión integrada **no** puede darse por probada por el PR. Corregir
   y volver a fusionar; el rojo no se deshace con un rerun.
