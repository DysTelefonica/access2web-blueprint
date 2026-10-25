# Evidencia del árbol integrado

Qué mide el merge gobernado después de fusionar, dónde queda registrado y
qué hacer cuando sale en rojo (#1265).

## Qué mide

El merge gobernado ya garantiza que el árbol integrado es el árbol que
revisó el CI: un PR atrasado respecto de su base se bloquea antes de
fusionar (`freshness_findings`, `behind_by > 0` → exit 1) y el merge se
pide con el SHA verificado del head, de modo que si la cabeza cambia entre
la comprobación y el PUT la API responde 409 y el fallo es cerrado.

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
el árbol que revisó el CI.

## Códigos de salida

| Código | Significado |
|---|---|
| 0 | Merge hecho y el árbol integrado es el del head verificado. |
| 1 | El contrato no se cumplió: no se fusiona. |
| 2 | Fallo cerrado: el entorno o la medición no son salvables. |
| 3 | Merge hecho, las issues enlazadas no se cerraron. |
| 4 | Merge hecho, el dispatch post-merge falló. |
| 5 | Merge hecho, **el árbol integrado no es el árbol que revisó el CI**. |

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
