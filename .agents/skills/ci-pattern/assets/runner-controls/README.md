# runner-controls — assets que el runner ejecuta sobre el evento

Assets del patrón que corren en el runner (HR-34) con el token del job: leen el
evento y el contrato del host y dejan un veredicto. Solo stdlib; el acceso a la
API vive en `github_rest.py` (reintentos, timeout y `User-Agent` propio).

## Contenido

| Fichero | Papel |
|---|---|
| `check_force_push_guard.py` | Guard de force-push y de borrado de una rama protegida por contrato |
| `check_push_compliance.py` | Detective post-push: lo que llega a la rama protegida debe ser un PR mergeado con sus checks en verde |
| `check_governed_merge.py` | Gate pre-merge: el PR entra por el camino gobernado y el merge sale auditado |
| `check_runner_binding.py` | Enlace entre el runner y el contrato del host |
| `github_rest.py` | Cliente de la API de GitHub que usan los tres assets de arriba |
| `safe_merge.py`, `safe_cleanup.py` | Fusión con permisos acotados y limpieza con salvaguarda de Leap |
| `tests/` | Suite ejecutable (solo stdlib, sin red) |

## El guard de force-push: qué mide y con qué evento (#1266)

El guard se dispara con dos formas de evento y cada una lleva el rango y las
banderas en un sitio distinto:

| Evento | Rango | Banderas `forced`/`deleted` | Veredicto |
|---|---|---|---|
| `push` | en el nivel superior (`before`, `after`) | las trae | `forced: true` o `deleted: true` sobre una rama protegida abre incidencia y sale 1; un push ordinario sale 0 |
| `workflow_dispatch` (el que despacha el merge gobernado, #817) | en `inputs` (`inputs.before`, `inputs.after`) | no las trae: no existen en ese payload | el rango bien formado sale 0 y dice qué midió; el rango mal formado falla cerrado con 2, nombrando la causa |

El rango se lee con `push_range`, el lector de `check_push_compliance.py` que ya
usa el detective post-push (#933): hay una sola implementación de «el rango de
este evento» y el guard no trae la suya. Un dispatch no trae las banderas, y su
ausencia se declara en la salida en vez de callar un veredicto que no se midió:

```
OK: dispatch on protected branch 'main' audited the range <before> -> <after>: well-formed range, no force-push nor deletion reported for it; the event carries no 'forced'/'deleted' flag
```

El guard no busca reescrituras dentro del rango despachado: las reescrituras
llegan como evento `push`, que es el camino que ya audita.

Salidas del guard: 0 veredicto conforme · 1 incidencia abierta (force-push o
borrado de una rama protegida) · 2 fallo cerrado: evento malformado o
incidencia que no se pudo abrir (HR-3, nunca verde sin medición).

## Destino en el consumer

Copie el directorio a `scripts/runner-controls/` del repo consumer (o la ruta
equivalente) y versione allí su propio contrato. Si el consumer también usa
`required-jobs/`, copie `runner-controls/` junto a él: `build_annotations.py`
carga `github_rest.py` por ruta relativa (`../runner-controls/github_rest.py`).
El workflow del consumer declara los `inputs` `before` y `after` bajo
`on.workflow_dispatch.inputs` si quiere que el guard de force-push mida el
rango del merge gobernado; declarar solo uno de los dos bloquea el merge.
