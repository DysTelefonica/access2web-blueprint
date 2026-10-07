<!--
`size-exception-reason:` es un campo de datos del cuerpo del PR: una sola
línea, en la primera columna, exactamente una aparición — nunca etiqueta ni
prosa dispersa (HR-8 del patrón). Deje el motivo vacío salvo que el diff supere
las 400 líneas y no quepa partirlo; el gate de tamaño del repositorio
(`scripts/check_pr_size.py`) pide además la etiqueta `size:exception`.
-->

size-exception-reason: <motivo en una sola línea, solo si el diff no puede partirse>

Closes #

## Chain Context

<!-- Obligatoria en la punta de una cadena y en todo PR con `chain:partial`
     (HR-54 del patrón): los ocho campos son dato y el diagrama lleva
     exactamente un 📍. Un PR suelto declara `position: 1/1` y
     `depends-on: none`. -->

- chain:
- position:
- base:
- depends-on: none
- follow-up: none
- starts-at:
- ends-with:
- review-budget:

```
📍 este PR
```

## Tests que prueban el cierre

<!-- Obligatoria si el cuerpo lleva `Closes #N` (HR-53 del patrón): nombre el
     fichero o el caso de test que demuestra el comportamiento. La prosa no
     cierra nada. -->

## Checklist

- [ ] La issue canónica vinculada está linkada y tiene `status:approved`.
- [ ] CI verde contra la base actual: `required`, `quality`, `codeql`, `review-budget` y los tres de `security`.
- [ ] Conventional Commits, sin atribución de IA.
- [ ] La rama remota se preserva tras el merge (nunca `git push origin --delete <rama>`).
- [ ] El merge lo lanza un mantenedor por el camino del host (ruleset `main-maintainers-and-admins-merge`). Ningún agente aplica `status:approved` ni mergea.
