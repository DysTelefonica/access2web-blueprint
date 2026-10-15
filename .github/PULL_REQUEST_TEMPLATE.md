<!--
El presupuesto de revisión es 400 líneas (additions + deletions) y se aplica en
el job `review-budget` del workflow de CI (`scripts/check_pr_size.py`).
La única manera de sobrepasarlo es partir el PR o aplicar la etiqueta
`size:exception` — la concede el mantenedor y se documenta en el cuerpo del
PR. Con la adopción del gobierno simplificado (#854) se retiró el campo
`size-exception-reason:` del cuerpo: ya no es gate.
-->

Closes #

## Chain Context

<!-- Obligatoria en la punta de una cadena y en todo PR con `chain:partial`
     (HR-54 del patrón): los ocho campos son dato y el diagrama de abajo lleva
     el marcador del PR actual una sola vez. Un PR suelto declara
     `position: 1/1` y `depends-on: none`. -->

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
