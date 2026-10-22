<!--
El presupuesto de revisión es 400 líneas (additions + deletions) y se aplica en
el job `pr-gate` del workflow de CI
(`.agents/skills/ci-pattern/assets/pr-size/check_pr_size.py`). Las rutas
vendorizadas `.agents/skills/**` y `skills/**` no cuentan (issue #918). El gate
lee solo la etiqueta `size:exception`, que pone el mantenedor; el cuerpo no
lleva ningún campo de excepción.
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

## Excepción de tamaño

<!-- Solo si el PR supera 400 líneas. Responda las cuatro preguntas; el mantenedor
     decide si concede la etiqueta `size:exception`. -->

1. ¿Qué parte puede ir antes en una PR propia sin romper `main` (refactor previo, puerto o servicio con sus tests, docs)?
2. ¿Se puede partir por sub-comportamiento o por capa (dominio → ruta → plantilla + e2e)?
3. ¿Hay cambios que no se pidieron (formato, renombrados, limpieza)? Quítalos antes de medir.
4. Si se fusiona solo la primera mitad, ¿qué se rompe exactamente en `main`? Nombra el test o el fichero concreto. Si no puedes nombrarlo, la PR se parte.

## Checklist

- [ ] La issue canónica vinculada está linkada y tiene `status:approved`.
- [ ] CI verde contra la base actual: `required`, `quality`, `codeql`, `pr-gate` y los tres de `security`.
- [ ] Conventional Commits, sin atribución de IA.
- [ ] La rama remota se preserva tras el merge (nunca `git push origin --delete <rama>`).
- [ ] El merge lo lanza un mantenedor por el camino del host (ruleset `main-maintainers-and-admins-merge`). Ningún agente aplica `status:approved` ni mergea.
