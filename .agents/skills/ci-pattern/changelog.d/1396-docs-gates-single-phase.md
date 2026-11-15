type: minor

`ci-pattern adoption check` acepta `--only-phase N`: mide la fase N y ninguna otra (#1396). El job `docs-gates` de un consumer invocaba `--phase 5`, que repite las fases 0 a 4, y las fases 2 y 3 ya las miden los jobs `env-isolation` y `governance-seals`, así que una misma causa daba dos rojos (medido en `DysTelefonica/cadete`: el re-sellado de #758 lo reportaron a la vez `governance-seals` y `docs-gates`). Con `--only-phase 5` cada job mide su fase una sola vez. `--phase N` no cambia: mide 0..N en orden y se detiene en la primera fase que falla, que es lo que sigue usando el job programado (`--phase 5 .`).

Dentro de la fase 5 se retira la segunda medición de la estructura de `AGENTS.md` (HR-55): la mide solo la fase 3. El porting-guide deja de publicar un paso propio para `check_where_to_work.py` (HR-77), que la fase 5 ya ejecuta con el modo del contrato.

Para el consumer, el cambio consiste en sustituir `--phase 5` por `--only-phase 5` en el job `docs-gates` y retirar el paso explícito de HR-77 si lo tiene. Con el comando anterior el job sigue en verde, solo que mide de más.
