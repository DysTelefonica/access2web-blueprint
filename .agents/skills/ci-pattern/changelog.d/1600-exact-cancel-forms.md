type: patch

**Qué.** La política de workflows (#978, #1598) aceptaba cualquier `cancel-in-progress` que mencionara `synchronize` o `github.ref`. Una expresión siempre verdadera, como `${{ github.ref == 'refs/heads/x' || true }}` o `${{ !startsWith(github.ref, 'refs/tags/') }}`, pasaba.
**Causa.** La regla buscaba una palabra en el texto en vez de comprobar qué hace la expresión.
**Arreglo.** `_cancels_on_any_event` evalúa la expresión. Entiende un subconjunto del lenguaje de GitHub: `github.event.action`, `github.event_name`, `github.ref`, literales, `!`, `==`, `!=`, `&&`, `||` y paréntesis; ni los nombres de contexto, ni los literales, ni la comparación de cadenas distinguen mayúsculas, como en GitHub. La expresión pasa en cualquiera de estos dos casos:
- es falsa para toda acción que deja el mismo SHA (todas salvo `synchronize`, incluidas `closed`, `enqueued` o `auto_merge_enabled` y una acción desconocida que GitHub añada después) en cada evento de PR que declara el workflow (`pull_request` y `pull_request_target`);
- lee el contexto `github.ref` (no una cadena que lo cite) y es falsa en toda ejecución sobre `refs/heads/<rama base>` (`push`, `workflow_dispatch`, `schedule`); la rama base es la de `--base-branch` del gate (por defecto, `main`).

Todo lo que queda fuera del subconjunto (funciones, otros contextos) es hallazgo. Las formas compuestas que ya usa la flota siguen pasando, porque son seguras: la de access2web-blueprint y la de APAP_WEB en `ci.yml`. Tests: `test_accepted_forms_pass`, `test_forms_that_still_cancel_are_findings` y los de la verificación independiente en `TestCancelInProgress` (`pull_request_target`, rama base, `github.ref` entre comillas, acciones del mismo SHA, mayúsculas).
**A partir de ahora.** `cancel-in-progress` en un workflow con `pull_request` se juzga por lo que la expresión hace, no por las palabras que contiene.
