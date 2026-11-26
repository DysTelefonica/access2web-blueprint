type: patch

**Qué.** El merge gobernado bloqueaba una PR con `required` en verde cuando un run gemelo del mismo SHA había salido `cancelled` por `concurrency` (ardelperal/APAP_WEB#4000, #4001, 2026-11-12): bloqueó los merges de la oleada.
**Causa.** `check_verdict` juzgaba el check-run de `required` con el `id` más alto. El del run cancelado suele ser el más reciente, y su `cancelled` no es un veredicto del SHA, sino el run que perdió la carrera.
**Arreglo.** Por cada nombre decide el check-run más reciente **no cancelado**. Si es verde, pasa; si es rojo, bloquea; si sigue en curso, se espera como `pending`. Un `cancelled` sin gemelo sigue decidiendo y bloquea. La regla está en `check_push_compliance._latest_by_name`, que comparten el merge gobernado y la auditoría de push. Tests: `test_cancelled_twin_of_a_green_run_does_not_block`, `test_cancelled_twin_of_a_running_run_is_waited_for`, `test_cancelled_without_a_twin_still_blocks` y `test_cancelled_twin_does_not_hide_a_red_run`.
**A partir de ahora.** Un `cancelled` solo es veredicto cuando no hay otro check-run del mismo nombre sobre el mismo SHA. El origen del run gemelo en el consumer se sigue en ardelperal/APAP_WEB#4031.
