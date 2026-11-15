type: patch

**Qué.** Cuando falta un check requerido y todos los runs de la cabeza terminaron, el merge gobernado mira el timeline de la PR. Si la base cambió después de esos runs, la pista dice que se cierre y se reabra la PR, o que se empuje un commit, en lugar de «re-run the whole workflow run».
**Causa.** Re-apuntar la base dispara `edited`, que el CI no escucha. La cabeza no tenía ningún run contra la base nueva, y relanzar el run viejo no servía (cadete#1074, #1495).
**Arreglo.** `base_changed_after` lee los eventos `base_ref_changed` del timeline. Sin timeline legible queda la pista de siempre, y el gate bloquea igual.
