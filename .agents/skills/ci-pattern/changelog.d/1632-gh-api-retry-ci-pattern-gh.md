type: patch

`assets/gh-api-retry/gh-api-retry.sh` toma el `gh` de `$CI_PATTERN_GH` si está fijada (una ruta `.py` corre con `python3`) y, si no, del PATH, como el resto de assets desde #1609. Antes siempre resolvía `gh` por el PATH, así que un `gh` real delante en el PATH se invocaba aunque el test fijara el falso. La suite demuestra que, con `CI_PATTERN_GH` apuntando al falso, el `gh` real no se llama. Refs #1632.
