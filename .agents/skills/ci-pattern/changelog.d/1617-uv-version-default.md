type: patch

El valor por defecto del parámetro `P31 uv_version` de `assets/parameters.schema.json` pasa de `0.9.28` a `0.13.4`, la versión a la que se mueve APAP_WEB (APAP_WEB#4008). Así la propagación ya no sugiere la versión antigua. Un consumidor que fije su propia versión no cambia (#1617).
