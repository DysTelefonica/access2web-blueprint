type: patch

Se borran 8 símbolos de los assets que nadie usaba (`GLOBAL_ENV_USE_RE`, `SHA_EXPRESSIONS`, `SCRAPING_PY`, `GOV_LABELS`, `CAPABILITY_AREAS`, `RECIPES`, `VALID_RULES`, `GateFinding`) y el guard de vulture pasa a escanear también los scripts Python sin extensión (`bin/ci-pattern`), con lo que `SNAPSHOT_CAPABILITY` deja de aparecer como muerto. `DEAD_GRANDFATHERED` baja de 11 a 2 entradas (#663).
