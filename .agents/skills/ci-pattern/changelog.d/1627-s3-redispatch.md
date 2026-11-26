type: minor

Asset nuevo `assets/prerelease/redispatch_nightly.py` (#1627, slice 3): relanza el workflow de la canalización nocturna sobre main con un solo comando. Es idempotente: si ya hay una ejecución nocturna activa para la cabeza de main, sale con SKIP y no relanza. `--dry-run` dice lo que haría y no llama a gh. Un `--repo` mal formado o un estado ilegible se rechaza con 2 sin relanzar. Refs #1627.
