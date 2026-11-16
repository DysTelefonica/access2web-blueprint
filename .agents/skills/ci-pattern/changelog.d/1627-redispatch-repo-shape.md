type: patch

`redispatch_nightly.py` valida `--repo` con `fullmatch`: rechaza un salto de línea final (`o/r\n`) y los segmentos solo de puntos (`../..`, `owner/.`). Un estado `action_required` de la nocturna activa también salta el relanzamiento. (#1627)
