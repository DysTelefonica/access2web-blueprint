type: minor

Merge gobernado (#787): deja de re-evaluar el tamaño del PR. Lo mide `pr-size` (que excluye `VENDORED_ROOTS` y exige la etiqueta `size:exception`) y el check `required` ya lo incluye; repetirlo aquí bloqueó la propagación vendorizada de APAP_WEB #2122 (19.786 líneas) que `pr-size` daba por buena. `--max-lines` se sigue aceptando, pero se ignora, para no romper los workflows que lo pasan.
