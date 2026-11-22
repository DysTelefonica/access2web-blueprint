type: minor

El merge gobernado (`check_governed_merge.py`) admite `--esperada <sha40>`: la cabeza que el
operador verificó antes de lanzar el merge (team-skills#1792). Un formato invalido falla
cerrado (2) antes de cualquier llamada de red; si la cabeza del PR no es la esperada, el gate
falla cerrado (2) en cuanto lee el PR, sin evaluar la política, sin leer la base y sin PUT; si
coinciden, el PUT lleva ese sha. Sin la opción, el comportamiento no cambia. Cierra por el
lado del cliente la ventana TOCTOU que la API ya cerraba con su 409.
