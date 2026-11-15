type: patch

El gate de presupuesto de revisión (`assets/pr-size/check_pr_size.py`) lee ahora las etiquetas del PR por API (`GET issues/N/labels`) en el modo `--event`, en vez de tomarlas del payload del evento (#1359).

El payload es una foto del instante en que se dispara el evento. Si el mantenedor pone `size:exception` justo después del `opened`, el run de ese evento no la ve y falla aunque la etiqueta ya esté puesta; un relanzamiento del mismo evento vuelve a leer el mismo payload y falla igual. Caso real: `DysTelefonica/cadete#715`, donde la etiqueta llegó 7 s después de abrir el PR y el `pr-size` del `opened` falló mientras el `pr-contract` del mismo run (que ya leía en vivo) pasó.

- Si la API no devuelve las etiquetas, el gate falla cerrado (exit 1), igual que con el recuento de ficheros.
- Una etiqueta del payload que ya no está en el PR no exime.
- El techo sigue siendo 400 líneas y `size:exception` sigue siendo la única excepción, aplicada solo por el mantenedor.
- Las ejecuciones locales (`--labels`, `$PR_LABELS`) no cambian.
