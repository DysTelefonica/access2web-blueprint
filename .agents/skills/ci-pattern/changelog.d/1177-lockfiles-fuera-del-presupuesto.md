type: patch

El gate de presupuesto de revisión (`assets/pr-size/check_pr_size.py`) excluye los lockfiles generados: `uv.lock`, `package-lock.json`, `pnpm-lock.yaml`, `poetry.lock` y `requirements*.lock` (#1177).

El techo sigue siendo 400 líneas y la etiqueta `size:exception` sigue siendo la única excepción: lo que cambia es qué se cuenta. El lockfile lo escribe el gestor de paquetes a partir del manifiesto, así que sus líneas no son trabajo de quien abre la PR, y un bump que mueve el árbol de dependencias se comía el presupuesto entero sin que nadie pueda partirlo: en `ardelperal/APAP_WEB#2424` (bump de mypy 2.3.1 → 2.4.0) 444 de las 448 líneas del diff eran `uv.lock` y las 4 restantes, `pyproject.toml`. Ese caso concreto ahora mide 4 y pasa sin etiqueta.

La regla vive en la fuente única del presupuesto (`assets/budget/vendored_budget.py`), así que los dos gates que la comparten miden lo mismo: `pr-size` y el veredicto A1 de `pr-contract`.

- Se reconocen por **nombre de fichero**, en cualquier directorio del repo.
- `requirements.txt` (y `requirements.in`) **sí** cuentan: solo la extensión `.lock` queda fuera.
- Un renombrado **desde una ruta de autoría hacia** un lockfile **sí** cuenta: el trabajo fue de autoría. Entre lockfiles no cuenta.
- Ningún otro fichero generado queda fuera: un `.min.js` o un `.pb.go` siguen contando.
- `pr-size` reporta las líneas excluidas **aparte** (`generated_lines` en el sobre y «N generated line(s) excluded» en el texto), para no llamarlas «vendored».

Tests: la suite de `pr-size` (modo `git` y modo API) y la de `budget` cubren el patrón, los límites (`requirements.txt` cuenta, `requirements-dev.lock` no), el renombrado desde autoría, el caso real de 444 + 4 → 4, y que 500 líneas de código normal sigan en rojo.
