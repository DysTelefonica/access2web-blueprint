type: patch

`check_release_ci.py` separa sus dos sujetos (#1423, verificador independiente). Con `--repo OWNER/NAME`, ese valor va a la API y el checkout local sale de `--git-dir`, que por defecto es el directorio actual. Con `--repo <directorio>`, la API toma `GITHUB_REPOSITORY`. Antes, un mismo valor tenía que servir de ruta git y de OWNER/NAME, y la llamada salía siempre con 2 fuera de Actions. La nueva opción `--tag-run` mide solo el run lanzado por el push del tag: otro run sobre el mismo commit (un push a main o un dispatch) puede haber saltado las suites.
