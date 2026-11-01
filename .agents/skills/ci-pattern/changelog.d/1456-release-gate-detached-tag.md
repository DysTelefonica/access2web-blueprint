type: patch

El gate `check_release_tag.py` ya funciona en el checkout que deja un push de tag (#1456). `actions/checkout` con `fetch-depth: 0` deja el tag en HEAD separado y solo trae las ramas remotas, sin rama local, así que el gate salía con 2 (`default branch 'main' does not exist`) en todo push de tag y el preflight de `release.yml.tmpl` nunca pasaba. Ahora mide HR-20 contra la rama local cuando existe y, si no, contra `refs/remotes/origin/<rama>`. Sigue saliendo con 2 cuando la rama no existe en ninguno de los dos sitios.
