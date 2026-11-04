type: patch

**Qué.** `safe_merge.py` ya no borra la rama remota después del merge e imprime `KEEP: remote branch '<rama>' is kept`. La rama local la decide `safe_cleanup` con la cabeza de la PR fusionada: si su punta no está preservada, se deja. La salida de `gh` se lee como UTF-8 con cualquier locale.
**Causa.** `cleanup_branch` borraba la rama remota (`DELETE git/refs/heads/<rama>`), en contra de la invariante de la flota, y la local con `git branch -D` sin comprobar nada (access2web-blueprint #1546, #1548, #1551, #1557). En Windows, un emoji del cuerpo de la PR rompía la decodificación cp1252 (#1491).
**Arreglo.** Se quita el borrado remoto, la limpieza local pasa a `safe_cleanup.cleanup(pr_head=...)` y `run_command` decodifica como UTF-8.
**A partir de ahora.** Las ramas fusionadas se quedan en `origin`.
