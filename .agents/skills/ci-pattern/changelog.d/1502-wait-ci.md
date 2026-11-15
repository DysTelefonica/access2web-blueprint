type: minor

**Qué.** Asset nuevo `assets/runner-controls/wait_ci.py`. Espera a que la CI de una PR aparezca y termine. Sale con 0 en verde, con 1 en rojo, sin checks o fuera de tiempo, y con 2 en uso o error de API.
**Causa.** `gh pr checks --watch` sale con 1 y `no checks reported` justo después del push, y `--required` dice `no required checks reported` mientras el agregador no se lanza. Cada agente escribía su propio bucle, y algunos daban la CI por verde con checks que todavía no habían aparecido (#1502). También pasa cuando ya hay un run verde y la suite de `tests` está aún en cola (team-skills #1500).
**Arreglo.** Lee los check runs y las check suites por la API (sin `gh`, igual en Windows y en Linux). Espera a los `--check` nombrados y a las suites de GitHub Actions que siguen arrancando dentro de `--appear-seconds`.
