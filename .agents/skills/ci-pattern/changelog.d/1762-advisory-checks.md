type: minor

**Qué.** El consumer puede declarar en su `host-contract.json` la clave opcional
`advisory_checks`: una lista de patrones de nombre (globs) de los check runs que no bloquean
el merge. `safe_merge.py`, `wait_ci.py` y el merge gobernado dejan de juzgarlos: no cuentan
para verde/rojo ni para pendiente, y se listan en una línea informativa. Un check
**requerido** (por `--required-check`, por `--check` o por el contrato) nunca queda tapado
por un patrón. Sin la clave, el comportamiento es idéntico al de antes (#1762, APAP_WEB#4411
y #4423).
**Causa.** Esos guards leen *todos* los check runs del sha, así que un job espejo en runners
de GitHub (sufijo `-gh`) en rojo, o con `startup_failure` cuando se agota el crédito, paraba
los merges de los agentes aunque `required` estuviera en verde.
**Arreglo.** El lector común (`assets/runner-controls/advisory_checks.py`) trae la declaración
del contrato y separa los runs que se juzgan de los que solo se listan; una declaración con
forma inválida (o un contrato ilegible) falla cerrado (exit 2) en vez de ignorar todo. La
clave entra en `host-contract.schema.json` y en `host-contract.example.json`, y el validador
del host-contract falla cerrado si su forma no es válida.
**A partir de ahora.** Un consumer con jobs espejo declara `advisory_checks` con sus patrones
(`["*-gh"]`) y sus merges dejan de depender de ellos; sin la clave, nada cambia.
