type: patch

`required_check_names` (`advisory_checks.py`) falla cerrado ante un `required_checks` mal formado, con la misma regla que `check_push_compliance.required_checks()`: lista, y cada entrada un objeto con `name` cadena no vacía. Antes descartaba en silencio las entradas rotas, y con ello un requerido podía quedar sin protección frente a un patrón advisory. `safe_merge.py` y `wait_ci.py` salen con exit 2 nombrando la entrada rota, sin dar verde ni fusionar. Sin la clave `required_checks` no hay requeridos, como antes (team-skills#1795).
