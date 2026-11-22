type: patch

**Qué.** El lector de `advisory_checks` falla cerrado ante un contrato ilegible: si el
host-contract existe pero no se puede leer (no es UTF-8) o su declaracion no tiene una
forma valida, el guard sale con 2 y nombra el motivo, en vez de ignorar la declaracion y
dejar el merge en manos de un check que no bloquea. El mismo lector devuelve ahora los
checks que el contrato declara como requeridos, para que un check requerido nunca
quede tapado por un patron advisory (#1762).
