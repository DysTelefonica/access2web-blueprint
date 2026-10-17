type: patch

Los 33 gates vendorizables que imprimían sin fijar la codificación ahora llaman a `_pin_output_encoding()` al inicio de `main()` (UTF-8 en stdout y stderr), con el patrón de issue-spec. La lista de deuda `UNPINNED_GRANDFATHERED` queda vacía y el guard exige el pin a todo gate nuevo (#663).
