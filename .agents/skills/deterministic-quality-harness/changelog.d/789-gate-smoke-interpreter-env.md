type: patch

Versión 2.2 (HR-44). `test_gate_smoke.py` ya no lanza `check_pr_size.py` con un entorno reducido a `PATH` y `HOME`: conserva `LD_LIBRARY_PATH`, `DYLD_LIBRARY_PATH` y `SYSTEMROOT`, que el propio intérprete necesita para arrancar (#789). Con un Python de `libpython` compartida (el de `actions/setup-python` en CI) el hijo moría con exit 127 antes de ejecutar el gate y 5 tests de `pr-size` fallaban sin tener nada que ver con el gate. El resto del entorno de CI sigue sin entrar, a propósito. Los 43 assets sin exención se re-estampan a `v2.2` por la regla del marcador.
