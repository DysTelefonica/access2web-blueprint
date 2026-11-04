type: patch

`governed-delivery` 1.5: la regla 2 dice cómo se comprueba el verde antes del merge: con `wait_ci.py` de `ci-pattern`, que espera a que los checks aparezcan y terminen, y nunca da verde con checks ausentes o pendientes (#1502).
