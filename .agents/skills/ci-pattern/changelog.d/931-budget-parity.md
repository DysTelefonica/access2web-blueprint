type: patch

`pr-size` y `pr-contract` (A1) cuentan el mismo diff: la regla de exclusión vive ahora en `personal/ardelperal/ci-pattern/assets/budget/vendored_budget.py` y ambos gates la importan. `pr-contract` acepta un `--origins-file` opcional (que genera el runner local, porque el módulo compartido no ejecuta `git`) que restaura el origen authored de un fichero vendorizado que llega en otro commit del PR, así que un renombrado desde autoría sigue contando. Sin ese fichero el presupuesto de `pr-contract` no cambia (#931, #623, #900).
