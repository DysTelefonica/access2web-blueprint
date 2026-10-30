type: patch

HR-64 (#1013): el gate y el generador de datos del rango calculan los ficheros y el numstat de la PR con `BASE...HEAD` (merge-base), no con `BASE..HEAD`. Una PR apilada que se reapunta a `main` tras mergearse su base dejaba de juzgar solo sus cambios y heredaba los ficheros que `main` ya había mergeado, con el falso positivo consiguiente. La ayuda de `--changed-files` y `--numstat-file` documenta el rango nuevo; los commits siguen con `BASE..HEAD`.
