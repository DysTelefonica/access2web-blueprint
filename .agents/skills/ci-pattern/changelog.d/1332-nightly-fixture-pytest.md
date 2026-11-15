type: patch

El fixture de los tres autotests del runner (`test-bash-parallel.sh`, `test-rc-aggregation.sh` y `test-suite-timing.sh`) deja de sembrar los stubs `.py` con un `sys.exit(0)`: un stub así no lo puede recoger pytest, que aborta la recogida con `SystemExit` y deja el repositorio falso en rojo aunque todas sus suites estén en PASS. El stub pasa a ser un módulo importable, que sale con 0 como script y que pytest recoge como un test verde (#1332).
