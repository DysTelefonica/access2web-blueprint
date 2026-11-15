type: patch

El fixture de `testing/suites/run-unit-suites/test-bash-parallel.sh` deja de depender de que `actionlint` esté instalado en el `PATH` del host: el repositorio falso incluye su propio stub y lo antepone al `PATH` del runner real (#1236). Así, el job `full` puede ejecutar los autotests del runner antes del paso que instala la herramienta, sin convertir su ausencia en un falso `FAIL actionlint`.
