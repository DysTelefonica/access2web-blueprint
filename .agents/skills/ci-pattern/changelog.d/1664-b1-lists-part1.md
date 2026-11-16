type: patch

Parte 1 de la división de #1664 (B1): las suites de 0100 a 0660 pasan de la lista literal de `testing/run-unit-suites.sh` a `testing/unit-suites.d/`. El orden de ejecución no cambia: el prefijo que queda en la lista literal va antes que los ficheros. `test-ci-workflow-pins` comprueba la suite del gate de contrato en la lista literal o en `unit-suites.d`, y que el entrypoint lee `unit-suites.d`.
