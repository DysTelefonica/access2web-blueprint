type: minor

El runner unitario ya no tiene lista literal de tests Python: cada suite vive en `testing/unit-suites.d/NNNN-<suite>.list` (#1664, PR-B). `test_unit_suite_order.py` comprueba la estructura: el orden de ejecución es la concatenación de los `.list` en orden C de nombre, con sus líneas en orden de fichero. Añadir una suite es añadir su fichero, sin tocar nada compartido ni una línea base. `test_every_test_runs.py` lee también esos ficheros.
