type: minor

El runner unitario acepta tests Python por suite en `testing/unit-suites.d/NNNN-<suite>.list` (#1664, PR-A). `run-unit-suites.sh` los expande en orden C después de su lista literal, y el gate HR-4 (`local-ci-parity`) los lee igual. La lista literal sigue en uso y el directorio solo tiene `0990-unit-suite-order.list`, así que el conjunto y el orden de tests no cambian. El test `test_unit_suite_order.py` lo fija contra la línea base de 145 entradas de #1669.

Fixes de PR-A de #1664: un patrón sin coincidencias ya no se pierde en silencio, los `.list` con CRLF se leen igual que con LF, y una línea comentada no cablea su test. Sin directorio o sin listas, el run falla.
