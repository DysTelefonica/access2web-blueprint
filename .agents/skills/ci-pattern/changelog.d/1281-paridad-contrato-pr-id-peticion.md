type: patch

El test de paridad del contrato de PR local (`assets/bin/tests/test_pr_contract_local.py`) compara
el `stderr` de **dos** ejecuciones independientes del gate: la local, en proceso, y el gate real,
como proceso aparte. Desde que el cliente REST reintenta las lecturas JSON (#1220), cada fallo de
llamada imprime por `stderr` un aviso con el `X-GitHub-Request-Id` de la respuesta, un
identificador que GitHub cambia en cada petición: las dos ejecuciones no pueden coincidir byte a
byte y el test salía rojo en `main` en toda máquina con un token de GitHub en el entorno y la API
contestando con error (401 por credenciales, 403 por cuota, 5xx). En la CI no se veía: el paso que
corre la suite unitaria no lleva token, así que el chequeo HR-7 ni siquiera llama a la API.

Ahora el test normaliza **solo** el valor del identificador antes de comparar (#1281). El resto
del texto —ruta, número de intento, estado HTTP, mensaje de GitHub y número de líneas de aviso— se
sigue comparando byte a byte, y no cambian ni el `stdout` ni el código de salida ni la exigencia
del aviso HR-7. El identificador sigue apareciendo en el aviso del reintento: es útil para
diagnosticar, así que el cliente REST no se toca. El arreglo vive en la suite, no en el producto:
un test que falla en `main` tapa los rojos reales.

Tests: `assets/bin/tests/test_pr_contract_local.py` — el caso de paridad, más tres casos nuevos
que fijan la normalización: dos identificadores distintos se normalizan igual, el estado HTTP y el
número de intento siguen mandando, y un texto sin identificador no cambia.
