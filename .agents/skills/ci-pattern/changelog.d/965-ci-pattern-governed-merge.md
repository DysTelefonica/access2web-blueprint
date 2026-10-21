type: patch

El texto generado del merge en el documento operativo (paso 9 del flujo y sección «Merge») decía que ningún agente mergea. Con la regla 2 de `governed-delivery` (#962), un agente puede lanzar el merge gobernado solo por orden explícita y registrada de un mantenedor con cuenta de mantenedor, por el camino gobernado y con CI en verde, nunca con `--admin`. Ambos textos del CLI y el test que los fija ahora dicen esa regla. Ningún gate ni método de merge cambia. Los consumers verán el cambio de su documento operativo generado en la siguiente propagación.
