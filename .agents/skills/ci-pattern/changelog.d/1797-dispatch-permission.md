type: minor

El texto generado del merge en el documento operativo (paso 9 del flujo y sección «Merge») deja de pedir una cuenta de mantenedor: el agente despacha el merge gobernado con una cuenta con permiso de escritura en el repositorio (`actions:write`), que es lo que exige `gh workflow run`; el merge lo ejecuta el `GITHUB_TOKEN` del job con sus gates (team-skills#1797). Coherente con la regla 2 de `governed-delivery`. La variante sin roles no cambia.
