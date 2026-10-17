type: patch

La cabecera `Link` del stub de la API de GitHub (`_github_stub.py`) ya no lleva nada que envíe el cliente: toma el SHA de las claves guardadas en `state['pulls_pages']`, la dirección del propio servidor y una plantilla de ruta fija. Antes seguía copiando el path de la petición a la cabecera (CodeQL py/http-response-splitting). Prueba nueva con `Host` ajeno o plegado, path con `%0d%0a` y SHA falsificado, y comprobación de que la paginación sigue funcionando (#694).
