type: patch

El propagador ya no omite en silencio los fixtures del catálogo que el `.gitignore` del consumer ignora. Cada fichero ignorado se nombra en el log y la corrida termina con error (exit 2) sin commitear esa skill, en vez de publicar una skill a medio materializar (#1222).
