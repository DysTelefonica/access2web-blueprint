type: patch

Detección de bloques cercados compartida (#712): `assets/markdown_fences.py` (CommonMark: tres o más comillas invertidas o tildes, cierre con el mismo carácter y al menos la misma longitud, bloque sin cerrar hasta el final) sustituye al recorte propio de `docs-codeguide`; los demás gates de documentación la importan en vez de copiarla.
