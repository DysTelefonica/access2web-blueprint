type: minor

`image-provenance` (#1389, 2/5): `ImageTree` lee el árbol desplegable y las etiquetas de un tarball de imagen (docker-archive y oci-archive), aplica las capas en orden con sus whiteouts y resuelve enlaces; `compare()` nombra los ficheros ausentes, extra y cambiados, y un cambio que solo afecta a los finales de línea sigue siendo hallazgo, nombrado como tal. Un archivo ilegible o sin ficheros en la ruta desplegable no es medible.

Tests: `assets/image-provenance/tests/test_image_tree.py` y `CompareTests` en `test_manifest.py`.
