type: minor

Asset nuevo `assets/image-provenance/` (#1389, 1/5): `image_provenance.py manifest` escribe el manifiesto de release de un consumer que construye la imagen fuera del CI: el sha256 de cada fichero del árbol desplegable leído de los objetos git del commit del tag (`git ls-tree` + `git cat-file`), nunca de un checkout, sin normalizar finales de línea. Tag ausente, árbol vacío o configuración inválida salen con 2.

Tests: `assets/image-provenance/tests/test_manifest.py`.
