type: patch

HR-63 (#757): `check_docs_codeguide.py` cumple el ratchet de ruff extendido de los consumers (PLR2004, PTH207, TRY003 y C901): constante con nombre, `Path.glob` en lugar de `glob.glob` (los ficheros ocultos también casan), excepciones propias que siguen siendo `ValueError` y `FileNotFoundError`, y la función principal dividida en sentido adelante y atrás sin cambiar el veredicto.
