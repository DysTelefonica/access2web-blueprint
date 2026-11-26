type: minor

`image-provenance` (#1389, 3/5): `image_provenance.py gate`, antes de `oc set image`, rechaza una imagen cuya revisión no es el commit de su tag de release, cuyo commit nunca se subió o cuyo tag no está en la rama por defecto (medido sobre un clon bare fresco del remoto), y una imagen cuyo árbol difiere del manifiesto de la release; imprime además el diff frente a la imagen en ejecución. Sale 0 correcto, 1 hallazgos con los ficheros nombrados y un resumen por tipo, 2 no medible.

Tests: `assets/image-provenance/tests/test_gate.py`.
