type: patch

Las 28 clases de test de los assets que vulture daba por muertas se renombran a `Test*` (heredan de `unittest.TestCase`, así que ya se ejecutaban; mismo número de pruebas por fichero) y salen de `DEAD_GRANDFATHERED`, que baja de 39 a 11 entradas (#663).
