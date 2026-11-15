type: patch

El propagador detecta el retarget del tramo 2 a la rama activa, publica un commit normal que conserva el árbol para activar el evento `synchronize` y verifica los `required_checks` del contrato del consumer sobre esa cabeza. Si un check falta o sigue pendiente, la cadena termina con error y no se declara lista (#1238).
