type: minor

Skill nueva `release-notes`: redacta las notas de una versión desde el rango entre dos tags, agrupadas en Seguridad, Errores de producto, Funcionalidad nueva y Repositorio, pruebas y CI, con cada entrada en antes, ahora y dónde. Perfiles web/contenedor (con la plantilla de `ci-pattern`) y VBA/Access. `assets/release_coverage.py` recoge el rango (`collect`) y falla nombrando cada PR, commit sin PR o issue que las notas no cuentan (`check`). Refs #1425.
