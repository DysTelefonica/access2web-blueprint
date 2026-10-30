type: minor

`image-provenance` (#1389, 4/5): `build_release_image.py` construye la imagen de release solo desde un clon limpio del tag: rechaza tag ausente, tag fuera de la rama por defecto y checkout sucio o distinto del contenido git del tag, y etiqueta `org.opencontainers.image.revision` con el commit del tag y `org.opencontainers.image.version` con el tag. `core.autocrlf=false` va en cada invocación de git, nunca en una configuración.

Tests: `assets/image-provenance/tests/test_build.py`.
