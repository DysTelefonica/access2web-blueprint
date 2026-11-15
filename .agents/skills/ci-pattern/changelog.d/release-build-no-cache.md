type: patch

`image-provenance`: `build_release_image.py` construye con `--pull --no-cache`. Antes reutilizaba la imagen base y las capas en caché de un build anterior, de modo que el `dnf update` del Dockerfile no se volvía a ejecutar y las correcciones de paquetes (CVE corregibles del escaneo de Quay) no llegaban a la imagen de release. `engine_cmd` no admite opciones después de `build`, por lo que el consumidor no podía corregirlo desde su configuración.

Tests: `assets/image-provenance/tests/test_build.py`.
