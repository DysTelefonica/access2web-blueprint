type: patch

**Qué.** El README de `image-provenance` documenta por qué un Dockerfile que construye podman usa nombres de imagen completos, y la decisión de no ampliar el gate HR-78 a los Dockerfiles.
**Causa.** Tras `podman system prune -af`, o con `--pull --no-cache` en una máquina nueva, el build de cadete falló en `COPY --from=composer:2.5`: podman no resuelve un nombre corto sin preguntar.
**Arreglo.** Cadete usa `docker.io/library/composer:2.5`. El README añade el inventario de nombres cortos de la flota a 2026-11-02 y deja escrita la decisión: el gate mide lo que descarga el CI, y los nombres cortos solo fallan en el build con podman, que solo hace cadete.
**A partir de ahora.** `FROM` y `COPY --from` llevan el registro en los Dockerfiles que se construyen con podman. Si otro consumer construye con podman, se revisa la decisión del gate. Fase 0 de CI/CD, punto 0.10.
