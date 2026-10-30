type: patch

Se retira `assets/version-bump/` (`check_version_bump.py` y su suite). Ningún workflow ni plantilla lo ejecutaba, y HR-44 ya dice que el PR de trabajo no fija `metadata.version`: la versión la sube el PR de release (#665). Se quita también su línea del runner de suites unitarias y su entrada del baseline de ruff vendorizado. Refs #1093.
