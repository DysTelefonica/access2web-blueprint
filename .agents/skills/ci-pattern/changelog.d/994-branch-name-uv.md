type: patch

El gate `assets/branch-name/check_branch_name.py` acepta también las ramas `dependabot/uv/...` del bot (solo con `--actor dependabot[bot]`), y el valor por defecto de P04 en `parameters.schema.json` lista el mismo ecosistema. Antes, un consumer con dependencias `uv` veía rechazadas sus PR de Dependabot (#994).
