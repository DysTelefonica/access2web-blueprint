type: patch

`release.yml.tmpl`: el preflight trae el tag anotado del remoto con `git fetch --force` antes del gate del tag. En un push de tag, `actions/checkout` crea `refs/tags/<tag>` como tag ligero sobre el commit, y `fetch-tags: true` no lo sobrescribe, así que `check_release_tag.py` fallaba con HR-19 «lightweight» sobre un tag anotado (primer tag de cadete con `release.yml`, v1.3.2). El token llega a git por variables de entorno de configuración, nunca en la línea de comandos ni persistido. Un push de tag usa el workflow del commit etiquetado: el arreglo rige desde el siguiente tag.

Tests: `assets/templates/tests/test_release_workflow_template.py`.
