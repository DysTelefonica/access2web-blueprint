type: patch

Añade la comprobación diaria de deploy del perfil pre-release (#1627, slice 7), con sus pruebas y modo de ensayo. `prerelease-deploy/check_daily_deploy.py` exige al menos un deploy exitoso a `production` el día UTC: un GitHub Deployment del entorno cuyo estado más reciente es `success`. Si no lo hay, abre una sola issue de máxima prioridad con la etiqueta fija `prerelease:night-red`. Falla cerrado si gh no responde y `--dry-run` no crea nada.
