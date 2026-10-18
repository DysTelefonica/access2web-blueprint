type: minor

Nuevo asset `assets/branch-name/check_branch_name.py`: gate del nombre de rama del PR con la regex como dato (`--pattern`, parámetro P01; por defecto `<type>/<slug>` sin número de issue). Exige el nombre completo, exime `main`, `archive/*`, `skill-fleet/*` y las ramas de Dependabot solo si `--actor` es `dependabot[bot]`, admite `--allow` para ramas anteriores a la convención y sale con 2 ante una regex ilegible o una rama vacía. Es el control compensatorio de los hosts sin rulesets (repos privados de plan Free), donde el ruleset que aplica la regex no existe (#806).
