type: minor

El perfil pre-release gana un informe nocturno en HTML (`assets/prerelease/night_report.py`, #1627 slice 4). Recoge los PRs desde la última noche buena, agrupados y enlazados a sus issues, el resultado de cada job caro y si hubo deploy. Tiene modo `--dry-run` que no escribe nada y sale con 2 ante una instantánea no válida.
