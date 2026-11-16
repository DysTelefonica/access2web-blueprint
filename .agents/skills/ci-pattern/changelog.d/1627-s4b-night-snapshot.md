type: minor

Instantánea de la noche del perfil pre-release: `build_night_snapshot.py` reúne por `gh api` (solo lectura) los PRs fusionados desde la última noche buena con sus issues, los jobs caros de la noche y el deploy a producción del día, en el formato que lee `night_report.py`. Falla cerrado si una lectura falla. (#1627)
