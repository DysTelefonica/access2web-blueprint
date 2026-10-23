type: patch

R5b de #1093: se retira `assets/hr-gate-matrix/`, el directorio que aún contenía el gate HR-1 (`check_structured_inputs.py`) y su suite. El gate no estaba cableado en ningún workflow, y la regla HR-1 no citaba ningún gate en el SKILL, así que su enforcement ya era documental. Se quitan también su línea de `testing/run-unit-suites.sh` y su entrada en `testing/vendorable-ruff-extended-baseline.json`, y la mención del directorio en `SKILL.md`.
