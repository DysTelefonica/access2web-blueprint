type: minor

`pr-contract` admite `"advise"` en cada entrada de `checks` (team-skills#1786): el control corre y sus hallazgos salen como `AVISO (advise, no bloquea)` por `stderr` sin cambiar el exit code. Cubre `closes`, `commits`, `issue_links`, `budget` y el resto de comprobaciones de la clave; `true`, `false` y ausente conservan su significado y un valor distinto sale 2. `check_e2e_evidence.py` gana `--advise` con la misma semántica para la evidencia e2e de UI (HR-58). Sin consumers que declaren `advise`, el comportamiento no cambia.
