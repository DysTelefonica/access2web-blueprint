type: minor

**Qué.** El consumer puede declarar en su `host-contract.json` la clave opcional
`advisory_checks`: una lista de patrones de nombre (globs) de los check runs que no
bloquean el merge. La clave entra en `host-contract.schema.json` y en
`host-contract.example.json`, y el validador del host-contract falla cerrado (exit 2) si
su forma no es valida. El lector comun vive en
`assets/runner-controls/advisory_checks.py` (#1762).
