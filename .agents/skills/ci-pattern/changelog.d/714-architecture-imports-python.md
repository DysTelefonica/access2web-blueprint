type: minor

HR-68 (#714): `docs-architecture` impone las reglas de imports del bloque `architecture` en Python: deriva la configuración de import-linter, falla si la versionada diverge o el linter no se ejecuta y nombra `fichero:línea` y la regla de cada import prohibido. La declaración `architecture` exige ahora `stack` y `linter_config`.
