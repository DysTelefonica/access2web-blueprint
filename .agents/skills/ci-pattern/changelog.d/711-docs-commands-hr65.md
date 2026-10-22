type: minor

HR-65 (#711): asset `docs-commands`. Extrae los comandos documentados de `README.md`, `AGENTS.md` y `docs/**` (nombre de la CLI como parámetro del consumer), los clasifica en executed / parse-only / presence-only y falla si el parser real rechaza uno o si no extrae ninguno. Por declaración `docs_commands` en el contrato de adopción, la fase 5 valida la declaración y el job requerido. Sin declaración avisa por stderr y sale 0.
