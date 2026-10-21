type: minor

El gate de paridad local/CI declara la capacidad que necesita cada comprobación (#917). Las exclusiones aceptan `requires` con el vocabulario cerrado `docker`, `runner`, `network` y `credentials`; la política declara `local.capabilities` con lo que tiene el entorno local. El gate falla si el entrypoint local ejecuta un comando que exige una capacidad no declarada, y declarar `requires` no exime de `side` ni de `reason`. El porting guide documenta cómo correr una comprobación sin Docker y el camino Windows/WSL, con su limitación conocida.
