type: patch

El gate de ensayo nocturno en seco (`assets/prerelease/night_gates.py`) reconoce los scripts nocturnos también donde los vendoriza un consumer: `NIGHT_PATHS` usa `**/ci-pattern/assets/prerelease/night*.py` y `.sh` en lugar de la ruta del catálogo, así que un PR que solo cambia esos scripts en `.agents/skills/ci-pattern/` ya no pasa sin ensayo. Refs #1690.
