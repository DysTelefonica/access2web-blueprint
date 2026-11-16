type: minor

El perfil pre-release gana `assets/prerelease/provision-night-runner.sh` (#1627 slice 6): aprovisiona el directorio del runner nocturno de forma idempotente. Comprueba las versiones de Python y Node contra `night-runner.pins`, sale con 2 nombrando lo que falta o no cuadra, y su modo `--dry-run` no escribe nada.
