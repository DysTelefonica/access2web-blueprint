type: minor

Nuevo control local `ci-pattern checkout-sync` (asset `assets/checkout-sync/`) que compara el checkout con `origin/<base>` sin red y distingue sincronizar la rama base de integrar `main` en una rama de trabajo; sale 2 si no puede medir. Se ejecuta en el arranque, justo después de `git pull --ff-only` (invariante del partial `web`). No es un gate de CI ni bloquea pushes (#914).
