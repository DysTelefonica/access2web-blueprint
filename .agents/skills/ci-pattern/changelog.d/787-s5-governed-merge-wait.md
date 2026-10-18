type: minor

Merge gobernado (#787): un check requerido que aún no existe o no ha terminado (`missing`/`pending`) se espera hasta `--wait-seconds` (300 por defecto, cada `--poll-seconds`, 10 por defecto) en lugar de bloquear el merge lanzado antes de que el CI publicara `required`. Un check en fallo no se espera, y un valor negativo falla cerrado (exit 2).
