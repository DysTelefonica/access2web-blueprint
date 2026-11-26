type: patch

El esquema del host contract declara `deploy_backup_publisher` como objeto opcional (`check_name` no vacío, `app_id` entero mayor o igual que 0, `todo` opcional, sin otras claves). `check_host_drift.py` valida esa forma y falla cerrado (exit 2) nombrando la clave si no se cumple; sin la clave, nada cambia, y una clave desconocida en la raíz sigue siendo exit 2. Así el contrato real de APAP_WEB deja de bloquear `host-readback`. Refs #1705.
