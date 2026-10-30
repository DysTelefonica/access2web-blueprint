type: patch

El gate `required` del asset `personal/ardelperal/ci-pattern/assets/required-jobs/check_required_jobs.py` admite la clave opcional `check_names` de la política: el nombre con el que el runner publica cada check, como dato declarado y no como código (#1219). Con la declaración, la causa de un check rojo nombra el check publicado y conserva el id del job (`lint (pull_request) [lint] (failure)`); sin ella la causa sale como antes. El README documenta los códigos de salida que el gate produce: 0, 1 y 2 (uso de argparse). No cambia qué jobs exige el gate ni la protección de rama.
