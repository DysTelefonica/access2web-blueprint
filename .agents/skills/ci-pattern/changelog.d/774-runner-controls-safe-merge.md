type: patch

HR-3 (#774): `runner-controls/safe_merge.py` reparte `main` en funciones por paso (`load_pull`, `base_findings`, `ci_findings`, `chain_findings`, `run`) y lanza todos los procesos por un único `run_command` que resuelve el ejecutable con `shutil.which`; el comportamiento, los mensajes y los códigos de salida no cambian. El baseline de ruff extendido baja de 402 a 386 hallazgos.
