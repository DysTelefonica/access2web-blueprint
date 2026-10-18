type: patch

HR-3 (#774): `runner-controls/safe_cleanup.py` reparte `cleanup` en pasos (`_bundle_step`, `_worktree_step`, `_delete_step`) y resuelve `git` con `shutil.which`; `check_runner_binding.py` reparte `main` en `unbound_findings`, `print_report` y `run`; `github_rest.py` agrupa cuerpo, cabeceras y método en `_Payload`. Comportamiento, mensajes y códigos de salida sin cambios. El baseline de ruff extendido baja de 386 a 375 hallazgos.
