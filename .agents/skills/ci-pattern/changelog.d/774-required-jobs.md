type: patch

HR-3 (#774): `required-jobs/check_required_jobs.py` reparte `load_policy` (`validate_events`, `validate_relaxation`), `evaluate` (`policy_violations`, `classify_conclusions`) y la carga de la política base (`load_base_policy`), y usa `Path.samefile`; el veredicto, el orden de las violaciones, los mensajes y los códigos de salida no cambian. El baseline de ruff extendido baja en 5 hallazgos (C901 x2, PLR0912, PTH121 y TRY301; queda solo TRY003, fuera de alcance).
