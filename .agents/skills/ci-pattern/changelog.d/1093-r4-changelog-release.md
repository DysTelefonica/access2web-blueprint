type: patch

HR-44 deja de bloquear cada PR de trabajo (#1093, R4). El paso del gate `changelog-fragment` sale de `tests.yml` y de la política de paridad local/CI. El fragmento sigue siendo obligatorio por convención en cada edición material, y el consolidador `consolidate_changelog` valida cada fragmento pendiente en el PR de release, con fallo cerrado si uno es inválido. El script del gate se conserva para ejecutarlo a mano.
