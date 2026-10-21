type: patch

#933: el dispatch post-merge de `check_governed_merge.py` se lanza sobre la rama base del PR (`ref` solo admite una rama o una etiqueta; con un SHA la API responde 422) y pasa en `inputs` el rango auditado —`before` (punta de la base antes del merge) y `after` (SHA fusionado)— a los workflows que los declaran bajo `workflow_dispatch`; `check_push_compliance.py` audita ese rango vía compare y falla cerrado si `before`/`after` faltan o no son SHA. Los workflows sin `before`/`after` siguen recibiendo solo `ref`.
