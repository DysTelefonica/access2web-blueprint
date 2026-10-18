type: minor

El gate `pr-size` deja de leer el cuerpo del PR: la única excepción al presupuesto es la etiqueta `size:exception` que aplica el mantenedor, tomada de `--labels` o de `$PR_LABELS` (separadas por comas). El campo `size-exception-reason:` ya no existe para el gate, y los lockfiles (`uv.lock`, `package-lock.json`…) cuentan como cualquier otro fichero, como dicen `parameters.md` y el esquema desde A1 (#518). Un consumer que pasaba `PR_BODY` debe pasar `PR_LABELS` (#806).
