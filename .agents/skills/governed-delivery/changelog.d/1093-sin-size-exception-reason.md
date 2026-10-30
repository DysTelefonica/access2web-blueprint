type: patch

La skill `governed-delivery` deja de mandar `size-exception-reason:` como vía de excepción al presupuesto de 400 líneas: la única es la etiqueta `size:exception`, que aplica el mantenedor, y el gate `check_pr_size.py` la lee de `--labels` o `$PR_LABELS` sin comprobar quién la puso (#1093).

La regla de aprobación se alinea con la decisión del dueño del 2026-10-23 (17:24, en los comentarios de #1093): `status:approved` no es puerta y la puede poner cualquiera, agentes incluidos; el merge sigue exigiendo la orden registrada de un mantenedor. La prohibición de la entrada `916-approval-path.md` queda sustituida por esta.
