type: patch

El texto generado del ciclo de vida (etapa 2, «Aprobación») y el de la sección de
issues dejan de presentar `status:approved` como una puerta: la etiqueta la puede poner
cualquiera, agentes incluidos (decisión del dueño del 2026-10-23, 17:24, en los comentarios
de #1093), ningún gate la exige y la PR no se bloquea por ella. La etapa 2 la presenta
como convención y la sección de pull requests declara qué bloquea de verdad: el enlace a
la issue, el presupuesto y los commits convencionales.

El docstring de `assets/bin/pr_preflight.py` deja de listar `status:approved` entre lo que
el preflight verifica: desde R1 ningún gate lee esa etiqueta ni la lista `approvers`.

Sin bump de `metadata.version`: lo consolida el PR de release (#665).
