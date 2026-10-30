type: patch

HR-53 (tests de cierre) y HR-64 (evidencia de documentación) pasan a AVISO en el gate de contrato de PR (`assets/pr-contract/check_pr_contract.py`): el gate los imprime con «AVISO (no bloquea)» y sale 0; el PR ya no se bloquea por ellos. Decisión del dueño (2026-10-23, comentario en DysTelefonica/team-skills#1093): «que pasen a avisos, menos los commits convencionales» — los commits convencionales (HR-14) siguen bloqueando, y el flujo de PR bloquea por el enlace a la issue, el tamaño y los commits convencionales (#1093 R1/R2).

Lo que no cambia: la declaración de `docs_evidence` inválida sigue saliendo 2 (HR-3) y la comprobación activa sin su fichero de datos también; HR-53 y HR-64 pasan a `documented-only` (HR-34), con la regla enunciada en el propio `SKILL.md` y el aviso como dato para quien revisa.

El preflight local (`ci-pattern pr preflight`) y el modo local (`ci-pattern pr contract`) siguen al gate: imprimen el aviso y salen 0.
