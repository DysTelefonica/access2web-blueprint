type: minor

**Qué.** El tipo de una PR sale de su título convencional (HR-93, #1726). El gate `pr-contract` mide el título: un título convencional válido basta para el merge sin etiqueta `type:*`; un título que no lo es bloquea con el motivo nombrado; y si hay etiqueta `type:*` y contradice el título, manda el título y se avisa sin bloquear. La checklist del tipo de HR-64 se elige por el título. La etiqueta `type:*` deja de ser requisito del contrato de merge y el documento operativo deja de pedirla al abrir el PR.

**Por qué.** Decisión del propietario del 2026-11-16 («que salga»): el tipo se declara en el título, no en una etiqueta. #1545 y #1528 quedan cerradas por eso.

**Dónde.** `assets/pr-contract/check_pr_contract.py` (`check_pr_title`, `pr_title_type`), `assets/pr-contract/tests/test_pr_title.py`, `assets/bin/ci-pattern` (etapa 6 del documento operativo), `SKILL.md` (HR-93).
