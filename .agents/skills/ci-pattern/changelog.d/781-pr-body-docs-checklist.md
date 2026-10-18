type: patch

Propagador (#781): el cuerpo de los PRs de la flota enlaza la checklist del tipo de cambio (`type:chore`) bajo «## Documentación» cuando el consumer exige HR-64: la lee de `docs_evidence.checklists` de la política de su propio `check_pr_contract` (`.github/pr-contract.policy.json` o `.github/pr-contract/*.json`), como ya hace con `size_exception_field`. Sin esa política el cuerpo no cambia. Antes, toda propagación a un consumer con HR-64 nacía con `pr-gate` en rojo.
