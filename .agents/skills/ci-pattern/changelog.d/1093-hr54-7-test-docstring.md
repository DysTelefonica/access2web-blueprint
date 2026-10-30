type: patch

Docstring de `test_pr_contract.py` (#1093): la clase `IssueLinkGateTests` decía que la mezcla `Closes`/`Refs` a la misma issue «es aviso, #787». No lo es desde `47f74411` (Refs #840, 2026-10-18): `issue-spec` no produce ni fallo ni aviso. Solo texto; el comportamiento no cambia, y lo fijan `test_closes_and_refs_for_the_same_issue_passes` e `test_closes_and_refs_same_issue_passes_without_warning`.
