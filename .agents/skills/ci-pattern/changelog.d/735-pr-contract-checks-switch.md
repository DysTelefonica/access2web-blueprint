type: minor

#735: `check_pr_contract.py` acepta la clave opcional `checks` en la política (`{comprobación: true|false}` sobre `labels`, `closes`, `chain_context`, `issue_links`, `closure_tests`, `docs_evidence`, `budget`, `commits`) para adoptar el contrato de forma incremental, por ejemplo solo HR-64. Sin `checks` el comportamiento no cambia; un nombre desconocido, un valor no booleano o las ocho apagadas salen 2 (HR-3) y las comprobaciones apagadas se listan por `stderr`.
