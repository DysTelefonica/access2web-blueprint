type: patch

**Qué.** El merge gobernado (`check_governed_merge.py`) ya no cierra issues cuando el PR lleva `chain:partial`, e imprime `NOT CLOSED: PR #N is a chain intermediate (chain:partial): only the chain tip closes its issues`.
**Causa.** El gate cerraba todo lo que devolvía `closingIssuesReferences`, y GitHub puede conservar ahí un «Closes #N» que ya se borró del cuerpo. En cadete, el intermedio #1065 cerró #826 (#1486).
**Arreglo.** Si el PR lleva `chain:partial`, el gate no consulta `closingIssuesReferences` ni cierra nada. La punta de la cadena sigue cerrando como antes.
