type: patch

La punta de una cadena nueva lista en `Closes` cada issue que sus eslabones referencian (#1190, decisión del dueño, 2026-10-24).

Una cadena nueva declara su id en el Chain Context de la punta (`- chain: <id>`). El gate de contrato de PR lee los PR del mismo id por API y avisa de los `Refs #N` de los eslabones que la punta no lleva en `Closes`. Es un aviso (HR-7, #1172): no cambia el código de salida, y el autor sigue escribiendo la lista. Las cadenas sin `chain:` en la punta no se miden, así que las ya abiertas no reciben avisos nuevos. Si la API no responde, el gate avisa y sigue; sin token el chequeo se omite.

Tests: `assets/pr-contract/tests/test_chain_closes.py` (ocho casos: la lista que falta, la lista completa, cadena sin id, otra cadena, `Closes` en un eslabón, fallo de API, ausencia de token y exclusión de la propia punta). La suite existente del gate se ajusta solo en el mock de API, que ahora sirve el listado de PR vacío.
