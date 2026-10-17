type: minor

HR-67 (#754): `check_docs_decisions.py` acepta una lista de slugs en `sustituida-por` (cada destino debe existir y la detección de ciclos los recorre todos) y deja de tomar por cita una palabra suelta tras «decisión:»; solo cuentan `decisión: AAAA-MM-DD-…`, la forma entre comillas invertidas y el texto de un enlace.
