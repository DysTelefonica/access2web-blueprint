type: minor

HR-63 (#727): `docs_map.code_roots` admite varias raíces de código (lista de `{root, glob}`) y `check_docs_codeguide.py` acepta pares `--code-root`/`--code-glob` repetibles; las claves escalares `code_root` y `code_glob` siguen valiendo para una raíz y mezclar las dos formas falla cerrado nombrando las claves.
