type: patch

`check_tree_integrity` leía el árbol de un commit en `tree.sha` de la raíz, campo que la API de GitHub no devuelve: el árbol vive en `commit.tree.sha`. Como el campo no existía, el chequeo nunca emitía veredicto. Ahora lee la forma real de la API y solo compara el commit de merge (`merge_commit_sha`) cuando la base del merge es ancestro del head (`compare` en `ahead` o `identical`). Los commits intermedios de la PR y los merges con main avanzado ya no abren incidente. Limitación conocida: en un rebase-merge el chequeo no aplica y un drift no se detecta (#928).
