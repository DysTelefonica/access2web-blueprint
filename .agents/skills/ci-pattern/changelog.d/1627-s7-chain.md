type: patch

Añade el script que abre cada eslabón de una cadena de PRs del perfil pre-release (#1627, slice 7), con sus pruebas y modo de ensayo. `prerelease-chain/open_chain_link.py` pone `chain:partial` y `Refs #N` en cada eslabón intermedio, sin palabra de cierre; solo la punta lleva `Closes #N`; no deja abrir un cuarto eslabón abierto. Falla cerrado si gh no responde y `--dry-run` no crea nada.
