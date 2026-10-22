type: minor

Propagador (#787): al terminar de abrir la cadena vigente de un consumer, cierra con un comentario los PRs abiertos de cadenas anteriores del mismo namespace (`<bot_branch>-vendored*`) que no son de la cadena nueva, solo del mismo repositorio (nunca forks) y conservando la rama en el remoto. Un fallo al listar o cerrar avisa y no detiene la corrida. Antes, una cadena de un catálogo antiguo convivía con la nueva hasta que alguien la cerraba a mano (APAP_WEB: 129 PRs).
