type: minor

El contrato de adopción admite la clave opcional `docs_where_to_work`, cuyo `mode` (`exigido` o `aviso`) decide si la fase 5 falla o solo avisa cuando la página `docs/codebase/where-to-work.md` no cuadra. La clave es opcional a propósito: sin ella el gate no se mide y ningún consumer cambia de comportamiento por propagar; el cambio de modo de cada repo va en su propio PR de adopción. El modo es dato del contrato, no cableado del consumer (#1201).
