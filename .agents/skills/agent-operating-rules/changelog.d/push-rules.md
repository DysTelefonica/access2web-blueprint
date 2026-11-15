type: minor

`agent-operating-rules` 1.1: HR-13 con las normas de push que el dueño fijó para los workers de APAP_WEB el 2026-11-01. Un push por tanda, no poner al día una rama hasta que se vaya a fusionar, empujar solo el eslabón que sigue de una cadena, comprobar en local antes de empujar y cancelar solo las ejecuciones de la propia rama que ya no sirven. Si la cola sigue llena, se piden más runners a infra-coord.
