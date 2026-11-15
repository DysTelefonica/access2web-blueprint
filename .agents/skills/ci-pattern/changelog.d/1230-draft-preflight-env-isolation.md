type: patch

El test del preflight en modo borrador aísla el entorno del proceso hijo y limita la trampa HTTP a un prefijo único por ejecución (#1230). El contrato API-free del draft no cambia: el test sigue rechazando cualquier llamada hecha con la URL de API vigente, pero ya no cuenta como llamada del preflight el tráfico ajeno que pueda alcanzar el puerto loopback efímero compartido.
