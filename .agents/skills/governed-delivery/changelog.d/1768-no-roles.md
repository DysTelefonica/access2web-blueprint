type: minor

Sin `agent_roles: enabled` en el `.team-skills.yaml` del consumidor, la skill que llega no tiene roles: el merge gobernado lo lanza el agente que lleva la issue con el perfil del administrador que lleva la sesión y la CI en verde (una orden permanente del `AGENTS.md` local cuenta como orden registrada), y `size:exception` la decide ese administrador. Con la clave, el texto es el de siempre (team-skills#1768).
