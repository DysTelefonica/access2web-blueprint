type: minor

Skill nueva `infra-disk-alerts`: atiende las alertas de disco que los hosts envían por Telegram. Reconoce el mensaje de `thinpool-guard` (thin pool de `pve`, `pve2` y `pve3`) y el de `vps-disk-guard.sh` (disco raíz de `rosa` y `oracle`), diagnostica en solo lectura en el host que lo envía y propone la acción, que se confirma antes de escribir. Nunca borra la instantánea `clean` de un runner efímero ni lanza `prune` de volúmenes. Versiona `assets/thinpool-guard.sh` (Telegram con Gmail de respaldo) con su test y las instrucciones de reinstalación. Refs #1430.
